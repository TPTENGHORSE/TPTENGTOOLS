import contextlib
import io
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import load_workbook

import Dataimport
from Dataimport import INVOICE_COLUMNS, deduplicate_invoice_rows


def invoice_row(source, description="Part", amount=10.0, invoice="INV-1"):
    row = {column: "" for column in INVOICE_COLUMNS}
    row.update(
        {
            "Invoice #": invoice,
            "Reference & Description": description,
            "Amount": amount,
            "PDF Source": source,
        }
    )
    return row


class DeduplicateInvoiceRowsTests(unittest.TestCase):
    def test_duplicates_from_different_pdfs_are_omitted(self):
        first = invoice_row("a.pdf")
        repeated = invoice_row("b.pdf")
        different_amount = invoice_row("b.pdf", amount=20.0)
        different_invoice = invoice_row("b.pdf", invoice="INV-2")

        rows, omitted = deduplicate_invoice_rows(
            [first, repeated, different_amount, different_invoice]
        )

        self.assertEqual(rows, [first, different_amount, different_invoice])
        self.assertEqual(omitted, {"b.pdf": 1})

    def test_identical_lines_in_the_same_pdf_are_preserved(self):
        first = invoice_row("a.pdf")
        second = invoice_row("a.pdf")
        copies = [invoice_row("b.pdf"), invoice_row("b.pdf")]

        rows, omitted = deduplicate_invoice_rows([first, second, *copies])

        self.assertEqual(rows, [first, second])
        self.assertEqual(omitted, {"b.pdf": 2})

    def test_more_occurrences_in_another_pdf_are_preserved(self):
        first = invoice_row("a.pdf")
        copies = [invoice_row("b.pdf"), invoice_row("b.pdf")]

        rows, omitted = deduplicate_invoice_rows([first, *copies])

        self.assertEqual(rows, [first, copies[1]])
        self.assertEqual(omitted, {"b.pdf": 1})

    def test_interleaved_sources_and_empty_input(self):
        first = invoice_row("a.pdf")
        copy = invoice_row("b.pdf")
        second = invoice_row("a.pdf")
        another_copy = invoice_row("b.pdf")

        rows, omitted = deduplicate_invoice_rows([first, copy, second, another_copy])

        self.assertEqual(rows, [first, second])
        self.assertEqual(omitted, {"b.pdf": 2})
        self.assertEqual(deduplicate_invoice_rows([]), ([], {}))

    def test_main_writes_deduplicated_rows_and_logs_omissions(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory)
            output = source / "output.xlsx"
            pdfs = [source / "a.pdf", source / "b.pdf"]

            def fake_process(pdf_path, source_dir, invoice_rows, *args):
                invoice_rows.append(invoice_row(pdf_path.name))
                return {
                    "PDF Source": pdf_path.name,
                    "Status": "Extraído",
                    "Details": "1 línea(s) de factura.",
                }

            with (
                patch("sys.argv", ["Dataimport.py", "--source", str(source), "--output", str(output)]),
                patch.object(Dataimport, "iter_pdf_paths", return_value=pdfs),
                patch.object(Dataimport, "process_pdf", side_effect=fake_process),
                contextlib.redirect_stdout(io.StringIO()) as stdout,
            ):
                Dataimport.main()

            workbook = load_workbook(output, read_only=True, data_only=True)
            try:
                self.assertEqual(workbook["Invoice lines"].max_row, 2)
                logs = list(workbook["Processing log"].iter_rows(values_only=True))
                self.assertIn("1 línea(s) omitida(s)", logs[2][2])
                self.assertIn("Líneas duplicadas omitidas: 1", stdout.getvalue())
            finally:
                workbook.close()


if __name__ == "__main__":
    unittest.main()
