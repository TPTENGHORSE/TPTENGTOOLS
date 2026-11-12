import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import load_workbook

import Dataimport


def table_line(reference, unit, quantity, unit_price, amount, widths=(30, 8, 12, 14, 12)):
    return "".join(
        str(value).ljust(width)
        for value, width in zip(
            (reference, unit, quantity, unit_price, amount), widths
        )
    ) + "VAT"


def header(widths=(30, 8, 12, 14, 12)):
    return table_line("Reference & Description", "UM", "Quantity", "Unit Price", "Amount", widths)


class InvoiceLineExtractionTests(unittest.TestCase):
    def test_extracts_numeric_columns_and_zero_values(self):
        text = "\n".join(
            [
                header(),
                table_line("PART-123", "PC", "123", "1.234,5678", "1.234,50"),
                table_line("PART-456", "KG", "0", "0,00", "0,00"),
                "Sub total",
            ]
        )
        self.assertEqual(
            Dataimport.extract_invoice_lines(text),
            [
                {"description": "PART-123", "unit": "PC", "quantity": 123.0,
                 "unit_price": 1234.5678, "amount": 1234.5},
                {"description": "PART-456", "unit": "KG", "quantity": 0.0,
                 "unit_price": 0.0, "amount": 0.0},
            ],
        )

    def test_repeated_headers_realign_columns(self):
        widths = (38, 10, 15, 16, 14)
        text = "\n".join(
            [
                header(),
                table_line("PART-123", "PC", "12", "5,00", "60,00"),
                header(widths),
                table_line("PART-456", "KG", "3", "7,50", "22,50", widths),
                "Sub total",
            ]
        )
        lines = Dataimport.extract_invoice_lines(text)
        self.assertEqual(len(lines), 2)
        self.assertEqual(lines[1], {
            "description": "PART-456", "unit": "KG", "quantity": 3.0,
            "unit_price": 7.5, "amount": 22.5,
        })

    def test_missing_and_invalid_values_remain_empty(self):
        text = "\n".join(
            [
                header(),
                table_line("FREIGHT", "", "", "", "42,00"),
                table_line("PART-123", "PC", "N/A", "bad", "15,00"),
                "Sub total",
            ]
        )
        lines = Dataimport.extract_invoice_lines(text)
        self.assertEqual(
            [(line["unit"], line["quantity"], line["unit_price"]) for line in lines],
            [("", None, None), ("PC", None, None)],
        )

    def test_deduplication_includes_new_columns(self):
        base = {column: "" for column in Dataimport.INVOICE_COLUMNS}
        base.update({"PDF Source": "a.pdf", "UM": "PC", "Quantity": 2.0, "Unit Price": 3.0})
        duplicate = {**base, "PDF Source": "b.pdf"}
        different_unit = {**duplicate, "UM": "KG"}
        different_quantity = {**duplicate, "Quantity": 3.0}
        different_price = {**duplicate, "Unit Price": 4.0}
        rows, omitted = Dataimport.deduplicate_invoice_rows(
            [base, duplicate, different_unit, different_quantity, different_price]
        )
        self.assertEqual(rows, [base, different_unit, different_quantity, different_price])
        self.assertEqual(omitted, {"b.pdf": 1})

    def test_workbook_contains_numeric_values_in_new_columns(self):
        row = {
            "PDF Source": "invoice.pdf", "UM": "PC", "Quantity": 3.0,
            "Unit Price": 2.3456, "Amount": 7.04,
        }
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "invoices.xlsx"
            Dataimport.write_workbook(path, [row], [])
            workbook = load_workbook(path, read_only=True, data_only=True)
            try:
                sheet = workbook["Invoice lines"]
                headers, values = list(sheet.iter_rows(values_only=True))
                result = dict(zip(headers, values))
                self.assertEqual(result["UM"], "PC")
                self.assertEqual(result["Quantity"], 3)
                self.assertEqual(result["Unit Price"], 2.3456)
            finally:
                workbook.close()

    def test_process_pdf_reports_missing_line_values(self):
        class Page:
            def extract_text(self, extraction_mode):
                return "invoice"

        class Reader:
            pages = [Page()]

        lines = [{"description": "Freight", "unit": "", "quantity": None,
                  "unit_price": None, "amount": 12.0}]
        rows = []
        path = Path("folder") / "invoice.pdf"
        with (
            patch.object(Dataimport, "PdfReader", return_value=Reader()),
            patch.object(Dataimport, "has_invoice_layout", return_value=True),
            patch.object(Dataimport, "extract_invoice_fields", return_value={"Invoice #": "123"}),
            patch.object(Dataimport, "find_shipper", return_value=("Shipper", "OK")),
            patch.object(Dataimport, "extract_invoice_lines", return_value=lines),
        ):
            log = Dataimport.process_pdf(path, Path("folder"), rows, {}, {}, {})
        self.assertEqual(log["Status"], "Extracción parcial")
        self.assertIn("UM: 1", log["Details"])
        self.assertIn("Quantity: 1", log["Details"])
        self.assertIn("Unit Price: 1", log["Details"])
        self.assertEqual(rows[0]["Amount"], 12.0)
        self.assertIsNone(rows[0]["Quantity"])


if __name__ == "__main__":
    unittest.main()
