import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import load_workbook

import Dataimport


class SourceIdTests(unittest.TestCase):
    def test_extracts_ids_from_source_path(self):
        examples = (
            (r"#123#\invoice.pdf", "#123#"),
            (r"#123-456#\invoice.pdf", "#123-456#"),
            (r"#123#\#456#\invoice.pdf", "#123#, #456#"),
            (r"#123#\copy #123#.pdf", "#123#"),
            (r"# 123 #\invoice.pdf", "# 123 #"),
            (r"folder\invoice.pdf", ""),
            (r"#123\folder#\invoice.pdf", ""),
        )
        for source, expected in examples:
            with self.subTest(source=source):
                self.assertEqual(Dataimport.extract_source_ids(source), expected)

    def test_distinct_ids_do_not_collapse_different_invoice_rows(self):
        first = {"ID": "#123#", "PDF Source": r"#123#\a.pdf"}
        same = {"ID": "#123#", "PDF Source": r"#123#\b.pdf"}
        different = {"ID": "#456#", "PDF Source": r"#456#\c.pdf"}
        rows, omitted = Dataimport.deduplicate_invoice_rows([first, same, different])
        self.assertEqual(rows, [first, different])
        self.assertEqual(omitted, {same["PDF Source"]: 1})

    def test_process_pdf_writes_id_and_reports_missing_id(self):
        class Page:
            def extract_text(self, extraction_mode):
                return "invoice"

        class Reader:
            pages = [Page()]

        fields = {"Ship To": "HORSE CHILE SPA, SAN RAFAEL CASILLA 11 A (0) LOS ANDES"}
        line = {"description": "Part", "unit": "PC", "quantity": 1.0,
                "unit_price": 2.0, "amount": 2.0}
        with (
            patch.object(Dataimport, "PdfReader", return_value=Reader()),
            patch.object(Dataimport, "has_invoice_layout", return_value=True),
            patch.object(Dataimport, "extract_invoice_fields", side_effect=lambda text: dict(fields)),
            patch.object(Dataimport, "find_shipper", return_value=("Shipper", "OK")),
            patch.object(Dataimport, "extract_invoice_lines", return_value=[line]),
        ):
            rows = []
            log = Dataimport.process_pdf(Path("root") / "#123#" / "invoice.pdf", Path("root"), rows, {}, {}, {})
            missing_rows = []
            missing_log = Dataimport.process_pdf(Path("root") / "invoice.pdf", Path("root"), missing_rows, {}, {}, {})
        self.assertEqual(rows[0]["ID"], "#123#")
        self.assertEqual(log["Status"], "Extraído")
        self.assertEqual(missing_rows[0]["ID"], "")
        self.assertEqual(missing_log["Status"], "Extracción parcial")
        self.assertIn("ID", missing_log["Details"])

        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "output.xlsx"
            Dataimport.write_workbook(path, rows, [log])
            workbook = load_workbook(path, read_only=True, data_only=True)
            try:
                headers, values = list(workbook["Invoice lines"].iter_rows(values_only=True))
                self.assertEqual(dict(zip(headers, values))["ID"], "#123#")
            finally:
                workbook.close()


if __name__ == "__main__":
    unittest.main()
