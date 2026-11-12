import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import load_workbook

import Dataimport


ADDRESSES = (
    ("HORSE BRASIL S A, AV RENAULT 1300,Roseira de Sao SeBO83070 -900,SAO JOSE DOS PINHAIS,,BR", "Horse Brasil"),
    ("HORSE BRASIL S A, AV RENAULT 1300,Roseira de Sao SeBO83070-900,SAO JOSE DOS PINHAIS,,BR", "Horse Brasil"),
    ("HORSE BRASIL S A, AV RENAULT 1300,Roseira de Sao SeBORDA DO CAMPO (Setor Parte (83070-900) SAO JOSE DOS PINHAIS", "Horse Brasil"),
    ("HORSE CHILE SPA, SAN RAFAEL 1769 VALPARAISO2100837,LOS ANDES,,CL", "Horse Chile"),
    ("HORSE CHILE SPA, SAN RAFAEL CASILLA 11 A (0) LOS ANDES", "Horse Chile"),
    ("HORSE CHILE SPA, SAN RAFAEL CASILLA 11 A 17890,LOS ANDES,,CL", "Horse Chile"),
    ("OYAK HORSE MAK. EK. A.S., OYAK HOR UM BURSA EKIPMANLARI TICARET e SANAYI A.S Balatosb (16372) BURSA", "Horse Turkey"),
    ("RENAULT ARGENTINA SA, Av Renault 2520 5000 Córdoba Córdoba () Cordoba", "Horse Argentina"),
)


class DestinationPlantTests(unittest.TestCase):
    def test_all_given_variants(self):
        for address, expected in ADDRESSES:
            with self.subTest(address=address):
                self.assertEqual(Dataimport.destination_plant(address), expected)

    def test_normalization_and_invalid_destination(self):
        address = ADDRESSES[-1][0]
        self.assertEqual(Dataimport.destination_plant(address.lower().replace("ó", "o")), "Horse Argentina")
        self.assertEqual(Dataimport.destination_plant(address.replace("ó", "�")), "Horse Argentina")
        self.assertEqual(Dataimport.destination_plant("HORSE BRASIL S A, otra dirección"), "")
        self.assertEqual(Dataimport.destination_plant(""), "")

    def test_processed_row_and_workbook_retain_original_destination(self):
        class Page:
            def extract_text(self, extraction_mode):
                return "invoice"

        class Reader:
            pages = [Page()]

        address = ADDRESSES[0][0]
        lines = [{"description": "Part", "unit": "PC", "quantity": 1.0,
                  "unit_price": 2.0, "amount": 2.0}]
        rows = []
        with (
            patch.object(Dataimport, "PdfReader", return_value=Reader()),
            patch.object(Dataimport, "has_invoice_layout", return_value=True),
            patch.object(Dataimport, "extract_invoice_fields", return_value={"Ship To": address}),
            patch.object(Dataimport, "find_shipper", return_value=("Shipper", "OK")),
            patch.object(Dataimport, "extract_invoice_lines", return_value=lines),
        ):
            log = Dataimport.process_pdf(Path("folder") / "#123#" / "a.pdf", Path("folder"), rows, {}, {}, {})
        self.assertEqual(log["Status"], "Extraído")
        self.assertEqual(rows[0]["Ship To"], address)
        self.assertEqual(rows[0]["Destin Plant"], "Horse Brasil")

        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "output.xlsx"
            Dataimport.write_workbook(path, rows, [log])
            workbook = load_workbook(path, read_only=True, data_only=True)
            try:
                headers, values = list(workbook["Invoice lines"].iter_rows(values_only=True))
                result = dict(zip(headers, values))
                self.assertEqual(result["Ship To"], address)
                self.assertEqual(result["Destin Plant"], "Horse Brasil")
            finally:
                workbook.close()

    def test_unmatched_destination_is_logged(self):
        class Page:
            def extract_text(self, extraction_mode):
                return "invoice"

        class Reader:
            pages = [Page()]

        rows = []
        with (
            patch.object(Dataimport, "PdfReader", return_value=Reader()),
            patch.object(Dataimport, "has_invoice_layout", return_value=True),
            patch.object(Dataimport, "extract_invoice_fields", return_value={"Ship To": "Otro destino"}),
            patch.object(Dataimport, "find_shipper", return_value=("Shipper", "OK")),
            patch.object(Dataimport, "extract_invoice_lines", return_value=[]),
        ):
            log = Dataimport.process_pdf(Path("folder") / "a.pdf", Path("folder"), rows, {}, {}, {})
        self.assertEqual(rows[0]["Destin Plant"], "")
        self.assertEqual(log["Status"], "Extracción parcial")
        self.assertIn("Destin Plant (destino sin equivalencia)", log["Details"])


if __name__ == "__main__":
    unittest.main()
