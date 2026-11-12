"""Extract air-freight invoice fields from PDFs into an Excel workbook."""

from __future__ import annotations

import argparse
import os
import re
import unicodedata
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, TypedDict

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from pypdf import PdfReader


SOURCE_DIR = Path(
    r"C:\Users\OLMEDOGALVEZJorgeLui\OneDrive - Horse\Shortcuts"
    r"\EMERGENCY AIFREIGHTS HORSE - Documentos\EMERGENCY AIFREIGHTS HORSE"
    r"\LATAM Emergency AFs\DOCUMENTATION"
)
OUTPUT_PATH = Path(__file__).with_name("Facturas_Aereas_Extraidas.xlsx")

INVOICE_FIELDS = (
    ("Invoice Date", re.compile(r"\binvoice\s*date\b", re.IGNORECASE)),
    ("Ship To", re.compile(r"\bship\s*to\b", re.IGNORECASE)),
    ("Net weight", re.compile(r"\bnet\s*weight\b", re.IGNORECASE)),
    ("Gross weight", re.compile(r"\bgross\s*weight\b", re.IGNORECASE)),
    (
        "Number of packages",
        re.compile(r"\bnumber\s+of\s+packages\b", re.IGNORECASE),
    ),
)
INVOICE_NUMBER_LABEL_RE = re.compile(
    r"invoice\s*(?:#|number\b|no\.?\b)", re.IGNORECASE
)
SHIPPER_LABEL_RE = re.compile(
    r"shipper(?:['’]s)?\s+name\s+(?:and|&)\s+(?:address|addr\.?)",
    re.IGNORECASE,
)
CONSIGNEE_LABEL_RE = re.compile(
    r"consignee(?:['’]s)?\s+name\s+(?:and|&)\s+(?:address|addr\.?)",
    re.IGNORECASE,
)
WAYBILL_NUMBER_RE = re.compile(
    r"\b(?:H|M)?AWB\b(?:(?:\s*(?:No\.?|Number))\s*[:#_]?\s*|"
    r"\s*[:#_]\s*|\s+)([A-Z0-9][A-Z0-9-]*)",
    re.IGNORECASE,
)
REFERENCE_HEADER_RE = re.compile(
    r"reference\s*(?:&|and)\s*description", re.IGNORECASE
)
MONEY_RE = re.compile(
    r"(?<!\d)(?:\d{1,3}(?:[.,]\d{3})+|\d+)[.,]\d{2}(?!\d)"
)
DATE_RE = re.compile(
    r"\b(?:\d{1,2}[-/. ](?:JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)"
    r"[-/. ]\d{2,4}|\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}|"
    r"\d{4}[-/.]\d{1,2}[-/.]\d{1,2})\b",
    re.IGNORECASE,
)

INVOICE_COLUMNS = [
    "ID",
    "Invoice Date",
    "Invoice #",
    "Ship To",
    "Destin Plant",
    "Shipper",
    "Net weight",
    "Gross weight",
    "Number of packages",
    "Reference & Description",
    "UM",
    "Quantity",
    "Unit Price",
    "Amount",
    "PDF Source",
]
LOG_COLUMNS = ["PDF Source", "Status", "Details"]


class InvoiceLine(TypedDict):
    description: str
    unit: str
    quantity: float | None
    unit_price: float | None
    amount: float | None


def clean_cell_text(value: str) -> str:
    return re.sub(r"\s+", " ", value).strip()


def extract_source_ids(source: str) -> str:
    matches = re.findall(r"#[^#\\/\r\n]+#", source)
    return ", ".join(dict.fromkeys(matches))


def normalized_destination(value: str) -> str:
    decomposed = unicodedata.normalize("NFKD", value).casefold()
    return "".join(character for character in decomposed if character.isalnum())


DESTINATION_PLANTS = {
    normalized_destination(address): plant
    for plant, addresses in (
        ("Horse Brasil", (
            "HORSE BRASIL S A, AV RENAULT 1300,Roseira de Sao SeBO83070 -900,SAO JOSE DOS PINHAIS,,BR",
            "HORSE BRASIL S A, AV RENAULT 1300,Roseira de Sao SeBO83070-900,SAO JOSE DOS PINHAIS,,BR",
            "HORSE BRASIL S A, AV RENAULT 1300,Roseira de Sao SeBORDA DO CAMPO (Setor Parte (83070-900) SAO JOSE DOS PINHAIS",
        )),
        ("Horse Chile", (
            "HORSE CHILE SPA, SAN RAFAEL 1769 VALPARAISO2100837,LOS ANDES,,CL",
            "HORSE CHILE SPA, SAN RAFAEL CASILLA 11 A (0) LOS ANDES",
            "HORSE CHILE SPA, SAN RAFAEL CASILLA 11 A 17890,LOS ANDES,,CL",
        )),
        ("Horse Turkey", (
            "OYAK HORSE MAK. EK. A.S., OYAK HOR UM BURSA EKIPMANLARI TICARET e SANAYI A.S Balatosb (16372) BURSA",
        )),
        ("Horse Argentina", (
            "RENAULT ARGENTINA SA, Av Renault 2520 5000 Córdoba Córdoba () Cordoba",
            "RENAULT ARGENTINA SA, Av Renault 2520 5000 C�rdoba C�rdoba () Cordoba",
        )),
    )
    for address in addresses
}


def destination_plant(ship_to: str) -> str:
    return DESTINATION_PLANTS.get(normalized_destination(ship_to), "")


def has_invoice_layout(text: str) -> bool:
    normalized = clean_cell_text(text)
    if not all(pattern.search(normalized) for _, pattern in INVOICE_FIELDS):
        return False

    return any(
        REFERENCE_HEADER_RE.search(line) and re.search(r"\bamount\b", line, re.I)
        for line in text.splitlines()
    )


def extract_labeled_value(
    lines: list[str], label_pattern: re.Pattern[str], value_pattern: re.Pattern[str]
) -> str:
    for line_index, line in enumerate(lines):
        match = label_pattern.search(line)
        if match is None:
            continue

        candidates = [line[match.end() :]]
        if line_index + 1 < len(lines):
            candidates.append(lines[line_index + 1])
        for candidate in candidates:
            value_match = value_pattern.search(candidate)
            if value_match:
                return clean_cell_text(value_match.group(0))
    return ""


def extract_invoice_date(lines: list[str]) -> datetime | str:
    for line_index, line in enumerate(lines):
        label_match = INVOICE_FIELDS[0][1].search(line)
        if label_match is None:
            continue

        dates = [
            (abs(candidate_index - line_index), match.group(0))
            for candidate_index, candidate in enumerate(lines)
            for match in DATE_RE.finditer(candidate)
            if abs(candidate_index - line_index) <= 1
        ]
        if not dates:
            return ""

        date_text = min(dates, key=lambda value: value[0])[1]
        for date_format in (
            "%d-%b-%Y",
            "%d-%b-%y",
            "%d-%B-%Y",
            "%d-%B-%y",
            "%d/%m/%Y",
            "%d/%m/%y",
            "%m/%d/%Y",
            "%m/%d/%y",
            "%d-%m-%Y",
            "%d-%m-%y",
            "%Y-%m-%d",
            "%Y/%m/%d",
        ):
            try:
                return datetime.strptime(date_text, date_format)
            except ValueError:
                continue
        return date_text
    return ""


def extract_ship_to(lines: list[str]) -> str:
    label_pattern = INVOICE_FIELDS[1][1]
    stop_pattern = re.compile(
        r"^(?:incoterm|it\s+trucks?\s+id|net\s*weight|gross\s*weight|"
        r"number\s+of\s+packages|payment)\b",
        re.IGNORECASE,
    )

    for line_index, line in enumerate(lines):
        match = label_pattern.search(line)
        if match is None:
            continue

        parts = [clean_cell_text(line[match.end() :])]
        for following_line in lines[line_index + 1 : line_index + 3]:
            next_text = clean_cell_text(following_line)
            if not next_text:
                continue
            if stop_pattern.match(next_text):
                break
            parts.append(next_text)
            break
        return ", ".join(part for part in parts if part)
    return ""


def extract_invoice_fields(text: str) -> dict[str, Any]:
    lines = text.splitlines()
    weight_pattern = re.compile(
        r"\d[\d.,]*(?:\s*(?:kg|kgs|kilograms?))?", re.IGNORECASE
    )
    package_pattern = re.compile(r"\d+")
    invoice_number_pattern = re.compile(r"[A-Z0-9][A-Z0-9/-]*", re.IGNORECASE)
    return {
        "Invoice Date": extract_invoice_date(lines),
        "Invoice #": extract_labeled_value(
            lines, INVOICE_NUMBER_LABEL_RE, invoice_number_pattern
        ),
        "Ship To": extract_ship_to(lines),
        "Net weight": extract_labeled_value(
            lines, INVOICE_FIELDS[2][1], weight_pattern
        ),
        "Gross weight": extract_labeled_value(
            lines, INVOICE_FIELDS[3][1], weight_pattern
        ),
        "Number of packages": extract_labeled_value(
            lines, INVOICE_FIELDS[4][1], package_pattern
        ),
    }


def normalized_hawb(value: str) -> str:
    return re.sub(r"[^A-Z0-9]", "", value.upper())


def extract_waybill_numbers(text: str) -> set[str]:
    return {
        normalized_hawb(match.group(1))
        for match in WAYBILL_NUMBER_RE.finditer(text)
        if normalized_hawb(match.group(1))
    }


def extract_awb_shipper(text: str) -> str:
    lines = text.splitlines()
    for line_index, line in enumerate(lines):
        shipper_label = SHIPPER_LABEL_RE.search(line)
        if shipper_label is None:
            continue

        account_label = re.search(
            r"shipper['’]s\s+account\s+number", line, re.IGNORECASE
        )
        left_column_end = account_label.start() if account_label else None
        if account_label is not None:
            for address_line in lines[line_index + 1 :]:
                if CONSIGNEE_LABEL_RE.search(address_line):
                    break
                waybill_heading = re.search(
                    r"\bAIR\s+WAYBILL\b", address_line, re.IGNORECASE
                )
                if (
                    waybill_heading is not None
                    and waybill_heading.start() > account_label.start()
                    and not address_line[
                        account_label.start() : waybill_heading.start()
                    ].strip()
                ):
                    left_column_end = waybill_heading.start()
                    break
        address_lines: list[str] = []
        for address_line in lines[line_index + 1 :]:
            if CONSIGNEE_LABEL_RE.search(address_line):
                break
            left_column = (
                address_line[:left_column_end]
                if left_column_end is not None
                else address_line
            )
            value = clean_cell_text(left_column)
            if value:
                address_lines.append(value)
        return ", ".join(address_lines)

    terms_index = next(
        (
            index
            for index, line in enumerate(lines)
            if re.search(
                r"\bIt is agreed that the goods described herein\b",
                line,
                re.IGNORECASE,
            )
        ),
        None,
    )
    if terms_index is None:
        return ""

    header_index = next(
        (
            index
            for index, line in enumerate(lines[:terms_index])
            if WAYBILL_NUMBER_RE.search(line)
        ),
        None,
    )
    if header_index is None:
        return ""

    body_lines = lines[header_index + 1 : terms_index]
    right_column_start = None
    for line in body_lines:
        tokens = list(re.finditer(r"\S+", line))
        for previous, current in zip(tokens, tokens[1:]):
            if current.start() - previous.end() >= 2:
                right_column_start = current.start()
                break
        if right_column_start is not None:
            break
    if right_column_start is None:
        return ""

    address_lines = [
        clean_cell_text(line[:right_column_start])
        for line in body_lines
        if clean_cell_text(line[:right_column_start])
    ]
    return ", ".join(address_lines)


def read_awb_record(pdf_path: Path) -> dict[str, str] | None:
    reader = PdfReader(str(pdf_path), strict=False)
    if not reader.pages:
        return None
    text = reader.pages[0].extract_text(extraction_mode="layout") or ""
    if not SHIPPER_LABEL_RE.search(text) and not (
        WAYBILL_NUMBER_RE.search(text)
        and re.search(r"\bIt is agreed that the goods described herein\b", text, re.I)
    ):
        return None

    shipper = extract_awb_shipper(text)
    if not shipper:
        return None

    waybill_numbers = extract_waybill_numbers(text)
    if not waybill_numbers:
        filename_number = WAYBILL_NUMBER_RE.search(
            pdf_path.stem.replace("_", " ")
        )
        if filename_number:
            waybill_numbers.add(normalized_hawb(filename_number.group(1)))
    return {
        "shipper": shipper,
        "waybill_numbers": ",".join(sorted(waybill_numbers)),
        "source": pdf_path.name,
    }


def shipment_folder(pdf_path: Path, source_dir: Path) -> Path:
    relative_parts = pdf_path.relative_to(source_dir).parts
    return source_dir / relative_parts[0] if len(relative_parts) > 1 else source_dir


def find_shipper(
    invoice_path: Path,
    source_dir: Path,
    invoice_text: str,
    pdfs_by_folder: dict[Path, list[Path]],
    awb_cache: dict[Path, list[dict[str, str]]],
    awb_error_cache: dict[Path, list[str]],
) -> tuple[str, str]:
    folder = shipment_folder(invoice_path, source_dir)
    if folder not in awb_cache:
        records: list[dict[str, str]] = []
        errors: list[str] = []
        for pdf_path in pdfs_by_folder.get(folder, []):
            if pdf_path == invoice_path:
                continue
            try:
                record = read_awb_record(pdf_path)
            except Exception as error:
                errors.append(f"{pdf_path.name}: {type(error).__name__}: {error}")
                continue
            if record is not None:
                records.append(record)
        awb_cache[folder] = records
        awb_error_cache[folder] = errors

    records = awb_cache[folder]
    errors = awb_error_cache[folder]
    error_details = (
        " Errores al leer posibles AWB: " + "; ".join(errors)
        if errors
        else ""
    )
    invoice_waybills = extract_waybill_numbers(invoice_text)
    if invoice_waybills:
        matches = [
            record
            for record in records
            if invoice_waybills.intersection(
                record["waybill_numbers"].split(",")
            )
        ]
        if not matches:
            return "", (
                "No se encontró en la carpeta un AWB con número "
                + ", ".join(sorted(invoice_waybills))
                + error_details
            )
    elif len(records) == 1:
        matches = records
    elif not records:
        return "", (
            "No se encontró un AWB con texto extraíble en la carpeta"
            + error_details
        )
    else:
        return "", (
            "Hay varios AWB en la carpeta y la factura no indica un número AWB"
            + error_details
        )

    shippers = {record["shipper"] for record in matches}
    if len(shippers) != 1:
        return "", "El HAWB coincide con más de un shipper" + error_details
    matching_awbs = ", ".join(record["source"] for record in matches)
    return (
        shippers.pop(),
        f"Shipper obtenido de {matching_awbs}." + error_details,
    )


def parse_amount(amount_text: str) -> float | None:
    value = amount_text.strip().replace(" ", "")
    if not value:
        return None

    if "," in value and "." in value:
        decimal_separator = "," if value.rfind(",") > value.rfind(".") else "."
        grouping_separator = "." if decimal_separator == "," else ","
        value = value.replace(grouping_separator, "")
        value = value.replace(decimal_separator, ".")
    elif "," in value:
        value = value.replace(".", "").replace(",", ".")

    try:
        return float(Decimal(value))
    except InvalidOperation:
        return None


def is_reference_code(value: str) -> bool:
    compact_value = value.replace(" ", "")
    return bool(
        len(compact_value) >= 6
        and re.fullmatch(r"[A-Z0-9][A-Z0-9._/-]*", compact_value, re.IGNORECASE)
        and not compact_value.upper().startswith("HS")
    )


def extract_invoice_lines(text: str) -> list[InvoiceLine]:
    lines = text.splitlines()
    header_index = next(
        (
            index
            for index, line in enumerate(lines)
            if REFERENCE_HEADER_RE.search(line) and re.search(r"\bamount\b", line, re.I)
        ),
        None,
    )
    if header_index is None:
        return []

    header = lines[header_index]
    reference_header = REFERENCE_HEADER_RE.search(header)
    if reference_header is None:
        return []
    def column_positions(line: str) -> tuple[int, int, int | None, int | None, int, int] | None:
        reference = REFERENCE_HEADER_RE.search(line)
        amount = re.search(r"\bAmount\b", line, re.IGNORECASE)
        if reference is None or amount is None:
            return None
        unit = re.search(r"\bUM\b", line[reference.end():amount.start()], re.IGNORECASE)
        if unit is None:
            return None
        unit_start = reference.end() + unit.start()
        quantity = re.search(r"\bQuantity\b", line[unit_start:amount.start()], re.IGNORECASE)
        price = re.search(r"\bUnit\s+Price\b", line[unit_start:amount.start()], re.IGNORECASE)
        quantity_start = unit_start + quantity.start() if quantity else None
        price_start = unit_start + price.start() if price else None
        if quantity_start is None or price_start is None or quantity_start >= price_start:
            quantity_start = None
            price_start = None
        tax = re.search(r"\bV\.?\s*A\.?\s*T\b", line[amount.start():], re.IGNORECASE)
        amount_end = amount.start() + tax.start() if tax else len(line)
        return reference.start(), unit_start, quantity_start, price_start, amount.start(), amount_end

    positions = column_positions(header)
    if positions is None:
        return []

    pending_descriptions: list[str] = []
    active_descriptions: list[str] = []
    active_amount: float | None = None
    active_unit = ""
    active_quantity: float | None = None
    active_unit_price: float | None = None
    records: list[InvoiceLine] = []

    def finish_record() -> None:
        nonlocal active_descriptions, active_amount, active_unit, active_quantity, active_unit_price
        description = " | ".join(
            part for part in (clean_cell_text(value) for value in active_descriptions) if part
        )
        if description:
            records.append({
                "description": description,
                "unit": active_unit,
                "quantity": active_quantity,
                "unit_price": active_unit_price,
                "amount": active_amount,
            })
        active_descriptions = []
        active_amount = None
        active_unit = ""
        active_quantity = None
        active_unit_price = None

    for line in lines[header_index + 1 :]:
        if re.search(r"\bsub\s*total\b", line, re.IGNORECASE):
            break

        if REFERENCE_HEADER_RE.search(line) and re.search(r"\bamount\b", line, re.I):
            if active_amount is not None:
                finish_record()
            positions = column_positions(line)
            if positions is None:
                raise ValueError("No se pudieron identificar las columnas del encabezado de factura.")
            continue

        reference_start, unit_start, quantity_start, price_start, amount_start, amount_end = positions
        description = clean_cell_text(line[reference_start:unit_start])
        amount_cell = line[amount_start:amount_end]
        amount_match = MONEY_RE.search(amount_cell)
        if amount_match:
            if active_amount is not None:
                finish_record()
            active_descriptions = pending_descriptions + ([description] if description else [])
            pending_descriptions = []
            active_amount = parse_amount(amount_match.group(0))
            if quantity_start is not None and price_start is not None:
                active_unit = clean_cell_text(line[unit_start:quantity_start])
                quantity_text = clean_cell_text(line[quantity_start:price_start])
                price_text = clean_cell_text(line[price_start:amount_start])
                if re.fullmatch(r"[+-]?\d[\d.,]*", quantity_text):
                    active_quantity = parse_amount(quantity_text)
                if re.fullmatch(r"[+-]?\d[\d.,]*", price_text):
                    active_unit_price = parse_amount(price_text)
        elif description:
            if active_amount is None:
                pending_descriptions.append(description)
            elif description.upper().startswith("HS") and re.match(
                r"HS\d", description, re.IGNORECASE
            ):
                active_descriptions.append(description)
            elif is_reference_code(description):
                pending_descriptions.append(description)
            else:
                active_descriptions.append(description)
        elif active_amount is not None:
            finish_record()

    if active_amount is not None:
        finish_record()
    return records


def iter_pdf_paths(source_dir: Path, processing_log: list[dict[str, str]]) -> list[Path]:
    pdf_paths: list[Path] = []

    def record_walk_error(error: OSError) -> None:
        processing_log.append(
            {
                "PDF Source": str(error.filename or source_dir),
                "Status": "No se pudo acceder a la carpeta",
                "Details": str(error),
            }
        )

    for directory, _, filenames in os.walk(source_dir, onerror=record_walk_error):
        pdf_paths.extend(
            Path(directory) / filename
            for filename in filenames
            if filename.lower().endswith(".pdf")
        )
    return sorted(pdf_paths)


def relative_source(path: Path, source_dir: Path) -> str:
    try:
        return str(path.relative_to(source_dir))
    except ValueError:
        return str(path)


def process_pdf(
    pdf_path: Path,
    source_dir: Path,
    invoice_rows: list[dict[str, Any]],
    pdfs_by_folder: dict[Path, list[Path]],
    awb_cache: dict[Path, list[dict[str, str]]],
    awb_error_cache: dict[Path, list[str]],
) -> dict[str, str]:
    source_name = relative_source(pdf_path, source_dir)
    try:
        reader = PdfReader(str(pdf_path), strict=False)
        if not reader.pages:
            return {
                "PDF Source": source_name,
                "Status": "PDF vacío",
                "Details": "El documento no contiene páginas.",
            }

        first_page = reader.pages[0].extract_text(extraction_mode="layout") or ""
        if not first_page.strip():
            return {
                "PDF Source": source_name,
                "Status": "Sin texto extraíble",
                "Details": "Puede ser un PDF escaneado; requiere OCR.",
            }
        if not has_invoice_layout(first_page):
            return {
                "PDF Source": source_name,
                "Status": "No coincide con el tipo de factura",
                "Details": "No contiene todas las etiquetas de la factura objetivo.",
            }

        page_texts = [first_page]
        page_texts.extend(
            page.extract_text(extraction_mode="layout") or ""
            for page in reader.pages[1:]
        )
        full_text = "\n".join(page_texts)
        fields = extract_invoice_fields(first_page)
        fields["ID"] = extract_source_ids(source_name)
        fields["Destin Plant"] = destination_plant(fields.get("Ship To", ""))
        fields["Shipper"], shipper_details = find_shipper(
            pdf_path,
            source_dir,
            full_text,
            pdfs_by_folder,
            awb_cache,
            awb_error_cache,
        )
        lines = extract_invoice_lines(full_text)
        missing_fields = [
            name for name, value in fields.items() if not value
        ]
        if not fields["Destin Plant"]:
            missing_fields.remove("Destin Plant")
            missing_fields.append("Destin Plant (destino sin equivalencia)")
        if not lines:
            missing_fields.append("Reference & Description / Amount")

        if lines:
            for line in lines:
                invoice_rows.append(
                    {
                        **fields,
                        "Reference & Description": line["description"],
                        "UM": line["unit"],
                        "Quantity": line["quantity"],
                        "Unit Price": line["unit_price"],
                        "Amount": line["amount"],
                        "PDF Source": source_name,
                    }
                )
        else:
            invoice_rows.append(
                {
                    **fields,
                    "Reference & Description": "",
                    "UM": "",
                    "Quantity": None,
                    "Unit Price": None,
                    "Amount": None,
                    "PDF Source": source_name,
                }
            )

        line_warnings = []
        for field, key in (("UM", "unit"), ("Quantity", "quantity"), ("Unit Price", "unit_price")):
            missing_count = sum(line[key] is None or line[key] == "" for line in lines)
            if missing_count:
                line_warnings.append(f"{field}: {missing_count} línea(s)")
        if line_warnings:
            missing_fields.append(
                "valores no encontrados (pueden no figurar): " + ", ".join(line_warnings)
            )

        if missing_fields:
            return {
                "PDF Source": source_name,
                "Status": "Extracción parcial",
                "Details": (
                    "No se pudieron extraer: "
                    + ", ".join(missing_fields)
                    + f". {shipper_details}"
                ),
            }
        return {
            "PDF Source": source_name,
            "Status": "Extraído",
            "Details": f"{len(lines)} línea(s) de factura. {shipper_details}",
        }
    except Exception as error:
        return {
            "PDF Source": source_name,
            "Status": "Error al leer el PDF",
            "Details": f"{type(error).__name__}: {error}",
        }


def deduplicate_invoice_rows(
    invoice_rows: list[dict[str, Any]],
) -> tuple[list[dict[str, Any]], dict[str, int]]:
    data_columns = [column for column in INVOICE_COLUMNS if column != "PDF Source"]
    occurrences_by_source: dict[tuple[tuple[Any, ...], str], int] = {}
    retained_counts: dict[tuple[Any, ...], int] = {}
    omitted_by_source: dict[str, int] = {}
    unique_rows: list[dict[str, Any]] = []

    for row in invoice_rows:
        key = tuple(row.get(column) for column in data_columns)
        source = row["PDF Source"]
        source_key = (key, source)
        occurrence = occurrences_by_source.get(source_key, 0) + 1
        occurrences_by_source[source_key] = occurrence
        if occurrence <= retained_counts.get(key, 0):
            omitted_by_source[source] = omitted_by_source.get(source, 0) + 1
            continue
        retained_counts[key] = occurrence
        unique_rows.append(row)

    return unique_rows, omitted_by_source


def style_worksheet(worksheet: Any) -> None:
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions
    header_fill = PatternFill(fill_type="solid", fgColor="1F4E78")
    for cell in worksheet[1]:
        cell.font = Font(color="FFFFFF", bold=True)
        cell.fill = header_fill
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    worksheet.row_dimensions[1].height = 30

    for column_index, column_cells in enumerate(worksheet.columns, start=1):
        max_length = max(
            (len(str(cell.value)) for cell in column_cells if cell.value is not None),
            default=0,
        )
        worksheet.column_dimensions[get_column_letter(column_index)].width = min(
            max(max_length + 2, 14), 60
        )


def write_workbook(
    output_path: Path,
    invoice_rows: list[dict[str, Any]],
    processing_log: list[dict[str, str]],
) -> None:
    workbook = Workbook()
    invoices_sheet = workbook.active
    invoices_sheet.title = "Invoice lines"
    invoices_sheet.append(INVOICE_COLUMNS)
    for row in invoice_rows:
        invoices_sheet.append([row.get(column) for column in INVOICE_COLUMNS])
        date_cell = invoices_sheet.cell(invoices_sheet.max_row, 1)
        if isinstance(date_cell.value, datetime):
            date_cell.number_format = "dd-mmm-yyyy"
        for column in ("Unit Price", "Amount"):
            cell = invoices_sheet.cell(invoices_sheet.max_row, INVOICE_COLUMNS.index(column) + 1)
            if cell.value is not None:
                cell.number_format = "#,##0.00########"

    log_sheet = workbook.create_sheet("Processing log")
    log_sheet.append(LOG_COLUMNS)
    for row in processing_log:
        log_sheet.append([row.get(column, "") for column in LOG_COLUMNS])

    style_worksheet(invoices_sheet)
    style_worksheet(log_sheet)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Extrae datos de facturas aéreas PDF y los guarda en Excel."
    )
    parser.add_argument(
        "--source",
        type=Path,
        default=SOURCE_DIR,
        help="Carpeta raíz que contiene los PDFs (incluye subcarpetas).",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=OUTPUT_PATH,
        help="Ruta del Excel que se generará.",
    )
    args = parser.parse_args()
    source_dir = args.source.expanduser()
    output_path = args.output.expanduser()
    if not source_dir.is_dir():
        parser.error(f"No existe la carpeta de documentos: {source_dir}")

    processing_log: list[dict[str, str]] = []
    pdf_paths = iter_pdf_paths(source_dir, processing_log)
    if not pdf_paths:
        parser.error(f"No se encontraron PDFs en: {source_dir}")

    invoice_rows: list[dict[str, Any]] = []
    pdfs_by_folder: dict[Path, list[Path]] = {}
    for pdf_path in pdf_paths:
        folder = shipment_folder(pdf_path, source_dir)
        pdfs_by_folder.setdefault(folder, []).append(pdf_path)
    awb_cache: dict[Path, list[dict[str, str]]] = {}
    awb_error_cache: dict[Path, list[str]] = {}
    for pdf_path in pdf_paths:
        processing_log.append(
            process_pdf(
                pdf_path,
                source_dir,
                invoice_rows,
                pdfs_by_folder,
                awb_cache,
                awb_error_cache,
            )
        )

    invoice_rows, omitted_by_source = deduplicate_invoice_rows(invoice_rows)
    for entry in processing_log:
        omitted_count = omitted_by_source.get(entry["PDF Source"], 0)
        if omitted_count and entry["Status"] in {"Extraído", "Extracción parcial"}:
            entry["Details"] += (
                f" {omitted_count} línea(s) omitida(s) por coincidir con otro PDF."
            )

    write_workbook(output_path, invoice_rows, processing_log)
    matched_count = sum(
        entry["Status"] in {"Extraído", "Extracción parcial"}
        for entry in processing_log
    )
    partial_count = sum(
        entry["Status"] == "Extracción parcial" for entry in processing_log
    )
    unreadable_count = sum(
        entry["Status"] in {"Error al leer el PDF", "Sin texto extraíble"}
        for entry in processing_log
    )
    print(f"PDFs revisados: {len(pdf_paths)}")
    print(f"Facturas identificadas: {matched_count}")
    print(f"Extracciones parciales: {partial_count}")
    print(f"PDFs ilegibles o sin texto: {unreadable_count}")
    print(f"Líneas duplicadas omitidas: {sum(omitted_by_source.values())}")
    print(f"Excel generado: {output_path.resolve()}")


if __name__ == "__main__":
    main()