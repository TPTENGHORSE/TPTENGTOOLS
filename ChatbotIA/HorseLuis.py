import io
import importlib
import importlib.util
import json
import os
import re
import sys
import unicodedata
from datetime import datetime
from pathlib import Path

import pandas as pd
import requests
import streamlit as st


MEMORY_FILE = os.path.join(os.path.dirname(__file__), "horseluis_memory.json")
DEFAULT_MODEL = "llama3"

# Reuse the AirCalc airfreight rate engine (AirCalc/AirCalc.py) instead of duplicating its logic.
_AIRCALC_DIR = os.path.normpath(os.path.join(os.path.dirname(__file__), "..", "AirCalc"))
if _AIRCALC_DIR not in sys.path:
    sys.path.insert(0, _AIRCALC_DIR)
_AIRCALC_IMPORT_ERROR = None
try:
    import AirCalc as aircalc  # noqa: E402
except Exception as exc:
    aircalc = None
    _AIRCALC_IMPORT_ERROR = exc

_PROJECT_ROOT = Path(__file__).resolve().parent.parent
_INTEGRATED_MODULES = {}


def _load_integrated_module(name: str, path: Path):
    if name in _INTEGRATED_MODULES:
        return _INTEGRATED_MODULES[name]
    if not path.is_file():
        raise FileNotFoundError(f"No se encontró la aplicación: {path}")

    module_name = f"_horseluis_{name}"
    spec = importlib.util.spec_from_file_location(module_name, path)
    if spec is None or spec.loader is None:
        raise ImportError(f"No se pudo cargar la aplicación: {path}")

    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    try:
        spec.loader.exec_module(module)
    except Exception:
        sys.modules.pop(module_name, None)
        raise
    _INTEGRATED_MODULES[name] = module
    return module


@st.cache_data(show_spinner=False)
def _load_vtt_data() -> pd.DataFrame:
    data_path = _PROJECT_ROOT / "VTT Tool" / "VTT DATA.xlsx"
    if not data_path.is_file():
        raise FileNotFoundError(f"No se encontró el fichero de VTT: {data_path}")
    return pd.read_excel(data_path)


def _chunk_text(text: str, size: int = 900, overlap: int = 150) -> list[str]:
    cleaned = re.sub(r"\s+", " ", text).strip()
    if not cleaned:
        return []
    chunks = []
    i = 0
    step = max(size - overlap, 1)
    while i < len(cleaned):
        chunks.append(cleaned[i : i + size])
        i += step
    return chunks


def _tokenize(text: str) -> set[str]:
    return set(re.findall(r"[a-z0-9]{3,}", text.lower()))


def _extract_text_and_tables(uploaded_file) -> tuple[str, dict[str, pd.DataFrame], str | None]:
    name = uploaded_file.name
    ext = name.rsplit(".", 1)[-1].lower() if "." in name else ""
    data = uploaded_file.getvalue()

    try:
        if ext in {"txt", "md", "log"}:
            return data.decode("utf-8", errors="ignore"), {}, None

        if ext == "csv":
            df = pd.read_csv(io.BytesIO(data))
            return df.to_csv(index=False), {"data": df}, None

        if ext in {"xlsx", "xls"}:
            sheets = pd.read_excel(io.BytesIO(data), sheet_name=None)
            parts = []
            for sheet_name, df in sheets.items():
                parts.append(f"Sheet: {sheet_name}\n{df.to_csv(index=False)}")
            return "\n\n".join(parts), sheets, None

        if ext == "pdf":
            try:
                pypdf = importlib.import_module("pypdf")
                PdfReader = getattr(pypdf, "PdfReader")
            except Exception:
                return "", {}, "PDF support requires 'pypdf' package."

            reader = PdfReader(io.BytesIO(data))
            pages = []
            for page in reader.pages:
                pages.append(page.extract_text() or "")
            return "\n".join(pages), {}, None

        return "", {}, f"Unsupported file type: .{ext}"
    except Exception as exc:
        return "", {}, f"Failed to read {name}: {exc}"


def _build_kb(files) -> tuple[list[dict], dict[str, dict[str, pd.DataFrame]], list[str]]:
    kb_chunks: list[dict] = []
    tabular_data: dict[str, dict[str, pd.DataFrame]] = {}
    warnings: list[str] = []

    for up_file in files:
        text, tables, err = _extract_text_and_tables(up_file)
        if err:
            warnings.append(f"{up_file.name}: {err}")
            continue

        if tables:
            tabular_data[up_file.name] = tables

        chunks = _chunk_text(text)
        for chunk in chunks:
            kb_chunks.append(
                {
                    "source": up_file.name,
                    "text": chunk,
                    "tokens": _tokenize(chunk),
                }
            )

    return kb_chunks, tabular_data, warnings


def _retrieve(kb_chunks: list[dict], question: str, top_k: int = 3) -> list[dict]:
    q_tokens = _tokenize(question)
    if not q_tokens:
        return []

    scored = []
    for chunk in kb_chunks:
        overlap = len(q_tokens & chunk["tokens"])
        if overlap > 0:
            scored.append((overlap, chunk))

    scored.sort(key=lambda item: item[0], reverse=True)
    return [item[1] for item in scored[:top_k]]


def _fallback_retrieve(kb_chunks: list[dict], top_k: int = 3) -> list[dict]:
    # Provide representative chunks when keyword overlap fails.
    if not kb_chunks:
        return []
    selected = []
    seen_sources = set()
    for chunk in kb_chunks:
        source = chunk.get("source", "")
        if source not in seen_sources:
            selected.append(chunk)
            seen_sources.add(source)
        if len(selected) >= top_k:
            return selected
    return kb_chunks[:top_k]


def _looks_like_doc_request(text: str) -> bool:
    t = text.lower()
    markers = [
        "documento adjunto",
        "adjunto",
        "archivo adjunto",
        "resumen del pdf",
        "resume el pdf",
        "attached document",
        "attached file",
        "summarize the pdf",
    ]
    return any(m in t for m in markers)


def _normalize_memories(raw_data) -> list[dict]:
    normalized = []
    next_id = 1
    if not isinstance(raw_data, list):
        return normalized

    for item in raw_data:
        if not isinstance(item, dict):
            continue
        text = str(item.get("text", "")).strip()
        if not text:
            continue
        normalized.append(
            {
                "id": int(item.get("id", next_id)),
                "text": text,
                "category": str(item.get("category", "general")),
                "priority": int(item.get("priority", 3)),
                "confidence": float(item.get("confidence", 0.8)),
                "source": str(item.get("source", "manual")),
                "uses": int(item.get("uses", 0)),
                "created_at": str(
                    item.get(
                        "created_at",
                        datetime.utcnow().isoformat(timespec="seconds") + "Z",
                    )
                ),
                "last_used": str(item.get("last_used", "")),
            }
        )
        next_id = max(next_id, normalized[-1]["id"] + 1)

    normalized.sort(key=lambda m: m["id"])
    return normalized


def _load_memory() -> list[dict]:
    if not os.path.exists(MEMORY_FILE):
        return []
    try:
        with open(MEMORY_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        return _normalize_memories(data)
    except Exception:
        return []


def _save_memory(memories: list[dict]) -> None:
    try:
        with open(MEMORY_FILE, "w", encoding="utf-8") as f:
            json.dump(memories, f, ensure_ascii=False, indent=2)
    except Exception:
        # Keep the app usable even if writing the memory file fails.
        pass


def _add_memory(
    memories: list[dict],
    text: str,
    source: str,
    category: str = "general",
    priority: int = 3,
    confidence: float = 0.9,
) -> list[dict]:
    clean = re.sub(r"\s+", " ", text).strip()
    if not clean:
        return memories
    lowered = clean.lower()
    existing = {m.get("text", "").strip().lower() for m in memories}
    if lowered in existing:
        return memories

    next_id = max([int(m.get("id", 0)) for m in memories], default=0) + 1
    memories.append(
        {
            "id": next_id,
            "text": clean,
            "category": category,
            "priority": int(max(min(priority, 5), 1)),
            "confidence": float(max(min(confidence, 1.0), 0.0)),
            "source": source,
            "uses": 0,
            "created_at": datetime.utcnow().isoformat(timespec="seconds") + "Z",
            "last_used": "",
        }
    )
    return memories


def _auto_extract_memory(user_text: str) -> str | None:
    patterns = [
        r"^\s*recuerda que\s+(.+)$",
        r"^\s*aprende que\s+(.+)$",
        r"^\s*remember that\s+(.+)$",
        r"^\s*my name is\s+(.+)$",
        r"^\s*mi nombre es\s+(.+)$",
    ]
    for pattern in patterns:
        match = re.search(pattern, user_text, flags=re.IGNORECASE)
        if match:
            return match.group(1).strip(" .")
    return None


def _retrieve_memory(memories: list[dict], question: str, top_k: int = 4) -> list[dict]:
    q_tokens = _tokenize(question)
    if not q_tokens:
        return []

    scored = []
    for mem in memories:
        text = str(mem.get("text", ""))
        overlap = len(q_tokens & _tokenize(text))
        if overlap > 0:
            score = (
                (overlap * 2.0)
                + (float(mem.get("priority", 3)) * 0.6)
                + (float(mem.get("confidence", 0.8)) * 0.8)
            )
            scored.append((score, mem))

    scored.sort(key=lambda item: item[0], reverse=True)
    return [item[1] for item in scored[:top_k]]


def _mark_memories_used(memories: list[dict], used_ids: set[int]) -> list[dict]:
    now = datetime.utcnow().isoformat(timespec="seconds") + "Z"
    updated = []
    for mem in memories:
        if int(mem.get("id", 0)) in used_ids:
            mem["uses"] = int(mem.get("uses", 0)) + 1
            mem["last_used"] = now
        updated.append(mem)
    return updated


def _apply_memory_edits(memories: list[dict], edited_df: pd.DataFrame) -> list[dict]:
    by_id = {int(m["id"]): m for m in memories}
    for _, row in edited_df.iterrows():
        mem_id = int(row["id"])
        if mem_id not in by_id:
            continue
        by_id[mem_id]["text"] = str(row["text"]).strip()
        by_id[mem_id]["category"] = str(row["category"]).strip() or "general"
        by_id[mem_id]["priority"] = int(max(min(int(row["priority"]), 5), 1))
        by_id[mem_id]["confidence"] = float(max(min(float(row["confidence"]), 1.0), 0.0))
    cleaned = [m for m in by_id.values() if m.get("text")]
    cleaned.sort(key=lambda m: int(m.get("id", 0)))
    return cleaned


def _delete_memories(memories: list[dict], ids_to_delete: set[int]) -> list[dict]:
    return [m for m in memories if int(m.get("id", 0)) not in ids_to_delete]


def _compute_quote(distance_km: float, rate_per_km: float, fixed_cost: float, fuel_pct: float) -> dict:
    linehaul = distance_km * rate_per_km
    subtotal = linehaul + fixed_cost
    fuel_surcharge = subtotal * (fuel_pct / 100.0)
    total = subtotal + fuel_surcharge
    return {
        "linehaul": round(linehaul, 2),
        "subtotal": round(subtotal, 2),
        "fuel_surcharge": round(fuel_surcharge, 2),
        "total": round(total, 2),
    }


def _airfreight_results_to_markdown(df: pd.DataFrame) -> str:
    header = "| " + " | ".join(df.columns.astype(str)) + " |"
    sep = "| " + " | ".join(["---"] * len(df.columns)) + " |"
    rows = []
    for _, row in df.iterrows():
        cells = []
        for col in df.columns:
            val = row[col]
            if col == "Total Airfreight":
                cells.append(f"{val:,.0f} €")
            elif col == "Eur/Kg":
                cells.append(f"{val:,.2f}")
            else:
                cells.append(str(val))
        rows.append("| " + " | ".join(cells) + " |")
    return "\n".join([header, sep, *rows])


def _airfreight_quote(origin: str, destin: str, weight: float, volume: float) -> str:
    """Price a lane using the same rate engine as AirCalc/AirCalc.py (Calculator_v2)."""
    if aircalc is None:
        return "El motor de tarifas AirCalc no está disponible (no se pudo importar AirCalc.py)."
    if not os.path.exists(aircalc.XLSM_PATH):
        return f"No se encontró el fichero de tarifas: {aircalc.XLSM_PATH.name}"

    rate_table = aircalc.load_rate_table()
    dhl_adder = aircalc.load_dhl_adder()

    lanes = rate_table[
        (rate_table["Country Code"] == origin) & (rate_table["Country Code Destin"] == destin)
    ]
    if lanes.empty:
        return f"No se encontraron lanes aéreos de {origin} a {destin}."

    cbw = max(volume * 167, weight)
    if cbw <= 0:
        return "Indica peso (kg) y/o volumen (m3) para calcular la tarifa."

    results = aircalc.compute_rates(lanes, cbw, dhl_adder)
    results = results[list(aircalc.DISPLAY_COLUMNS.keys())].rename(columns=aircalc.DISPLAY_COLUMNS)
    results["Service Level"] = results["Service Level"].str.title()
    results["Lane Type"] = results["Lane Type"].str.title()
    service_level_rank = results["Service Level"].map(aircalc.SERVICE_LEVEL_ORDER).fillna(
        len(aircalc.SERVICE_LEVEL_ORDER)
    )
    lane_type_rank = results["Lane Type"].map(aircalc.LANE_TYPE_ORDER).fillna(len(aircalc.LANE_TYPE_ORDER))
    results = (
        results.assign(_slr=service_level_rank, _ltr=lane_type_rank)
        .sort_values(["ILN/Supplier", "_ltr", "Origin Airport", "Destin Airport", "_slr"])
        .drop(columns=["_slr", "_ltr"])
        .reset_index(drop=True)
    )

    summary = f"**CBW (peso facturable):** {cbw:,.2f} kg — {origin} → {destin}\n\n"
    return summary + _airfreight_results_to_markdown(results)


_AIRFREIGHT_WEIGHT_RE = re.compile(r"(\d+(?:[.,]\d+)?)\s*kg", re.IGNORECASE)
_AIRFREIGHT_VOLUME_RE = re.compile(r"(\d+(?:[.,]\d+)?)\s*(?:m3|m³|metros?\s*c[uú]bicos)", re.IGNORECASE)
_AIRFREIGHT_ROUTE_RE = re.compile(
    r"(?:de|desde|from|origen|origin)\s*[:\-]?\s*([a-zA-Z]{2})\b.{0,40}?"
    r"(?:a|hacia|to|destino|destination)\s*[:\-]?\s*([a-zA-Z]{2})\b",
    re.IGNORECASE,
)


def _extract_airfreight_request(text: str, valid_origins: set, valid_destins: set) -> dict | None:
    weight_match = _AIRFREIGHT_WEIGHT_RE.search(text)
    volume_match = _AIRFREIGHT_VOLUME_RE.search(text)
    if not weight_match and not volume_match:
        return None

    route_match = _AIRFREIGHT_ROUTE_RE.search(text)
    if not route_match:
        return None

    origin = route_match.group(1).upper()
    destin = route_match.group(2).upper()
    if origin not in valid_origins or destin not in valid_destins:
        return None

    weight = float(weight_match.group(1).replace(",", ".")) if weight_match else 0.0
    volume = float(volume_match.group(1).replace(",", ".")) if volume_match else 0.0
    return {"origin": origin, "destin": destin, "weight": weight, "volume": volume}


def _normalize_text(text: str) -> str:
    normalized = unicodedata.normalize("NFKD", text.casefold())
    return "".join(char for char in normalized if not unicodedata.combining(char))


def _mentions_any(text: str, phrases: tuple[str, ...]) -> bool:
    normalized = _normalize_text(text)
    return any(_normalize_text(phrase) in normalized for phrase in phrases)


def _parse_number(value: str) -> float:
    normalized = value.strip().replace(" ", "")
    if "," in normalized and "." in normalized:
        decimal_separator = "," if normalized.rfind(",") > normalized.rfind(".") else "."
        thousands_separator = "." if decimal_separator == "," else ","
        normalized = normalized.replace(thousands_separator, "").replace(decimal_separator, ".")
    elif "," in normalized:
        trailing_digits = len(normalized) - normalized.rfind(",") - 1
        normalized = normalized.replace(",", "" if trailing_digits == 3 else ".")
    return float(normalized)


def _labeled_number(text: str, labels: tuple[str, ...]) -> float | None:
    number = r"([-+]?\d[\d.,]*)"
    for label in sorted(labels, key=len, reverse=True):
        pattern = rf"(?<!\w){re.escape(label)}(?!\w)\s*(?::|=|is|es)?\s*{number}"
        match = re.search(pattern, _normalize_text(text), flags=re.IGNORECASE)
        if match:
            return _parse_number(match.group(1))
    return None


def _fcf_calculation_context(question: str) -> str:
    calculations = _load_integrated_module(
        "fcf_calculations", _PROJECT_ROOT / "FCF_TDC" / "calculations.py"
    )
    scenario_matches = list(
        re.finditer(r"\b(?:escenario|scenario)\s*([ab])\b", question, flags=re.IGNORECASE)
    )
    scenario_text = {}
    for index, match in enumerate(scenario_matches):
        end = scenario_matches[index + 1].start() if index + 1 < len(scenario_matches) else len(question)
        scenario_text[match.group(1).upper()] = question[match.start():end]

    field_labels = {
        "ocean": ("flete maritimo", "ocean freight", "ocean"),
        "inland": ("transporte interior", "transporte terrestre", "inland transport", "inland"),
        "other": ("logistica adicional", "coste adicional", "costo adicional", "additional logistics", "otros costes", "otros costos"),
        "clt": ("lead time", "clt", "plazo total", "dias de entrega"),
    }
    values = {}
    for scenario in ("A", "B"):
        segment = scenario_text.get(scenario, "")
        values[scenario] = {
            name: _labeled_number(segment, labels)
            for name, labels in field_labels.items()
        }

    cargo_value = _labeled_number(
        question,
        ("valor de carga por contenedor", "valor mercancia por contenedor", "cargo value per container", "valor por contenedor"),
    )
    annual_volume = _labeled_number(
        question,
        ("contenedores anuales", "volumen anual", "annual container volume", "volumen de contenedores"),
    )
    financing_rate = _labeled_number(
        question, ("wacc", "tasa financiera", "annual financing rate", "tasa de financiacion")
    )

    context_lines = [
        "FCF_TDC compara dos escenarios con las funciones calculate_scenario y calculate_comparison.",
        "Entradas necesarias por escenario: flete marítimo, transporte interior, logística adicional y CLT en días.",
        "Entradas globales: valor de carga por contenedor, volumen anual de contenedores y WACC/tasa anual.",
        "Las salidas incluyen reducción de CLT, capital circulante liberado, ahorro anual de financiación, ahorro anual de transporte y beneficio económico anual neto.",
    ]
    required = all(
        value is not None
        for scenario in values.values()
        for value in scenario.values()
    ) and cargo_value is not None and annual_volume is not None and financing_rate is not None
    if not required:
        context_lines.append(
            "No se ejecutó una comparación: faltan entradas explícitas. Pide solo los valores que falten; no inventes datos."
        )
        return "\n".join(context_lines)

    if financing_rate > 1:
        financing_rate /= 100
    scenario_metrics = {}
    for label in ("A", "B"):
        inputs = values[label]
        scenario = calculations.ScenarioInputs(
            ocean_freight_per_container=inputs["ocean"],
            inland_transport_per_container=inputs["inland"],
            additional_logistics_per_container=inputs["other"],
            customer_lead_time_days=inputs["clt"],
        )
        scenario_metrics[label] = calculations.calculate_scenario(
            scenario,
            cargo_value_per_container=cargo_value,
            containers_consumed_per_day=annual_volume / calculations.DEFAULT_DAYS_PER_YEAR,
            annual_container_volume=annual_volume,
            annual_financing_rate=financing_rate,
        )
    comparison = calculations.calculate_comparison(
        scenario_metrics["A"],
        scenario_metrics["B"],
        annual_container_volume=annual_volume,
    )
    context_lines.append(
        "Cálculo ejecutado con los valores indicados "
        f"(valor/contenedor={cargo_value:g}, contenedores/año={annual_volume:g}, "
        f"tasa anual={financing_rate:.2%}; días/año={calculations.DEFAULT_DAYS_PER_YEAR})."
    )
    for label in ("A", "B"):
        metrics = scenario_metrics[label]
        context_lines.append(
            f"Escenario {label}: CLT={metrics.analysis_days:g} días; "
            f"coste de transporte/contenedor={metrics.transportation_cost_per_container:,.2f}; "
            f"valor de inventario en tránsito={metrics.pipeline_inventory_value:,.2f}; "
            f"coste económico anual total={metrics.annual_total_economic_cost:,.2f}."
        )
    context_lines.append(
        "Comparación B frente a A: "
        f"reducción CLT={comparison['transit_time_reduction_days']:,.2f} días; "
        f"capital circulante liberado={comparison['working_capital_released']:,.2f}; "
        f"ahorro anual de financiación={comparison['annual_financing_savings']:,.2f}; "
        f"ahorro anual de transporte={comparison['annual_transportation_savings']:,.2f}; "
        f"beneficio económico anual neto={comparison['net_annual_economic_benefit']:,.2f}."
    )
    return "\n".join(context_lines)


def _empower_calculation_context(question: str) -> str:
    empower = _load_integrated_module(
        "empower3d", _PROJECT_ROOT / "Packaging" / "Empower3D.py"
    )
    container = None
    normalized_question = _normalize_text(question)
    for name in empower.DIMENSIONES_OPERATIVAS:
        if _normalize_text(name) in normalized_question:
            container = name
            break
    if container is None:
        aliases = (
            ("Mega Trailer 90m3", ("mega trailer", "90m3")),
            ("Container 40 HC", ("40 hc", "40 high cube")),
            ("Container 20 Ft Std", ("20 ft", "20ft", "20 pies")),
            ("Trailer 40m3", ("trailer 40m3", "40 m3 trailer")),
            ("India 32-ft MXL", ("india 32", "32-ft mxl")),
        )
        for name, terms in aliases:
            if _mentions_any(question, terms) and name in empower.DIMENSIONES_OPERATIVAS:
                container = name
                break

    dimensions_match = re.search(
        r"(?<!\d)(\d{2,5})\s*(?:mm)?\s*[x×]\s*(\d{2,5})\s*(?:mm)?\s*[x×]\s*(\d{2,5})\s*(?:mm)?\b",
        normalized_question,
    )
    pn_weight = _labeled_number(question, ("pn weight", "peso pn", "peso de la pieza", "peso pieza"))
    ucm_weight = _labeled_number(question, ("ucm weight", "peso ucm", "peso del ucm", "tara ucm"))
    pieces_per_ucm = _labeled_number(
        question, ("pn/ucm", "piezas por ucm", "piezas/ucm", "parts per ucm")
    )
    if pieces_per_ucm is None:
        pieces_match = re.search(
            r"(\d+(?:[.,]\d+)?)\s*(?:piezas?|parts|pn)\s*(?:por|/)\s*ucm",
            normalized_question,
        )
        if pieces_match:
            pieces_per_ucm = _parse_number(pieces_match.group(1))
    stacking = _labeled_number(question, ("stacking", "apilamiento", "capas"))

    context_lines = [
        "Empower3D calcula cuántas UCM caben por volumen usando la función calcula_cajas, "
        "y limita el resultado por el peso máximo del contenedor.",
        "Para calcular hacen falta tipo de contenedor, dimensiones de UCM (largo x ancho x alto en mm), "
        "peso de PN (kg), peso de UCM vacía (kg) y PN/UCM. El apilamiento por defecto es 1 capa.",
        "Dimensiones y pesos máximos de la aplicación: "
        + "; ".join(
            f"{name}: interior {empower.DIMENSIONES_INTERNAS[name]} mm, "
            f"operativa {empower.DIMENSIONES_OPERATIVAS[name]} mm, "
            f"peso máximo {empower.PESOS_MAXIMOS[name]} kg"
            for name in empower.DIMENSIONES_INTERNAS
        )
        + ".",
    ]
    if (
        container is None
        or dimensions_match is None
        or pn_weight is None
        or ucm_weight is None
        or pieces_per_ucm is None
    ):
        context_lines.append("No se ejecutó el cálculo porque faltan entradas explícitas.")
        return "\n".join(context_lines)

    box_dimensions = tuple(int(value) for value in dimensions_match.groups())
    if min(box_dimensions) <= 0 or pn_weight <= 0 or ucm_weight < 0 or pieces_per_ucm <= 0:
        raise ValueError("Las dimensiones, el peso de PN y PN/UCM deben ser positivos; el peso de UCM no puede ser negativo.")
    if not pieces_per_ucm.is_integer():
        raise ValueError("PN/UCM debe ser un número entero.")
    max_stacking = int(stacking) if stacking is not None else 1
    if stacking is not None and not stacking.is_integer():
        raise ValueError("El número de capas de apilamiento debe ser entero.")
    if max_stacking <= 0:
        raise ValueError("El apilamiento debe ser de al menos 1 capa.")

    operative_dimensions = empower.DIMENSIONES_OPERATIVAS[container]
    capacity_by_volume, rotation, distribution = empower.calcula_cajas(
        operative_dimensions, box_dimensions, 9999
    )
    max_stacking_possible = distribution[0][2]
    capacity_by_volume, rotation, distribution = empower.calcula_cajas(
        operative_dimensions, box_dimensions, min(max_stacking, max_stacking_possible)
    )
    box_weight = pn_weight * pieces_per_ucm + ucm_weight
    if box_weight <= 0:
        raise ValueError("El peso total de una UCM debe ser mayor que cero.")
    max_by_weight = int(empower.PESOS_MAXIMOS[container] // box_weight)
    realistic_ucm = min(capacity_by_volume, max_by_weight)
    context_lines.append(
        f"Cálculo Empower3D ejecutado: contenedor={container}; dimensiones UCM={box_dimensions} mm; "
        f"peso UCM cargada={box_weight:,.3f} kg; orientación={rotation}; "
        f"capacidad por volumen={capacity_by_volume}; capacidad por peso={max_by_weight}; "
        f"capacidad final={realistic_ucm} UCM, limitada por "
        f"{'peso' if realistic_ucm == max_by_weight else 'volumen'}; "
        f"PN totales={int(pieces_per_ucm) * realistic_ucm}."
    )
    return "\n".join(context_lines)


def _vtt_context(question: str) -> str:
    data = _load_vtt_data()
    preferred_columns = [
        "ID",
        "Destin Country",
        "POL",
        "POD",
        "Carrier",
        " Shipper",
        "Transit time",
        "9 ETD>ETA",
        "15 Due Date",
        "16 Manufacturing",
        "Safety stock",
        "Parts Vanning",
        "Expiration Date",
    ]
    columns = [column for column in preferred_columns if column in data.columns]
    query_tokens = _tokenize(question)
    requested_id = re.search(r"\b[A-Z]\d{3}\b", question.upper())
    scored_rows = []
    for index, row in data.iterrows():
        row_values = [str(row[column]) for column in columns if pd.notna(row[column])]
        row_tokens = _tokenize(" ".join(row_values))
        score = len(query_tokens & row_tokens)
        if requested_id and str(row.get("ID", "")).upper() == requested_id.group(0):
            score += 100
        if score:
            scored_rows.append((score, index, row))
    scored_rows.sort(key=lambda item: item[0], reverse=True)

    lines = [
        f"VTT Tool: datos leídos de VTT DATA.xlsx ({len(data)} registros).",
        "Campos consultados: ID, destino, POL, POD, carrier, shipper, transit time, ETD/ETA, due date, manufacturing y safety stock.",
    ]
    if not scored_rows:
        lines.append("No se encontró ningún registro con coincidencia en estos datos para la consulta.")
        return "\n".join(lines)

    for _, _, row in scored_rows[:4]:
        values = [
            f"{column.strip()}: {row[column]}"
            for column in columns
            if pd.notna(row[column]) and str(row[column]).strip()
        ]
        lines.append("Registro: " + "; ".join(values))
    return "\n".join(lines)


def _aircalc_context(question: str) -> str:
    if aircalc is None:
        detail = f": {_AIRCALC_IMPORT_ERROR}" if _AIRCALC_IMPORT_ERROR else ""
        raise ImportError(f"No se pudo importar AirCalc{detail}")
    if not aircalc.XLSM_PATH.is_file():
        raise FileNotFoundError(f"No se encontró el fichero de tarifas AirCalc: {aircalc.XLSM_PATH}")

    rate_table = aircalc.load_rate_table()
    route_match = _AIRFREIGHT_ROUTE_RE.search(question)
    if route_match:
        origin, destin = route_match.group(1).upper(), route_match.group(2).upper()
        if (
            origin in set(rate_table["Country Code"].dropna().unique())
            and destin in set(rate_table["Country Code Destin"].dropna().unique())
        ):
            lanes = rate_table[
                (rate_table["Country Code"] == origin)
                & (rate_table["Country Code Destin"] == destin)
            ]
            if lanes.empty:
                return f"AirCalc no tiene lanes para {origin} → {destin}."
            return (
                f"AirCalc usa CBW=max(volumen en m³ × 167, peso en kg) y el motor de tarifas del tender 2026. "
                f"Para {origin} → {destin} hay {len(lanes)} lanes. "
                "Detalle disponible: "
                + "; ".join(
                    ", ".join(
                        f"{label}={row[column]}"
                        for column, label in (
                            ("Shipper/ILN", "ILN"),
                            ("Type of Flow", "lane"),
                            ("Service Level (Normal, Economy and Express)", "service"),
                            ("Origin Airport", "origin airport"),
                            ("Destination Airport", "destination airport"),
                            ("Provider_Name", "carrier"),
                            ("Total TT (days)", "transit days"),
                        )
                        if column in lanes.columns and pd.notna(row[column])
                    )
                    for _, row in lanes.head(8).iterrows()
                )
                + ". Para calcular el precio se necesitan peso y/o volumen."
            )
    return (
        "AirCalc calcula peso facturable CBW=max(volumen en m³ × 167, peso en kg) "
        "y aplica el motor de tarifas del tender 2026. Para una cotización indica "
        "país de origen, país de destino y peso (kg) y/o volumen (m³)."
    )


def _integration_contexts(question: str) -> tuple[list[tuple[str, str]], list[str]]:
    contexts = []
    errors = []
    integrations = [
        (
            "AirCalc",
            ("aircalc", "airfreight", "air freight", "aereo", "vuelo", "tarifa aerea", "cotizacion aerea"),
            _aircalc_context,
        ),
        (
            "FCF_TDC",
            ("fcf", "finanzas", "financiero", "wacc", "capital circulante", "ahorro", "escenario", "lead time", "clt", "coste logistico", "costo logistico", "coste de transporte", "costo de transporte", "working capital", "economic benefit"),
            _fcf_calculation_context,
        ),
        (
            "Empower3D",
            ("empower3d", "packaging", "embalaje", "ucm", "caja", "palet", "pallet", "container 40", "container 20", "40 hc", "20 ft", "20ft", "40m3", "90m3", "32-ft mxl"),
            _empower_calculation_context,
        ),
        (
            "VTT Tool",
            ("vtt", "pol", "pod", "transit time", "fecha de llegada", "due date", "etd", "eta", "shipper", "carrier", "ruta maritima"),
            _vtt_context,
        ),
    ]
    for source, triggers, context_builder in integrations:
        is_vtt_id = source == "VTT Tool" and re.search(r"\b[A-Z]\d{3}\b", question.upper())
        packaging_context = (
            source == "Empower3D"
            and _mentions_any(question, ("contenedor", "container"))
            and (
                not _mentions_any(
                    question,
                    ("fcf", "finanzas", "wacc", "flete", "coste", "costo", "ahorro", "escenario"),
                )
                or _mentions_any(question, ("dimensiones", "medidas", "cajas", "ucm", "peso"))
            )
        )
        if not _mentions_any(question, triggers) and not is_vtt_id and not packaging_context:
            continue
        try:
            contexts.append((source, context_builder(question)))
        except Exception as exc:
            errors.append(f"{source}: {type(exc).__name__}: {exc}")
    return contexts, errors


def run():
    st.markdown("## HorseLuis")
    st.caption(
        "Asistente con memoria, documentos y consultas conectadas a AirCalc, FCF_TDC, "
        "Empower3D y VTT Tool."
    )

    OLLAMA_URL = "http://localhost:11434/api/chat"

    if "messages" not in st.session_state:
        st.session_state["messages"] = []
    if "kb_chunks" not in st.session_state:
        st.session_state["kb_chunks"] = []
    if "kb_sources" not in st.session_state:
        st.session_state["kb_sources"] = []
    if "tabular_data" not in st.session_state:
        st.session_state["tabular_data"] = {}
    if "memory_entries" not in st.session_state:
        st.session_state["memory_entries"] = _load_memory()
    if "auto_learn" not in st.session_state:
        st.session_state["auto_learn"] = True
    if "model_name" not in st.session_state:
        st.session_state["model_name"] = DEFAULT_MODEL
    if "temperature" not in st.session_state:
        st.session_state["temperature"] = 0.2

    with st.sidebar:
        st.header("Assistant")
        st.text_input("Model", key="model_name")
        st.slider("Temperature", 0.0, 1.5, key="temperature", step=0.1)
        st.checkbox("Strict document mode", key="strict_docs_mode", value=False)
        colsb1, colsb2 = st.columns(2)
        with colsb1:
            if st.button("New chat"):
                st.session_state["messages"] = []
                st.rerun()
        with colsb2:
            if st.button("Clear memory"):
                st.session_state["memory_entries"] = []
                _save_memory([])
                st.rerun()

        st.divider()
        st.subheader("Knowledge Base")
        files = st.file_uploader(
            "Upload docs",
            type=["txt", "csv", "xlsx", "xls", "pdf", "md", "log"],
            accept_multiple_files=True,
        )
        if st.button("Index documents"):
            if not files:
                st.warning("Upload at least one document.")
            else:
                kb_chunks, tabular_data, warnings = _build_kb(files)
                st.session_state["kb_chunks"] = kb_chunks
                st.session_state["tabular_data"] = tabular_data
                st.session_state["kb_sources"] = sorted({c["source"] for c in kb_chunks})
                st.success(
                    f"Indexed {len(kb_chunks)} chunks from {len(st.session_state['kb_sources'])} files."
                )
                for warning in warnings:
                    st.warning(warning)

            st.caption("Index documents procesa los archivos subidos y los vuelve consultables en el chat.")

        st.caption(
            "Indexed files: " + ", ".join(st.session_state["kb_sources"])
            if st.session_state["kb_sources"]
            else "No indexed files"
        )

        st.divider()
        st.subheader("Calculation Tools")
        distance_km = st.number_input("Distance (km)", min_value=0.0, value=100.0)
        rate_per_km = st.number_input("Rate per km", min_value=0.0, value=2.0)
        fixed_cost = st.number_input("Fixed cost", min_value=0.0, value=35.0)
        fuel_pct = st.number_input("Fuel surcharge %", min_value=0.0, value=8.0)
        if st.button("Calculate quote"):
            quote = _compute_quote(distance_km, rate_per_km, fixed_cost, fuel_pct)
            tool_msg = (
                "Quote calculation\n"
                f"Linehaul: {quote['linehaul']}\n"
                f"Subtotal: {quote['subtotal']}\n"
                f"Fuel surcharge: {quote['fuel_surcharge']}\n"
                f"Total: {quote['total']}"
            )
            st.session_state["messages"].append(
                {"role": "assistant", "content": tool_msg, "sources": ["quotation_tool"]}
            )
            st.success(f"Total: {quote['total']}")

        st.divider()
        st.subheader("Airfreight Tool (AirCalc)")
        if aircalc is None or not os.path.exists(getattr(aircalc, "XLSM_PATH", "")):
            st.caption("AirCalc no está disponible: revisa la importación y el fichero de tarifas.")
        else:
            air_rate_table = aircalc.load_rate_table()
            air_origin_options = sorted(air_rate_table["Country Code"].dropna().unique())
            air_origin = st.selectbox("Origin Country", air_origin_options, key="air_origin")
            air_destin_options = sorted(
                air_rate_table.loc[
                    air_rate_table["Country Code"] == air_origin, "Country Code Destin"
                ].dropna().unique()
            )
            air_destin = st.selectbox("Destin Country", air_destin_options, key="air_destin")
            air_weight = st.number_input("Weight (kg)", min_value=0.0, value=0.0, key="air_weight")
            air_volume = st.number_input("Volume (m³)", min_value=0.0, value=0.0, step=0.1, key="air_volume")
            if st.button("Calculate airfreight quote"):
                air_msg = _airfreight_quote(air_origin, air_destin, air_weight, air_volume)
                st.session_state["messages"].append(
                    {"role": "assistant", "content": air_msg, "sources": ["AirCalc"]}
                )
                st.rerun()

        st.divider()
        st.subheader("Excel Tool")
        tabular_data = st.session_state["tabular_data"]
        if tabular_data:
            source_name = st.selectbox("File", options=list(tabular_data.keys()))
            sheet_name = st.selectbox("Sheet", options=list(tabular_data[source_name].keys()))
            df_sheet = tabular_data[source_name][sheet_name]
            numeric_columns = [
                col
                for col in df_sheet.columns
                if pd.api.types.is_numeric_dtype(df_sheet[col])
            ]
            if numeric_columns:
                col_name = st.selectbox("Numeric column", options=numeric_columns)
                op_name = st.selectbox("Operation", options=["sum", "avg", "min", "max"])
                if st.button("Run Excel tool"):
                    series = df_sheet[col_name].dropna()
                    if op_name == "sum":
                        result = float(series.sum())
                    elif op_name == "avg":
                        result = float(series.mean())
                    elif op_name == "min":
                        result = float(series.min())
                    else:
                        result = float(series.max())

                    excel_msg = (
                        f"Excel tool result\nFile: {source_name}\nSheet: {sheet_name}\n"
                        f"Column: {col_name}\nOperation: {op_name}\nResult: {round(result, 4)}"
                    )
                    st.session_state["messages"].append(
                        {"role": "assistant", "content": excel_msg, "sources": [source_name]}
                    )
                    st.success(f"Result: {round(result, 4)}")
            else:
                st.info("No numeric columns available in this sheet.")
        else:
            st.caption("Index CSV/XLSX files to enable Excel tools.")

    with st.expander("Long-term memory", expanded=False):
        st.checkbox("Auto-learn from 'recuerda que' / 'remember that'", key="auto_learn")
        mem_col1, mem_col2, mem_col3, mem_col4 = st.columns([2, 1, 1, 1])
        with mem_col1:
            teach_text = st.text_input("Teach a memory", key="teach_memory_input")
        with mem_col2:
            teach_category = st.selectbox(
                "Category",
                options=["general", "customer", "pricing", "route", "personal"],
                key="teach_memory_category",
            )
        with mem_col3:
            teach_priority = st.slider("Priority", min_value=1, max_value=5, value=3, key="teach_priority")
        with mem_col4:
            teach_conf = st.slider("Confidence", min_value=0.0, max_value=1.0, value=0.9, step=0.1, key="teach_conf")

        mem_actions1, mem_actions2 = st.columns(2)
        with mem_actions1:
            if st.button("Save memory") and teach_text.strip():
                st.session_state["memory_entries"] = _add_memory(
                    st.session_state["memory_entries"],
                    teach_text,
                    source="manual",
                    category=teach_category,
                    priority=teach_priority,
                    confidence=teach_conf,
                )
                _save_memory(st.session_state["memory_entries"])
                st.success("Memory saved")
        with mem_actions2:
            if st.button("Clear all memory"):
                st.session_state["memory_entries"] = []
                _save_memory([])
                st.warning("Memory cleared")

        mem_df = pd.DataFrame(st.session_state["memory_entries"])
        if not mem_df.empty:
            editable_cols = ["id", "text", "category", "priority", "confidence", "source", "uses"]
            editor_df = mem_df[editable_cols].copy()
            edited_df = st.data_editor(
                editor_df,
                use_container_width=True,
                num_rows="fixed",
                disabled=["id", "source", "uses"],
                key="memory_editor",
            )
            apply_col1, apply_col2 = st.columns(2)
            with apply_col1:
                if st.button("Apply memory edits"):
                    st.session_state["memory_entries"] = _apply_memory_edits(
                        st.session_state["memory_entries"], edited_df
                    )
                    _save_memory(st.session_state["memory_entries"])
                    st.success("Memory updated")
            with apply_col2:
                delete_ids = st.multiselect("Delete by id", options=editor_df["id"].tolist())
                if st.button("Delete selected") and delete_ids:
                    st.session_state["memory_entries"] = _delete_memories(
                        st.session_state["memory_entries"], set(int(i) for i in delete_ids)
                    )
                    _save_memory(st.session_state["memory_entries"])
                    st.success("Selected memories deleted")

    for msg in st.session_state["messages"]:
        role = "assistant" if msg["role"] == "assistant" else "user"
        with st.chat_message(role):
            st.markdown(msg["content"])
            if msg.get("sources"):
                st.caption("Sources: " + ", ".join(msg["sources"]))
            if msg.get("memories"):
                st.caption("Memory used: " + "; ".join(msg["memories"]))

    user_input = st.chat_input("Send a message to HorseLuis")

    if user_input:
        st.session_state["messages"].append({"role": "user", "content": user_input})

        if aircalc is not None and os.path.exists(getattr(aircalc, "XLSM_PATH", "")):
            air_rate_table = aircalc.load_rate_table()
            air_request = _extract_airfreight_request(
                user_input,
                set(air_rate_table["Country Code"].dropna().unique()),
                set(air_rate_table["Country Code Destin"].dropna().unique()),
            )
            if air_request:
                air_msg = _airfreight_quote(**air_request)
                st.session_state["messages"].append(
                    {"role": "assistant", "content": air_msg, "sources": ["AirCalc"]}
                )
                st.rerun()

        if _looks_like_doc_request(user_input) and not st.session_state["kb_chunks"]:
            bot_reply = (
                "No hay documentos indexados todavía. "
                "Sube tu archivo en 'Upload docs' y luego pulsa 'Index documents'."
            )
            st.session_state["messages"].append(
                {"role": "assistant", "content": bot_reply, "sources": []}
            )
            st.rerun()

        if st.session_state.get("auto_learn", True):
            extracted = _auto_extract_memory(user_input)
            if extracted:
                st.session_state["memory_entries"] = _add_memory(
                    st.session_state["memory_entries"],
                    extracted,
                    source="auto",
                    category="personal",
                    priority=4,
                    confidence=0.8,
                )
                _save_memory(st.session_state["memory_entries"])

        integration_contexts, integration_errors = _integration_contexts(user_input)
        for integration_error in integration_errors:
            st.warning(f"No se pudo consultar una aplicación integrada: {integration_error}")

        retrieved = _retrieve(st.session_state["kb_chunks"], user_input, top_k=3)
        if not retrieved and st.session_state["kb_chunks"]:
            retrieved = _fallback_retrieve(st.session_state["kb_chunks"], top_k=3)
        context_blocks = [f"Source: {c['source']}\n{c['text']}" for c in retrieved]
        context_blocks.extend(
            f"Application source: {source}\n{context}"
            for source, context in integration_contexts
        )
        if integration_errors:
            context_blocks.extend(
                f"Application source unavailable: {error}" for error in integration_errors
            )
        context_text = "\n\n---\n\n".join(context_blocks)
        retrieved_memory = _retrieve_memory(st.session_state["memory_entries"], user_input, top_k=4)
        memory_text = "\n".join(f"- {m['text']}" for m in retrieved_memory)

        strict_docs_mode = bool(st.session_state.get("strict_docs_mode", False))
        if strict_docs_mode:
            system_prompt = (
                "You are HorseLuis, a helpful assistant. "
                "Use uploaded documents, integrated application context and long-term memory as primary evidence. "
                "Treat application calculations and records as authoritative; do not invent missing values. "
                "If a requested application source is unavailable, tell the user using its reported error. "
                "If the answer is not present in the available context, say clearly that it was not found there."
            )
        else:
            system_prompt = (
                "You are HorseLuis, a helpful general-purpose assistant. "
                "Use uploaded documents, integrated application context and long-term memory when relevant, "
                "but you can also answer general questions with your own model knowledge. "
                "When sources are used, prefer them for specific factual details. "
                "Treat application calculations and records as authoritative; do not invent missing values. "
                "If a requested application source is unavailable, report its error and do not guess."
            )

        user_prompt = user_input
        prompt_parts = []
        if memory_text:
            prompt_parts.append("Long-term memory facts:\n" + memory_text)
        if context_text:
            prompt_parts.append("Context from documents and connected applications:\n" + context_text)
        if prompt_parts:
            user_prompt = "\n\n".join(prompt_parts) + f"\n\nQuestion: {user_input}"

        ollama_messages = [{"role": "system", "content": system_prompt}]
        for m in st.session_state["messages"][:-1]:
            ollama_messages.append({"role": m["role"], "content": m["content"]})
        ollama_messages.append({"role": "user", "content": user_prompt})

        payload = {
            "model": st.session_state.get("model_name", DEFAULT_MODEL),
            "messages": ollama_messages,
            "options": {"temperature": float(st.session_state.get("temperature", 0.2))},
        }

        full_reply = ""
        with st.chat_message("assistant"):
            response_placeholder = st.empty()
        try:
            response = requests.post(OLLAMA_URL, json=payload, timeout=120, stream=True)
            response.raise_for_status()
            for line in response.iter_lines():
                if not line:
                    continue
                try:
                    data = json.loads(line.decode("utf-8"))
                    content = data.get("message", {}).get("content")
                    if content:
                        full_reply += content
                        response_placeholder.markdown(full_reply + "▌")
                except Exception:
                    continue
            bot_reply = full_reply if full_reply else "[No reply]"
            response_placeholder.markdown(bot_reply)
        except Exception as exc:
            bot_reply = f"Error: {exc}"
            response_placeholder.markdown(bot_reply)

        sources = sorted(
            {c["source"] for c in retrieved}
            | {source for source, _ in integration_contexts}
            | {error.split(":", 1)[0] for error in integration_errors}
        )
        memory_ids_used = {int(m["id"]) for m in retrieved_memory}
        memories_used = [m["text"] for m in retrieved_memory]
        st.session_state["memory_entries"] = _mark_memories_used(
            st.session_state["memory_entries"], memory_ids_used
        )
        _save_memory(st.session_state["memory_entries"])
        st.session_state["messages"].append(
            {
                "role": "assistant",
                "content": bot_reply,
                "sources": sources,
                "memories": memories_used,
            }
        )

if __name__ == "__main__":
    run()
