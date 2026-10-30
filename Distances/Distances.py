import os
from math import radians, cos, sin, asin, sqrt
from typing import Optional, Tuple

import pandas as pd

# Optional offline geocoding fallback
try:
    import pgeocode  # type: ignore
except Exception:  # pragma: no cover
    pgeocode = None


GEO_FILE_NAME = "GEO_LOCATIONS.xlsx"
GEO_SHEET = "GEO"


def haversine_km(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    """Great-circle distance between two points on Earth in kilometers."""
    # convert decimal degrees to radians
    lon1, lat1, lon2, lat2 = map(radians, [lon1, lat1, lon2, lat2])
    # haversine formula
    dlon = lon2 - lon1
    dlat = lat2 - lat1
    a = sin(dlat / 2) ** 2 + cos(lat1) * cos(lat2) * sin(dlon / 2) ** 2
    c = 2 * asin(sqrt(a))
    km = 6371.0 * c
    return km


class GeoIndex:
    """In-memory index of optional local geocoding data.

    Expected Excel structure (optional): GEO_LOCATIONS.xlsx, sheet 'GEO' with columns:
      - type: one of {ZIP, CITY, PORT, PLANT, COUNTRY}
      - country_code: ISO2 code (e.g., ES, FR)
      - key: lookup key (ZIP string, city/port/plant name uppercase, country code)
      - lat: float
      - lon: float
    """

    def __init__(self, df: Optional[pd.DataFrame] = None):
        self._by = {}
        if df is not None and not df.empty:
            for _, r in df.iterrows():
                t = str(r.get("type", "")).strip().upper()
                cc = str(r.get("country_code", "")).strip().upper()
                key = str(r.get("key", "")).strip().upper()
                lat = r.get("lat"); lon = r.get("lon")
                try:
                    lat = float(lat); lon = float(lon)
                except Exception:
                    continue
                self._by.setdefault(t, {}).setdefault(cc, {})[key] = (lat, lon)

    @staticmethod
    def load_from_dir(base_dir: str) -> "GeoIndex":
        path = os.path.join(base_dir, GEO_FILE_NAME)
        if not os.path.exists(path):
            return GeoIndex()
        try:
            df = pd.read_excel(path, sheet_name=GEO_SHEET)
            return GeoIndex(df)
        except Exception:
            return GeoIndex()

    def lookup(self, typ: str, country_code: str, key: str) -> Optional[Tuple[float, float]]:
        t = str(typ).strip().upper()
        cc = (country_code or "").strip().upper()
        k = (key or "").strip().upper()
        return self._by.get(t, {}).get(cc, {}).get(k)

    def lookup_country(self, country_code: str) -> Optional[Tuple[float, float]]:
        cc = (country_code or "").strip().upper()
        # For country, the 'key' is the same as country_code
        return self._by.get("COUNTRY", {}).get(cc, {}).get(cc)


def normalize_zip(zip_code: str) -> str:
    """Normalize a ZIP/postal code into a simple alphanumeric uppercase token without spaces.
    If too short/long or clearly malformed, return empty string to trigger fallbacks.
    """
    if not zip_code:
        return ""
    z = str(zip_code).strip().upper().replace(" ", "").replace("-", "")
    # Keep only simple A-Z0-9 to avoid odd characters
    z = "".join([ch for ch in z if ch.isalnum()])
    if len(z) < 3 or len(z) > 10:
        return ""
    return z


def resolve_point(geo: GeoIndex,
                  country_code: str,
                  zip_code: Optional[str] = None,
                  city: Optional[str] = None,
                  plant: Optional[str] = None,
                  port: Optional[str] = None) -> Tuple[Optional[float], Optional[float], str]:
    """Resolve a location to coordinates using the optional local GEO index.

    Tries in order: ZIP -> CITY -> PLANT -> PORT -> COUNTRY centroid.
    Returns (lat, lon, source_tag)
    """
    # ZIP
    z = normalize_zip(zip_code or "")
    if z:
        p = geo.lookup("ZIP", country_code, z)
        if p:
            return p[0], p[1], f"zip:{z}"
        # pgeocode fallback by ZIP (offline)
        if pgeocode is not None:
            try:
                nomi = pgeocode.Nominatim((country_code or "").strip().upper())
                row = nomi.query_postal_code(z)
                if row is not None and pd.notna(row.latitude) and pd.notna(row.longitude):
                    return float(row.latitude), float(row.longitude), f"pgeocode_zip:{z}"
            except Exception:
                pass
    # City
    if city:
        c = str(city).strip().upper()
        p = geo.lookup("CITY", country_code, c)
        if p:
            return p[0], p[1], f"city:{c}"
        # pgeocode fallback by City centroid
        if pgeocode is not None:
            try:
                nomi = pgeocode.Nominatim((country_code or "").strip().upper())
                df = getattr(nomi, "_data", None)
                if df is not None and isinstance(df, pd.DataFrame) and not df.empty and "place_name" in df.columns:
                    exact = df[df["place_name"].astype(str).str.upper() == c]
                    subset = exact if not exact.empty else df[df["place_name"].astype(str).str.upper().str.contains(c, na=False)]
                    subset = subset.dropna(subset=["latitude", "longitude"])
                    if not subset.empty:
                        lat = float(pd.to_numeric(subset["latitude"]).mean())
                        lon = float(pd.to_numeric(subset["longitude"]).mean())
                        tag = "pgeocode_city_exact" if not exact.empty else "pgeocode_city_contains"
                        return lat, lon, f"{tag}:{c}"
            except Exception:
                pass
    # Plant
    if plant:
        pl = str(plant).strip().upper()
        p = geo.lookup("PLANT", country_code, pl)
        if p:
            return p[0], p[1], f"plant:{pl}"
    # Port
    if port:
        po = str(port).strip().upper()
        p = geo.lookup("PORT", country_code, po)
        if p:
            return p[0], p[1], f"port:{po}"
    # Country centroid
    p = geo.lookup_country(country_code)
    if p:
        return p[0], p[1], f"country:{country_code}"
    return None, None, ""


def road_km_between(p1: Tuple[float, float], p2: Tuple[float, float], road_factor: float = 1.30) -> float:
    """Approximate road distance using haversine * road_factor (default 1.30)."""
    lat1, lon1 = p1
    lat2, lon2 = p2
    geo_km = haversine_km(lat1, lon1, lat2, lon2)
    return geo_km * float(road_factor)


# ---------------------------------------------------------------------------
# Truck routing (HERE Routing API v8) + Excel template / batch processing
# ---------------------------------------------------------------------------
HERE_ROUTING_URL = "https://router.hereapi.com/v8/routes"
ORS_ROUTING_URL = "https://api.openrouteservice.org/v2/directions/driving-hgv"
TEMPLATE_FILE_NAME = "TRUCK_ROUTES_TEMPLATE.xlsx"
ROUTES_SHEET = "ROUTES"

# Default megatrailer profile (tractor + semitrailer, 5 axles, 40 t legal limit)
DEFAULT_TARE_KG = 16000
DEFAULT_HEIGHT_CM = 400
DEFAULT_WIDTH_CM = 255
DEFAULT_LENGTH_CM = 1650
DEFAULT_AXLES = 5

INPUT_COLUMNS = [
    "ID",
    "ORIGIN_COUNTRY", "ORIGIN_ZIP", "ORIGIN_LAT", "ORIGIN_LON",
    "DEST_COUNTRY", "DEST_ZIP", "DEST_LAT", "DEST_LON",
    "CARGO_WEIGHT_KG", "TRUCK_TARE_KG", "GROSS_WEIGHT_KG",
    "HEIGHT_CM", "WIDTH_CM", "LENGTH_CM", "AXLE_COUNT",
]
OUTPUT_COLUMNS = [
    "ORIGIN_SOURCE", "DEST_SOURCE",
    "ORIGIN_LAT_USED", "ORIGIN_LON_USED", "DEST_LAT_USED", "DEST_LON_USED",
    "ROUTE_KM", "DURATION_H", "STRAIGHT_KM", "METHOD", "STATUS",
]


def _num(value, default=None):
    """Return a float from a cell value or `default` when empty/invalid."""
    try:
        if value is None or pd.isna(value) or str(value).strip() == "":
            return default
        return float(value)
    except (TypeError, ValueError):
        return default


def _text(value) -> str:
    """Return a clean string from a cell value ('' when empty/NaN)."""
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return ""
    return str(value).strip()


def resolve_location(geo: GeoIndex, row, prefix: str) -> Tuple[float, float, str]:
    """Resolve ORIGIN/DEST using coordinates and/or ZIP, whichever is available.

    Coordinates are used when present and valid; otherwise country + ZIP is used.
    Returns (lat, lon, source) where source is 'COORDS' or 'ZIP'.
    """
    lat, lon = _num(row[f"{prefix}_LAT"]), _num(row[f"{prefix}_LON"])
    if lat is not None and lon is not None and -90 <= lat <= 90 and -180 <= lon <= 180:
        return lat, lon, "COORDS"

    zip_code = _text(row[f"{prefix}_ZIP"])
    country = _text(row[f"{prefix}_COUNTRY"])
    if zip_code and country:
        zlat, zlon, _ = resolve_point(geo, country, zip_code)
        if zlat is not None and zlon is not None:
            return zlat, zlon, "ZIP"
        raise ValueError(f"{prefix}: ZIP '{zip_code}' ({country}) not found")
    if zip_code and not country:
        raise ValueError(f"{prefix}: ZIP given without COUNTRY")
    raise ValueError(f"{prefix}: provide LAT+LON or COUNTRY+ZIP")


def here_truck_route(origin: Tuple[float, float],
                     destination: Tuple[float, float],
                     gross_weight_kg: float,
                     height_cm: float = DEFAULT_HEIGHT_CM,
                     width_cm: float = DEFAULT_WIDTH_CM,
                     length_cm: float = DEFAULT_LENGTH_CM,
                     axle_count: int = DEFAULT_AXLES,
                     api_key: Optional[str] = None,
                     timeout: int = 30) -> Tuple[float, float]:
    """Truck-restricted road route. Returns (distance_km, duration_hours).

    Raises RuntimeError if no key is configured or the API call fails.
    """
    import requests

    key = api_key or os.environ.get("HERE_API_KEY")
    if not key:
        raise RuntimeError("HERE_API_KEY not set")
    params = {
        "origin": f"{origin[0]},{origin[1]}",
        "destination": f"{destination[0]},{destination[1]}",
        "transportMode": "truck",
        "return": "summary",
        "vehicle[grossWeight]": int(round(gross_weight_kg)),
        "vehicle[height]": int(round(height_cm)),
        "vehicle[width]": int(round(width_cm)),
        "vehicle[length]": int(round(length_cm)),
        "vehicle[axleCount]": int(axle_count),
        "apiKey": key,
    }
    resp = requests.get(HERE_ROUTING_URL, params=params, timeout=timeout)
    if resp.status_code != 200:
        raise RuntimeError(f"HERE error {resp.status_code}: {resp.text[:200]}")
    routes = resp.json().get("routes") or []
    if not routes:
        raise RuntimeError("HERE returned no route")
    sections = routes[0]["sections"]
    meters = sum(s["summary"]["length"] for s in sections)
    seconds = sum(s["summary"]["duration"] for s in sections)
    return meters / 1000.0, seconds / 3600.0


def ors_truck_route(origin: Tuple[float, float],
                    destination: Tuple[float, float],
                    gross_weight_kg: float,
                    height_cm: float = DEFAULT_HEIGHT_CM,
                    width_cm: float = DEFAULT_WIDTH_CM,
                    length_cm: float = DEFAULT_LENGTH_CM,
                    axle_count: int = DEFAULT_AXLES,
                    api_key: Optional[str] = None,
                    timeout: int = 60) -> Tuple[float, float]:
    """Truck route via OpenRouteService (free key, profile driving-hgv).

    Units expected by ORS: tonnes and metres. Coordinates are [lon, lat].
    Returns (distance_km, duration_hours).
    """
    import requests

    key = api_key or os.environ.get("ORS_API_KEY")
    if not key:
        raise RuntimeError("ORS_API_KEY not set")
    body = {
        "coordinates": [[origin[1], origin[0]], [destination[1], destination[0]]],
        "radiuses": [5000, 5000],
        "options": {"profile_params": {"restrictions": {
            "weight": round(gross_weight_kg / 1000.0, 2),
            "height": round(height_cm / 100.0, 2),
            "width": round(width_cm / 100.0, 2),
            "length": round(length_cm / 100.0, 2),
        }}},
    }
    resp = requests.post(ORS_ROUTING_URL, json=body, timeout=timeout,
                         headers={"Authorization": key, "Content-Type": "application/json"})
    if resp.status_code != 200:
        raise RuntimeError(f"ORS error {resp.status_code}: {resp.text[:200]}")
    routes = resp.json().get("routes") or []
    if not routes:
        raise RuntimeError("ORS returned no route")
    summary = routes[0]["summary"]
    return summary["distance"] / 1000.0, summary["duration"] / 3600.0


def create_template(path: str) -> str:
    """Write the Excel template (inputs sheet with example rows + instructions)."""
    example = pd.DataFrame([
        {"ID": "EX1", "ORIGIN_COUNTRY": "IT", "ORIGIN_ZIP": "", "ORIGIN_LAT": 41.08256, "ORIGIN_LON": 16.780552,
         "DEST_COUNTRY": "FR", "DEST_ZIP": "", "DEST_LAT": 49.27653, "DEST_LON": 4.05909,
         "CARGO_WEIGHT_KG": 24000, "TRUCK_TARE_KG": 16000, "GROSS_WEIGHT_KG": None,
         "HEIGHT_CM": 400, "WIDTH_CM": 255, "LENGTH_CM": 1650, "AXLE_COUNT": 5},
        {"ID": "EX2", "ORIGIN_COUNTRY": "ES", "ORIGIN_ZIP": "28001", "ORIGIN_LAT": None, "ORIGIN_LON": None,
         "DEST_COUNTRY": "FR", "DEST_ZIP": "75001", "DEST_LAT": None, "DEST_LON": None,
         "CARGO_WEIGHT_KG": 24000, "TRUCK_TARE_KG": None, "GROSS_WEIGHT_KG": None,
         "HEIGHT_CM": None, "WIDTH_CM": None, "LENGTH_CM": None, "AXLE_COUNT": None},
    ], columns=INPUT_COLUMNS)
    notes = pd.DataFrame({"INSTRUCCIONES": [
        "Origen/destino: rellena LAT+LON, o COUNTRY (ISO2) + ZIP, o ambos. Cada punto (origen y destino) puede usar un metodo distinto.",
        "Si hay LAT+LON validas se usan estas (mas precisas); si no, se usa COUNTRY+ZIP. Las columnas ORIGIN_SOURCE/DEST_SOURCE indican cual se uso.",
        "Con solo coordenadas, COUNTRY y ZIP pueden quedar vacios; con solo ZIP, COUNTRY es obligatorio.",
        "Las coordenadas se escriben en grados decimales (latitud, longitud).",
        "GROSS_WEIGHT_KG = peso total del conjunto cargado. Si esta vacio se calcula CARGO_WEIGHT_KG + TRUCK_TARE_KG.",
        f"Valores vacios por defecto: TARE={DEFAULT_TARE_KG} kg, ALTURA={DEFAULT_HEIGHT_CM} cm, "
        f"ANCHO={DEFAULT_WIDTH_CM} cm, LARGO={DEFAULT_LENGTH_CM} cm, EJES={DEFAULT_AXLES}.",
        "Dimensiones en cm, pesos en kg.",
        "Rutas reales de camion: define ORS_API_KEY (OpenRouteService, gratis) o HERE_API_KEY; sin ninguna se usa una aproximacion.",
        "Las columnas de resultado (ROUTE_KM, DURATION_H, METHOD, STATUS...) se anaden al procesar.",
    ]})
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        example.to_excel(xw, sheet_name=ROUTES_SHEET, index=False)
        notes.to_excel(xw, sheet_name="INSTRUCCIONES", index=False)
        for ws in xw.book.worksheets:
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = max(14, len(str(col[0].value)) + 2)
    return path


def process_excel(input_path: str, output_path: Optional[str] = None,
                  road_factor: float = 1.30, sleep_s: float = 0.0) -> pd.DataFrame:
    """Read the template, compute truck distances per row and save the result."""
    import time

    base_dir = os.path.dirname(os.path.abspath(input_path))
    geo = GeoIndex.load_from_dir(base_dir)
    df = pd.read_excel(input_path, sheet_name=ROUTES_SHEET, dtype={"ORIGIN_ZIP": str, "DEST_ZIP": str})
    for col in INPUT_COLUMNS:
        if col not in df.columns:
            df[col] = None
    for col in OUTPUT_COLUMNS:
        df[col] = None
    # Free option first (OpenRouteService), then HERE; none -> approximation
    if os.environ.get("ORS_API_KEY"):
        provider = ("ORS_HGV", ors_truck_route)
    elif os.environ.get("HERE_API_KEY"):
        provider = ("HERE_TRUCK", here_truck_route)
    else:
        provider = None

    for i, r in df.iterrows():
        try:
            olat, olon, osrc = resolve_location(geo, r, "ORIGIN")
            dlat, dlon, dsrc = resolve_location(geo, r, "DEST")
            o, d = (olat, olon), (dlat, dlon)
            df.loc[i, "ORIGIN_SOURCE"] = osrc
            df.loc[i, "DEST_SOURCE"] = dsrc
            df.loc[i, ["ORIGIN_LAT_USED", "ORIGIN_LON_USED"]] = o
            df.loc[i, ["DEST_LAT_USED", "DEST_LON_USED"]] = d
            straight = haversine_km(*o, *d)
            df.loc[i, "STRAIGHT_KM"] = round(straight, 1)

            gross = _num(r["GROSS_WEIGHT_KG"])
            if gross is None:
                cargo = _num(r["CARGO_WEIGHT_KG"], 0.0)
                gross = cargo + _num(r["TRUCK_TARE_KG"], DEFAULT_TARE_KG)

            if provider:
                km, hours = provider[1](
                    o, d, gross,
                    height_cm=_num(r["HEIGHT_CM"], DEFAULT_HEIGHT_CM),
                    width_cm=_num(r["WIDTH_CM"], DEFAULT_WIDTH_CM),
                    length_cm=_num(r["LENGTH_CM"], DEFAULT_LENGTH_CM),
                    axle_count=int(_num(r["AXLE_COUNT"], DEFAULT_AXLES)),
                )
                df.loc[i, "ROUTE_KM"] = round(km, 1)
                df.loc[i, "DURATION_H"] = round(hours, 2)
                df.loc[i, "METHOD"] = provider[0]
            else:
                df.loc[i, "ROUTE_KM"] = round(straight * road_factor, 1)
                df.loc[i, "METHOD"] = f"APPROX_HAVERSINE_x{road_factor}"
            df.loc[i, "STATUS"] = "OK"
        except Exception as exc:
            df.loc[i, "STATUS"] = f"ERROR: {exc}"
        if sleep_s:
            time.sleep(sleep_s)

    out = output_path or os.path.splitext(input_path)[0] + "_RESULT.xlsx"
    df.to_excel(out, sheet_name=ROUTES_SHEET, index=False)
    return df


if __name__ == "__main__":
    import sys

    here = os.path.dirname(os.path.abspath(__file__))
    if len(sys.argv) > 1 and sys.argv[1] == "template":
        print("Template created:", create_template(os.path.join(here, TEMPLATE_FILE_NAME)))
    elif len(sys.argv) > 1:
        result = process_excel(sys.argv[1])
        print(result[["ID", "ROUTE_KM", "DURATION_H", "METHOD", "STATUS"]].to_string(index=False))
    else:
        print("Usage: python Distances.py template | python Distances.py <input.xlsx>")


