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


