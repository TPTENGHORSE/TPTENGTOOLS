import pandas as pd
from Quotations.generate_quote import QTOOL_DIR, INPUT_FILE, find_qtool_data_file
from Quotations.Distances import GeoIndex, resolve_point, road_km_between


def coerce_coord(val, kind: str) -> float | None:
    try:
        if pd.isna(val):
            return None
        if isinstance(val, str):
            s = val.strip().replace(" ", "").replace(",", ".")
            f = float(s)
        else:
            f = float(val)
        if kind == "lat" and abs(f) > 90 and abs(f) <= 180000:
            f = f / 1000.0
        if kind == "lon" and abs(f) > 180 and abs(f) <= 360000:
            f = f / 1000.0
        if kind == "lat" and abs(f) <= 90:
            return f
        if kind == "lon" and abs(f) <= 180:
            return f
        return None
    except Exception:
        return None


def build_geo_index() -> GeoIndex:
    data_file = find_qtool_data_file()
    geo_rows = []
    if data_file:
        try:
            df_hp = pd.read_excel(data_file, sheet_name="HORSE-PUERTO")
            for _, rr in df_hp.iterrows():
                cc = str(rr.get("Country Code", "")).strip().upper()
                plant = str(rr.get("Plant", "")).strip().upper()
                lat = rr.get("Plant Lat"); lon = rr.get("Plant Long")
                if cc and plant and pd.notna(lat) and pd.notna(lon):
                    geo_rows.append({"type":"PLANT","country_code":cc,"key":plant,"lat":float(lat),"lon":float(lon)})
            # Also index ZIP and City from HP if available
            if "Plant ZIP Code" in df_hp.columns:
                for _, rr in df_hp.iterrows():
                    cc = str(rr.get("Country Code", "")).strip().upper()
                    z = str(rr.get("Plant ZIP Code", "")).strip().upper()
                    lat = rr.get("Plant Lat"); lon = rr.get("Plant Long")
                    if cc and z and pd.notna(lat) and pd.notna(lon):
                        geo_rows.append({"type":"ZIP","country_code":cc,"key":z,"lat":float(lat),"lon":float(lon)})
            if "Plant City" in df_hp.columns:
                for _, rr in df_hp.iterrows():
                    cc = str(rr.get("Country Code", "")).strip().upper()
                    c = str(rr.get("Plant City", "")).strip().upper()
                    lat = rr.get("Plant Lat"); lon = rr.get("Plant Long")
                    if cc and c and pd.notna(lat) and pd.notna(lon):
                        geo_rows.append({"type":"CITY","country_code":cc,"key":c,"lat":float(lat),"lon":float(lon)})
        except Exception:
            pass
        try:
            df_ports = pd.read_excel(data_file, sheet_name="Ports Locations")
            for _, rr in df_ports.iterrows():
                code = str(rr.get("POL/POD", "")).strip().upper()
                lat = coerce_coord(rr.get("LAT"), "lat")
                lon = coerce_coord(rr.get("LONG"), "lon")
                cc = code[:2] if len(code) >= 2 else str(rr.get("Country", "")).strip().upper()
                if code and lat is not None and lon is not None:
                    geo_rows.append({"type":"PORT","country_code":cc,"key":code,"lat":float(lat),"lon":float(lon)})
        except Exception:
            pass
    if geo_rows:
        return GeoIndex(pd.DataFrame(geo_rows))
    return GeoIndex.load_from_dir(QTOOL_DIR)


def parse_city_zip(city_val: str, zip_val: str) -> tuple[str, str]:
    import re
    city_raw = (city_val or "").strip()
    zip_raw = (zip_val or "").strip()
    if city_raw and not zip_raw:
        m = re.findall(r"(\d{4,6})", city_raw)
        if m:
            zip_raw = m[-1]
            city_raw = re.sub(r"[\s,-]*" + re.escape(zip_raw) + r"\b", "", city_raw).strip()
    return city_raw, zip_raw


def port_coords_from_ports_locations(port_code: str) -> tuple[float | None, float | None]:
    data_file = find_qtool_data_file()
    if not data_file:
        return None, None
    try:
        df_ports = pd.read_excel(data_file, sheet_name="Ports Locations")
        m = df_ports[df_ports["POL/POD"].astype(str).str.upper() == str(port_code).strip().upper()]
        if not m.empty:
            lat = coerce_coord(m.iloc[0].get("LAT"), "lat")
            lon = coerce_coord(m.iloc[0].get("LONG"), "lon")
            if lat is not None and lon is not None:
                return lat, lon
    except Exception:
        return None, None
    return None, None


def compute_leg1_for_mundhwa():
    # Load input template
    df = pd.read_excel(INPUT_FILE, sheet_name="Input")
    # Column names per template
    oc_code = "Origin Country code"
    oc_name = "Origin Country"
    city_col = "Origin City"
    zip_col = "Origin ZIP Code"
    # Filter Mundhwa rows in India
    mask = (
        df.get(oc_code, pd.Series()).astype(str).str.upper().eq("IN") |
        df.get(oc_name, pd.Series()).astype(str).str.upper().eq("INDIA")
    ) & df.get(city_col, pd.Series()).astype(str).str.upper().str.contains("MUNDHWA", na=False)
    sel = df[mask].copy()
    if sel.empty:
        print("No hay filas con India / Mundhwa en el template.")
        return
    geo = build_geo_index()
    POL = "INENR"
    # Port coords via index then ports locations
    plat, plon, psrc = resolve_point(geo, "IN", port=POL)
    if plat is None:
        f_lat, f_lon = port_coords_from_ports_locations(POL)
        if f_lat is not None:
            plat, plon, psrc = f_lat, f_lon, "ports_locations"
    if plat is None:
        print("No se pudieron resolver coords de POL INENR")
        return
    for i, row in sel.iterrows():
        oc = str(row.get(oc_code, "")).strip().upper() or "IN"
        city = str(row.get(city_col, ""))
        zipc = str(row.get(zip_col, ""))
        city_clean, zip_clean = parse_city_zip(city, zipc)
        # Resolve origin
        olat = olon = None
        osrc = ""
        if zip_clean:
            lat, lon, src = resolve_point(geo, oc, zip_code=zip_clean)
            if lat is not None and src.startswith("zip:"):
                olat, olon, osrc = lat, lon, src
        if olat is None and city_clean:
            lat, lon, src = resolve_point(geo, oc, city=city_clean)
            if lat is not None and src.startswith("city:"):
                olat, olon, osrc = lat, lon, src
        # Alias fallback: Mundhwa -> Pune for India
        if olat is None and city_clean and oc == "IN":
            alias = None
            if city_clean.strip().upper() == "MUNDHWA":
                alias = "PUNE"
            if alias:
                lat, lon, src = resolve_point(geo, oc, city=alias)
                if lat is not None and src.startswith("city:"):
                    olat, olon, osrc = lat, lon, src + "(alias)"
        # Known city coordinates (last resort)
        if olat is None and oc == "IN":
            if city_clean.strip().upper() == "MUNDHWA":
                olat, olon, osrc = 18.5366, 73.9152, "known_city"
            elif city_clean.strip().upper() == "PUNE":
                olat, olon, osrc = 18.5204, 73.8567, "known_city"
        if olat is None:
            # Try HP city/zip already indexed into GEO
            if zip_clean:
                lat, lon, src = resolve_point(geo, oc, zip_code=zip_clean)
                if lat is not None and src.startswith("zip:"):
                    olat, olon, osrc = lat, lon, src
            if olat is None and city_clean:
                lat, lon, src = resolve_point(geo, oc, city=city_clean)
                if lat is not None and src.startswith("city:"):
                    olat, olon, osrc = lat, lon, src
        if olat is None:
            print(f"Fila {i}: origen no resuelto (ciudad='{city}', zip='{zipc}')")
            continue
        km = road_km_between((olat, olon), (plat, plon))
        print(f"Fila {i}: Origen({osrc})=({olat:.6f},{olon:.6f}) → POL({psrc})=({plat:.6f},{plon:.6f})  Leg1={km:.2f} km")


if __name__ == "__main__":
    compute_leg1_for_mundhwa()
