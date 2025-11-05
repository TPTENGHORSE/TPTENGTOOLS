import pandas as pd
from Quotations.generate_quote import QTOOL_DIR, find_qtool_data_file
from Quotations.Distances import GeoIndex, resolve_point, road_km_between


def port_coords_from_ports_locations(port_code: str) -> tuple[float | None, float | None]:
    data_file = find_qtool_data_file()
    if not data_file:
        return None, None
    try:
        df_ports = pd.read_excel(data_file, sheet_name="Ports Locations")
    except Exception:
        return None, None
    try:
        m = df_ports[df_ports["POL/POD"].astype(str).str.upper() == str(port_code).strip().upper()]
        if not m.empty:
            lat = m.iloc[0].get("LAT"); lon = m.iloc[0].get("LONG")
            if pd.notna(lat) and pd.notna(lon):
                return float(lat), float(lon)
    except Exception:
        return None, None
    return None, None


def build_geo_index() -> GeoIndex:
    """Build a GeoIndex including PLANT (from HORSE-PUERTO) and PORT (from Ports Locations),
    mirroring the generator's construction."""
    data_file = find_qtool_data_file()
    geo_rows = []
    if data_file:
        # Plants
        try:
            df_hp = pd.read_excel(data_file, sheet_name="HORSE-PUERTO")
            for _, rr in df_hp.iterrows():
                cc = str(rr.get("Country Code", "")).strip().upper()
                plant = str(rr.get("Plant", "")).strip().upper()
                lat = rr.get("Plant Lat"); lon = rr.get("Plant Long")
                if cc and plant and pd.notna(lat) and pd.notna(lon):
                    geo_rows.append({
                        "type": "PLANT",
                        "country_code": cc,
                        "key": plant,
                        "lat": float(lat),
                        "lon": float(lon),
                    })
        except Exception:
            pass
        # Ports
        try:
            df_ports = pd.read_excel(data_file, sheet_name="Ports Locations")
            def _coerce_coord(val, kind: str):
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
            for _, rr in df_ports.iterrows():
                code = str(rr.get("POL/POD", "")).strip().upper()
                lat = _coerce_coord(rr.get("LAT"), "lat")
                lon = _coerce_coord(rr.get("LONG"), "lon")
                # Country may be name; use UN/LOC prefix from code when possible
                cc = code[:2] if len(code) >= 2 else str(rr.get("Country", "")).strip().upper()
                if code and lat is not None and lon is not None:
                    geo_rows.append({
                        "type": "PORT",
                        "country_code": cc,
                        "key": code,
                        "lat": float(lat),
                        "lon": float(lon),
                    })
        except Exception:
            pass
    if geo_rows:
        return GeoIndex(pd.DataFrame(geo_rows))
    return GeoIndex.load_from_dir(QTOOL_DIR)


def compute_leg1_km(origin_cc: str, origin_city: str, origin_zip: str, pol_code: str, plant: str | None = None):
    geo = build_geo_index()
    # Resolve origin: try ZIP then City
    olat, olon, osrc = resolve_point(geo, origin_cc, zip_code=origin_zip, city=origin_city)
    if olat is None and plant:
        olat, olon, osrc = resolve_point(geo, origin_cc, plant=plant)
    # Resolve POL via index, fallback to Ports Locations
    plat, plon, psrc = resolve_point(geo, origin_cc, port=pol_code)
    if plat is None:
        f_lat, f_lon = port_coords_from_ports_locations(pol_code)
        if f_lat is not None:
            plat, plon, psrc = f_lat, f_lon, "ports_locations"
    if olat is None or plat is None:
        return None, osrc, psrc
    # Debug coordinates
    print(f"Origen coords: ({olat:.6f}, {olon:.6f}) from {osrc}")
    print(f"POL coords:    ({plat:.6f}, {plon:.6f}) from {psrc}")
    km = road_km_between((olat, olon), (plat, plon))
    # Also compute swapped in case source LAT/LONG are reversed
    km_swapped = road_km_between((olat, olon), (plon, plat))
    # Return both via source tag hint
    if km_swapped < km:
        return km_swapped, osrc, psrc + "(swapped)"
    return km, osrc, psrc


if __name__ == "__main__":
    # Case: ES Valladolid 47008 -> POL ESVLC
    km, osrc, psrc = compute_leg1_km("ES", "Valladolid", "47008", "ESVLC", plant="HORSE MOTORES")
    if km is None:
        print(f"No se pudo calcular: origen={osrc} pol={psrc}")
    else:
        print(f"Leg1 km: {km:.2f}  (origen={osrc} -> pol={psrc})")
