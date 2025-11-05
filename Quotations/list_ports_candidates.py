import os
import sys
import pandas as pd

# Reuse finder and constants from generate_quote
try:
    from .generate_quote import find_qtool_data_file
except Exception:
    sys.path.append(os.path.dirname(os.path.dirname(__file__)))
    from Quotations.generate_quote import find_qtool_data_file  # type: ignore

def main(country_code: str = "IN"):
    data_file = find_qtool_data_file()
    if not data_file:
        print("No se encontró QUOTATION TOOL DATA.xlsx")
        return
    try:
        df_ports = pd.read_excel(data_file, sheet_name="Ports Locations")
    except Exception as e:
        print(f"No se pudo leer 'Ports Locations': {e}")
        return
    if df_ports is None or df_ports.empty:
        print("Ports Locations vacío")
        return
    cc = (country_code or "").strip().upper()
    cand = []
    seen = set()
    for _, rr in df_ports.iterrows():
        try:
            code = str(rr.get("POL/POD", "")).strip().upper()
            if not code:
                continue
            ctry = str(rr.get("Country", "")).strip().upper()
            # match by Country value or by UN/LOCODE prefix
            ok = (ctry == cc or ctry == "INDIA" or ctry == "IND") or (code.startswith(cc))
            if ok:
                if code not in seen:
                    seen.add(code)
                    cand.append(code)
        except Exception:
            continue
    print(f"Candidatos POL para {cc}: {len(cand)}")
    if cand:
        print("Lista:", ", ".join(sorted(cand)))

if __name__ == "__main__":
    cc = sys.argv[1] if len(sys.argv) > 1 else "IN"
    main(cc)
