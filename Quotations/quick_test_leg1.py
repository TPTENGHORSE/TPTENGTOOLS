import os
import pandas as pd
from Quotations.generate_quote import build_output, next_output_path, QTOOL_DIR

# Single test row matching the user's example
row = {
    "pn": "TEST-PN",
    "designation": "Test Item",
    "supplier_plant": "",  # not used here; we'll rely on city/zip
    "incoterm": "FCA",
    "origin_country_code": "IN",
    "origin_country": "India",
    "origin_city": "Mundhwa",
    "origin_zip": "34190",
    "dest_plant": "Horse Valladolid",
    "dest_country_code": "ES",
    "dest_country": "Spain",
    "dest_city": "Valladolid",
    "dest_zip": "47008",
}

df = pd.DataFrame([row])

out_path = next_output_path(QTOOL_DIR)
build_output(df, out_path)

# Read back result to print Leg1 Distance (km)
q = pd.read_excel(out_path, sheet_name="Quote")
cols = list(q.columns)
leg1_col = next((c for c in cols if c.strip().lower() == "leg1 distance (km)".lower()), None)
pol_col = "POL" if "POL" in cols else None
pod_col = "POD" if "POD" in cols else None
if leg1_col:
    leg1_km = q.iloc[0][leg1_col]
    pol = q.iloc[0][pol_col] if pol_col else None
    pod = q.iloc[0][pod_col] if pod_col else None
    print(f"Output file: {out_path}")
    print(f"POL={pol} POD={pod} Leg1 Distance (km)={leg1_km}")
else:
    print(f"Output file: {out_path}")
    print("Leg1 Distance (km) column not found")
