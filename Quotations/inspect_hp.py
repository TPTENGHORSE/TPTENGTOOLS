import pandas as pd
from Quotations.generate_quote import find_qtool_data_file

def main():
    path = find_qtool_data_file()
    if not path:
        print("No data file found")
        return
    try:
        df = pd.read_excel(path, sheet_name="HORSE-PUERTO")
    except Exception as e:
        print(f"Error reading HORSE-PUERTO: {e}")
        return
    # Show ES plants and any that look like Horse Motores
    mask_es = df.get("Country Code", pd.Series()).astype(str).str.upper().eq("ES")
    mask_name = df.get("Plant", pd.Series()).astype(str).str.upper().str.contains("HORSE|VALLADOLID|MOTO", regex=True)
    view = df[mask_es | mask_name].copy()
    cols = [c for c in ["Country Code","Country","Plant","Plant Lat","Plant Long"] if c in view.columns]
    print(view[cols].head(50).to_string(index=False))

if __name__ == "__main__":
    main()
