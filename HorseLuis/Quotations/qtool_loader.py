import os
import sys
import pandas as pd


STD_COLS = {
    "Part Number (PN)": "pn",
    "Part Designation": "designation",
    "Supplier/Plant": "supplier_plant",
    "Incoterm": "incoterm",
    "Origin Country code": "origin_country_code",
    "Origin Country": "origin_country",
    "Origin City": "origin_city",
    "Origin ZIP Code": "origin_zip",
    "Destination Plant": "dest_plant",
    "Destination country code": "dest_country_code",
    "Destination Country": "dest_country",
    "Destination City": "dest_city",
    "Destinartion ZIP Code": "dest_zip",
    "Anual Needs (PN / Year)": "annual_needs",
    "Daily Need (PN / Day)": "daily_need",
    "PN Unit cost (€)": "unit_cost_eur",
    "Packaging Code": "packaging_code",
}


def load_input_template(path: str, sheet: str = "Input") -> pd.DataFrame:
    if not os.path.exists(path):
        raise FileNotFoundError(path)
    df_full = pd.read_excel(path, sheet_name=sheet)
    # Keep only known columns, rename to standard
    keep = [c for c in df_full.columns if c in STD_COLS]
    df = df_full[keep].rename(columns=STD_COLS)
    # Fallback: if 'incoterm' missing (header mismatch), try case-insensitive or column M (index 12)
    if "incoterm" not in df.columns:
        # 1) Try case-insensitive name match in original
        incoterm_col = None
        for col in df_full.columns:
            if str(col).strip().lower() == "incoterm":
                incoterm_col = col
                break
        # 2) Fallback to column M (0-based index 12) if available
        if incoterm_col is None and df_full.shape[1] > 12:
            incoterm_series = df_full.iloc[:, 12]
        elif incoterm_col is not None:
            incoterm_series = df_full[incoterm_col]
        else:
            incoterm_series = None
        if incoterm_series is not None:
            df["incoterm"] = incoterm_series.astype(str).str.strip().str.upper()
    # Trim strings
    for c in [
        "pn","designation","supplier_plant","incoterm","origin_country_code","origin_country","origin_city","origin_zip",
        "dest_plant","dest_country_code","dest_country","dest_city","dest_zip","packaging_code"
    ]:
        if c in df.columns:
            df[c] = df[c].astype(str).str.strip()
    # Normalize incoterm to uppercase if present (ensure consistency)
    if "incoterm" in df.columns:
        df["incoterm"] = df["incoterm"].astype(str).str.strip().str.upper()
    # Numeric coercions
    for c in ["annual_needs","daily_need","unit_cost_eur"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    # Drop fully empty rows
    df = df.dropna(how="all")
    return df


def main():
    if len(sys.argv) < 2:
        print("Usage: python Quotations/qtool_loader.py <path-to-upload_Quotation Template.xlsx>")
        sys.exit(2)
    path = sys.argv[1]
    df = load_input_template(path)
    out_dir = os.path.join(os.path.dirname(__file__), "_out")
    os.makedirs(out_dir, exist_ok=True)
    out_csv = os.path.join(out_dir, "normalized_input.csv")
    df.to_csv(out_csv, index=False, encoding="utf-8-sig")
    print(f"Normalized input written: {out_csv}")


if __name__ == "__main__":
    main()
