import pandas as pd
from Quotations.generate_quote import find_qtool_data_file


def main():
    path = find_qtool_data_file()
    if not path:
        print("No QUOTATION TOOL DATA found")
        return
    xls = pd.ExcelFile(path)
    print("Sheets:", xls.sheet_names)
    for name in xls.sheet_names:
        try:
            df = xls.parse(name, nrows=3)
            print(f"\nSheet: {name}")
            print("Columns:", list(df.columns))
            print(df.head(3).to_string(index=False))
        except Exception as e:
            print(f"Error reading {name}: {e}")


if __name__ == "__main__":
    main()
