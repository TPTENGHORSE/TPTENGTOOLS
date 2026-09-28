"""Apply review fixes to the 2027 airfreight quotation template.

Reads 'Airfreights quotation template_2027.xlsx' from the ongoing tender
folder and writes a corrected copy (never overwrites the original) with:
  1. Input columns (V:AX on 'AIR') unlocked + sheet protection enabled.
  2. Dropdown validation for 'Type of Flow' (D) and 'Service Level' (T).
  3. Sensible freeze panes (header rows + reference columns visible).
  4. Missing top-level category labels (AF2, AP2).
  5. Merged duplicate sub-headers (Z3:AA3, AC3:AD3, AW3:AX3).
  6. Number formats on the pricing/count input columns.
  7. Whole-number validation on 'Nº Transshipment' (AQ), matching AR.
  8. Fixed the 'Sea-Air' row 25 formula to match the rest of the column.
"""

from pathlib import Path

from openpyxl import load_workbook
from openpyxl.worksheet.datavalidation import DataValidation

SOURCE_PATH = Path(
    r"C:\Users\OLMEDOGALVEZJorgeLui\OneDrive - Horse\Exchange VRAC\02_Engineering Department"
    r"\04. Tender Process\05. ONGOING TENDERS\2027\Airfreights_All_OVS_Flows"
    r"\Airfreights quotation template_2027.xlsx"
)
OUTPUT_PATH = SOURCE_PATH.with_name("Airfreights quotation template_2027_CORRECTED.xlsx")

FIRST_DATA_ROW = 5
LAST_DATA_ROW = 235

# Rate/count columns on 'AIR' that suppliers fill in and that were left as "General".
CURRENCY_COLUMNS = ["V", "W", "X", "Z", "AA", "AB", "AC", "AD", "AE", "AF"] + [
    "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN"
] + ["AS", "AT", "AU", "AW", "AX"]
INTEGER_COLUMNS = ["U", "Y", "AO", "AQ", "AR", "AV"]


def fix_air_sheet(ws) -> None:
    # 1) Unlock the supplier input columns, lock everything else, then protect the sheet.
    for row in ws.iter_rows(min_row=FIRST_DATA_ROW, max_row=LAST_DATA_ROW, min_col=1, max_col=50):
        for cell in row:
            col_letter = cell.column_letter
            cell.protection = cell.protection.copy(
                locked=col_letter not in CURRENCY_COLUMNS and col_letter not in INTEGER_COLUMNS
            )
    ws.protection.sheet = True
    ws.protection.formatCells = False
    ws.protection.formatColumns = False
    ws.protection.formatRows = False
    ws.protection.selectLockedCells = True
    ws.protection.selectUnlockedCells = True

    # 2) Dropdown validation for Type of Flow (D) and Service Level (T).
    flow_dv = DataValidation(type="list", formula1='"Main,Backup"', allow_blank=False, showErrorMessage=True)
    flow_dv.error = "Select Main or Backup."
    flow_dv.errorTitle = "Invalid Type of Flow"
    ws.add_data_validation(flow_dv)
    flow_dv.add(f"D{FIRST_DATA_ROW}:D{LAST_DATA_ROW}")

    service_dv = DataValidation(
        type="list", formula1='"Economy,Normal,Express"', allow_blank=False, showErrorMessage=True
    )
    service_dv.error = "Select Economy, Normal or Express."
    service_dv.errorTitle = "Invalid Service Level"
    ws.add_data_validation(service_dv)
    service_dv.add(f"T{FIRST_DATA_ROW}:T{LAST_DATA_ROW}")

    # 3) Freeze the 4 header rows + reference columns instead of the stray 'D110'.
    ws.freeze_panes = "E5"

    # 4) Fill in the missing top-level category labels.
    ws["AF2"] = "Air Freight - Euros €"
    ws["AP2"] = "Carrier & Transit Info"

    # 5) Merge duplicate sub-headers that are currently two separate identical cells.
    for rng, keep in (("Z3:AA3", "Z3"), ("AC3:AD3", "AC3"), ("AW3:AX3", "AW3")):
        second_cell = rng.split(":")[1]
        ws[second_cell] = None
        ws.merge_cells(rng)
        ws[keep].alignment = ws[keep].alignment.copy(horizontal="center")

    # 6) Number formats for rate/fee columns (2 decimals) and count columns (integers).
    for col_letter in CURRENCY_COLUMNS:
        for row in range(FIRST_DATA_ROW, LAST_DATA_ROW + 1):
            ws[f"{col_letter}{row}"].number_format = "#,##0.00"
    for col_letter in INTEGER_COLUMNS:
        for row in range(FIRST_DATA_ROW, LAST_DATA_ROW + 1):
            ws[f"{col_letter}{row}"].number_format = "0"

    # 7) Whole-number validation on Transshipment count (AQ), matching Flight Frequency (AR).
    aq_dv = DataValidation(type="whole", operator="greaterThanOrEqual", formula1="0", allow_blank=True)
    aq_dv.error = "Enter a whole number (0 or more)."
    aq_dv.errorTitle = "Invalid Transshipment count"
    ws.add_data_validation(aq_dv)
    aq_dv.add(f"AQ{FIRST_DATA_ROW}:AQ{LAST_DATA_ROW}")


def fix_sea_air_sheet(ws) -> None:
    # 8) Row 25 used '=ROUND(Q25/P25,1)' (Annual Weight) instead of the rest of the
    # column's '=+R{row}/P{row}' pattern (Chargeable Weight). Align it with the rest.
    ws["S25"] = "=+R25/P25"


def main() -> None:
    wb = load_workbook(SOURCE_PATH)
    fix_air_sheet(wb["AIR"])
    fix_sea_air_sheet(wb["Sea-Air"])
    wb.save(OUTPUT_PATH)
    print(f"Corrected copy written to: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
