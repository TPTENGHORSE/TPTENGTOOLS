"""Excel export of the CLT dashboard with live formulas for every calculation."""

from __future__ import annotations

from datetime import datetime, timedelta
from io import BytesIO

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

HEADER_FILL = PatternFill("solid", fgColor="1F2937")
HEADER_FONT = Font(bold=True, color="FFFFFF")
INPUT_FILL = PatternFill("solid", fgColor="FEF9C3")
TITLE_FONT = Font(bold=True, size=14)


def _header(ws, row: int, values: list[str]) -> None:
    for col, value in enumerate(values, start=1):
        cell = ws.cell(row=row, column=col, value=value)
        cell.fill, cell.font = HEADER_FILL, HEADER_FONT
        cell.alignment = Alignment(horizontal="center")


def _widths(ws, widths: list[int]) -> None:
    for index, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(index)].width = width


def _put(ws, row: int, label: str, value, note: str = "", fmt: str | None = None,
         is_input: bool = False) -> None:
    ws.cell(row=row, column=1, value=label)
    cell = ws.cell(row=row, column=2, value=value)
    if fmt:
        cell.number_format = fmt
    if is_input:
        cell.fill = INPUT_FILL
    ws.cell(row=row, column=3, value=note)


def build_workbook(
    *,
    wacc_percent: float,
    use_carrying_rate: bool,
    carrying_percent: float,
    annual_days: float,
    annual_volume: float,
    plant_weeks: float,
    cargo_value: float,
    currency: str,
    scenario_a,
    scenario_b,
    vtt: dict | None,
    gantt_steps: list[dict] | None,
    gantt_labels: list[str],
) -> bytes:
    """Return an .xlsx file whose results are formulas over the input cells."""
    money = f'"{currency}" #,##0.00'
    wb = Workbook()

    # ---- Inputs --------------------------------------------------------
    ws = wb.active
    ws.title = "Inputs"
    ws["A1"], ws["A1"].font = "1. Global assumptions (yellow cells are inputs)", TITLE_FONT
    _header(ws, 2, ["Item", "Value", "Formula / note"])
    _put(ws, 3, "WACC (% per year)", wacc_percent, is_input=True)
    _put(ws, 4, "Use separate carrying cost rate", bool(use_carrying_rate), is_input=True)
    _put(ws, 5, "Inventory carrying cost rate (% per year)", carrying_percent, is_input=True)
    _put(ws, 6, "Financing rate used (decimal)", "=IF(B4,B5,B3)/100",
         "Carrying rate if selected, otherwise WACC, divided by 100", "0.00%")
    _put(ws, 7, "Days per year", annual_days, is_input=True)
    _put(ws, 8, "Annual container volume", annual_volume, is_input=True)
    _put(ws, 9, "Plant working weeks per year", plant_weeks, is_input=True)
    _put(ws, 10, "Cargo value per container", cargo_value, fmt=money, is_input=True)
    _put(ws, 11, "Currency", currency, is_input=True)
    _put(ws, 12, "Containers consumed per day", "=B8/B7", "Annual volume ÷ days per year", "#,##0.000")
    _put(ws, 13, "Containers per week", "=B8/B9", "Annual volume ÷ plant weeks", "#,##0.00")
    _widths(ws, [46, 20, 70])

    # ---- VTT -----------------------------------------------------------
    if vtt:
        wv = wb.create_sheet("VTT")
        wv["A1"], wv["A1"].font = "2. Route and lead-time change (VTT)", TITLE_FONT
        _header(wv, 2, ["Item", "Value", "Formula / note"])
        rows = [
            (3, "VTT route", vtt["route"], "", None, True),
            (4, "Current POL>POD (days)", vtt["current_pol_pod"], "", None, True),
            (5, "New POL>POD (days)", vtt["new_pol_pod"], "", None, True),
            (6, "POL>POD reduction (days)", "=B4-B5", "Current - new", None, False),
            (7, "Day of customer order", vtt["start_day"], "", None, True),
            (8, "Base final day (rounding)", vtt["base_final_day"], "", None, True),
            (9, "Base day before rounding (transport to plant)", vtt["base_pre_rounding"], "", None, True),
            (10, "Base cut-off day", vtt["base_cutoff_day"], "", None, True),
            (11, "Base ETD day", vtt["base_etd_day"], "", None, True),
            (12, "First day to POL", vtt["first_day_to_pol"], "", None, True),
            (13, "Cut-off weekday (Mon=0 … Sun=6)", vtt["cutoff_weekday"], vtt["cutoff_name"], None, True),
            (14, "ETD weekday (Mon=0 … Sun=6)", vtt["etd_weekday"], vtt["etd_name"], None, True),
            (15, "New cut-off day", "=B12+MOD(B13-MOD(B12-1,7),7)",
             "First day to POL + days until selected cut-off weekday", None, False),
            (16, "Days from cut-off to ETD weekday", "=MOD(B14-MOD(B15-1,7),7)", "", None, False),
            (17, "New ETD day", "=B15+IF(B16=0,7,B16)",
             "If both weekdays match, ETD is the following week", None, False),
            (18, "Total shift (days)", "=B11-B17+B6", "Base ETD - new ETD + POL>POD reduction", None, False),
            (19, "Sunday offset", "=MOD(B8,7)", "Base final day mod 7", None, False),
            (20, "New day before rounding", "=B9-B18", "Base day before rounding - total shift", None, False),
            (21, "New final day", "=IF(B18=0,B8,B19+7*-INT(-(B20-B19)/7))",
             "First Sunday on or after the new day before rounding", None, False),
            (22, "Current CLT (days)", "=B8-B7+1", "Final day - order day + 1", None, False),
            (23, "Current rounding (days)", "=B8-B9", "", None, False),
            (24, "New rounding to Sunday (days)", "=B21-B20", "", None, False),
            (25, "New CLT (days)", "=B21-B7+1", "", None, False),
            (26, "CLT change vs current (days)", "=B25-B22", "", None, False),
        ]
        for row, label, value, note, fmt, is_input in rows:
            _put(wv, row, label, value, note, fmt, is_input)

        wv["A28"], wv["A28"].font = "Effective CLT reduction vs. sea-transit reduction", Font(bold=True)
        _header(wv, 29, ["New POL>POD (days)", "POL>POD reduction (days)", "Total shift",
                         "Day before rounding", "Final day", "Rounding (days)",
                         "New CLT (days)", "CLT reduction (days)"])
        for i, reduction in enumerate(range(0, 15)):
            r = 30 + i
            wv.cell(r, 2, reduction)
            wv.cell(r, 1, f"=$B$4-B{r}")
            wv.cell(r, 3, f"=$B$11-$B$17+B{r}")
            wv.cell(r, 4, f"=$B$9-C{r}")
            wv.cell(r, 5, f"=IF(C{r}=0,$B$8,$B$19+7*-INT(-(D{r}-$B$19)/7))")
            wv.cell(r, 6, f"=E{r}-D{r}")
            wv.cell(r, 7, f"=E{r}-$B$7+1")
            wv.cell(r, 8, f"=$B$22-G{r}")
        _widths(wv, [46, 24, 50, 20, 14, 16, 16, 20])

        if gantt_steps:
            wg = wb.create_sheet("VTT Timeline")
            wg["A1"], wg["A1"].font = "VTT timeline with new POL>POD and schedule", TITLE_FONT
            first_col = 5
            num_days = max(110, max(s["final"] for s in gantt_steps) + 7)
            today = datetime.today()
            wg["A2"], wg["B2"] = "Day 1 (Monday)", (today - timedelta(days=today.weekday())).date()
            wg["B2"].number_format = "dd/mm/yyyy"
            wg["B2"].fill = INPUT_FILL
            _header(wg, 4, ["Steps", "Day", "Day+", "Final Day"])
            for c in range(1, 5):
                wg.cell(3, c).fill = HEADER_FILL
            thin = Side(style="thin", color="DDDDDD")
            for n in range(1, num_days + 1):
                col = first_col + n - 1
                letter = get_column_letter(col)
                wg.column_dimensions[letter].width = 4.5
                week = wg.cell(3, col, f"=\"W\"&WEEKNUM($B$2+{n - 1},21)")
                week.fill, week.font = PatternFill("solid", fgColor="FFFBE6"), Font(bold=True, size=8)
                wg.cell(4, col, n).alignment = Alignment(horizontal="center")
                wg.cell(4, col).fill = PatternFill("solid", fgColor="F5F5F5")
                wg.cell(4, col).font = Font(size=8)
            for i, (label, step) in enumerate(zip(gantt_labels, gantt_steps)):
                r = 5 + i
                wg.cell(r, 1, label)
                wg.cell(r, 2, step["day"])
                wg.cell(r, 3, step["plus"])
                wg.cell(r, 4, step["final"])
                for c in range(first_col, first_col + num_days):
                    wg.cell(r, c).border = Border(left=thin, right=thin, top=thin, bottom=thin)
                for seg in step["segments"]:
                    color = seg["fill"].lstrip("#").upper()
                    for d in range(seg["start"], seg["end"] + 1):
                        wg.cell(r, first_col + d - 1).fill = PatternFill("solid", fgColor=color)
            wg.cell(5 + len(gantt_steps) + 1, 1, "Green = process days, blue = security time")
            _widths(wg, [34, 10, 10, 12])
            wg.freeze_panes = wg.cell(5, first_col)

    # ---- Scenarios -----------------------------------------------------
    wsn = wb.create_sheet("Scenarios", 1)
    wsn["A1"], wsn["A1"].font = "3. Scenario comparison", TITLE_FONT
    _header(wsn, 2, ["Metric", "Scenario A (current)", "Scenario B (proposed)", "Change (B − A)",
                     "Formula"])
    clt_a = "=VTT!B22" if vtt else scenario_a.customer_lead_time_days
    clt_b = "=VTT!B25" if vtt else scenario_b.customer_lead_time_days
    inputs = [
        (3, "Ocean freight per container", scenario_a.ocean_freight_per_container,
         scenario_b.ocean_freight_per_container),
        (4, "Port-to-plant transportation per container",
         scenario_a.inland_transport_per_container, scenario_b.inland_transport_per_container),
        (5, "Additional logistics cost per container",
         scenario_a.additional_logistics_per_container, scenario_b.additional_logistics_per_container),
        (6, "CLT (days)", clt_a, clt_b),
    ]
    for r, label, a, b in inputs:
        wsn.cell(r, 1, label)
        wsn.cell(r, 2, a).fill = INPUT_FILL
        wsn.cell(r, 3, b).fill = INPUT_FILL
        wsn.cell(r, 4, f"=C{r}-B{r}")
        for c in (2, 3, 4):
            wsn.cell(r, c).number_format = "#,##0" if r == 6 else money
    wsn.cell(6, 5, "From the VTT sheet" if vtt else "Manual input")
    calc = [
        (7, "Transport cost / container", "={c}3+{c}4+{c}5", "Ocean + inland + additional"),
        (8, "Pipeline inventory value", "=Inputs!$B$10*Inputs!$B$12*{c}6",
         "Cargo value × containers per day × CLT days"),
        (9, "Finance cost / container", "=Inputs!$B$10*Inputs!$B$6*{c}6/Inputs!$B$7",
         "Cargo value × rate × CLT days ÷ days per year"),
        (10, "Annual financing cost", "={c}8*Inputs!$B$6", "Pipeline value × annual rate"),
        (11, "Annual transportation cost", "={c}7*Inputs!$B$8", "Transport per container × annual volume"),
        (12, "Economic cost / container", "={c}7+{c}9", "Transport + finance per container"),
        (13, "Annual economic logistics cost", "={c}11+{c}10", "Annual transport + annual financing"),
    ]
    for r, label, formula, note in calc:
        wsn.cell(r, 1, label)
        for col in ("B", "C"):
            cell = wsn[f"{col}{r}"]
            cell.value = formula.format(c=col)
            cell.number_format = money
        wsn.cell(r, 4, f"=C{r}-B{r}").number_format = money
        wsn.cell(r, 5, note)
    for r in range(3, 14):
        wsn.cell(r, 4).fill = PatternFill("solid", fgColor="F1F5F9")
    _widths(wsn, [46, 24, 24, 20, 50])

    chart = BarChart()
    chart.type, chart.grouping, chart.overlap = "col", "stacked", 100
    chart.title = f"Annual cost by scenario ({currency})"
    chart.add_data(Reference(wsn, min_col=1, max_col=3, min_row=10, max_row=11), from_rows=True,
                   titles_from_data=True)
    chart.set_categories(Reference(wsn, min_col=2, max_col=3, min_row=2))
    chart.width, chart.height = 16, 8
    wsn.add_chart(chart, "A16")

    # ---- Summary -------------------------------------------------------
    wsum = wb.create_sheet("Decision summary", 2)
    wsum["A1"], wsum["A1"].font = "4. Decision summary (Scenario B vs. A)", TITLE_FONT
    _header(wsum, 2, ["Indicator", "Value", "Formula / note"])
    S = "Scenarios!"
    summary = [
        (3, "CLT reduction (days)", f"={S}B6-{S}C6", "CLT A - CLT B", "#,##0.00"),
        (4, "CLT reduction (%)", f"=IF({S}B6=0,0,B3/{S}B6)", "CLT reduction ÷ CLT A", "0.00%"),
        (5, "CLT reduction (weeks)", "=B3/7", "", "#,##0.00"),
        (6, "Working capital released", "=Inputs!B10*Inputs!B13*B5",
         "Cargo value × containers per week × weeks of CLT reduction", money),
        (7, "Working capital released (pipeline difference)", f"={S}B8-{S}C8",
         "Pipeline value A - pipeline value B", money),
        (8, "Working capital released per 1 week of CLT", "=Inputs!B10*Inputs!B13",
         "Cargo value × containers per week", money),
        (9, "Annual financing savings", f"={S}B10-{S}C10", "Financing cost A - B", money),
        (10, "Transportation savings per container", f"={S}B7-{S}C7", "Transport A - B", money),
        (11, "Transportation savings (annual)", f"={S}B11-{S}C11",
         "Negative = additional annual transport cost", money),
        (12, "Net annual economic benefit", f"={S}B13-{S}C13", "Annual economic cost A - B", money),
        (13, "Preferred scenario", '=IF(B12>0,"B",IF(B12<0,"A","Tie"))', "", None),
        (14, "Financing-only break-even freight premium / container", f"={S}B9-{S}C9",
         "Finance cost per container A - B", money),
        (15, "Annual maximum freight budget", "=B14*Inputs!B8", "Break-even premium × annual volume", money),
        (16, "Actual additional transport cost / container", f"={S}C7-{S}B7", "Transport B - A", money),
        (17, "Actual additional annual transport cost", f"={S}C11-{S}B11", "", money),
        (18, "Transport premium check",
         '=IF(B16<B14,"Below break-even premium",IF(B16=B14,"At break-even","Exceeds break-even premium"))',
         "Actual premium vs financing-only break-even", None),
    ]
    for r, label, formula, note, fmt in summary:
        _put(wsum, r, label, formula, note, fmt)
    _widths(wsum, [56, 26, 64])

    # ---- Notes ---------------------------------------------------------
    wn = wb.create_sheet("Assumptions")
    notes = [
        "Calculation breakdown & assumptions",
        "Pipeline inventory value = cargo value × containers consumed per day × CLT days.",
        "Working capital released = annual volume ÷ plant weeks × cargo value × (CLT reduction ÷ 7).",
        "Per-container financing cost = cargo value × annual rate × CLT days ÷ days per year.",
        "Annual financing cost = pipeline inventory value × annual rate.",
        "Transportation cost per container = ocean freight + inland transportation + additional logistics.",
        "Net annual economic benefit = Scenario A annual economic cost - Scenario B annual economic cost.",
        "Inventory value and released working capital are balance-sheet measures, not P&L savings.",
        "Safety stock, service level, obsolescence, warehousing, demurrage, detention and stockouts are not modeled.",
    ]
    for i, text in enumerate(notes, start=1):
        wn.cell(i, 1, text)
    wn["A1"].font = TITLE_FONT
    wn.column_dimensions["A"].width = 110

    buffer = BytesIO()
    wb.save(buffer)
    return buffer.getvalue()
