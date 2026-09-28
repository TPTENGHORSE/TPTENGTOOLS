"""Maritime transit time (POL -> POD) analysis for BuyCo departures data.

Reads the ocean freight tracking export (ATD = actual departure from POL,
ATA = actual arrival at POD) and computes, per POL>POD flow, both central
tendency and tail statistics so the recommended planning transit time
accounts for variability instead of relying only on the average.
"""

from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.worksheet.worksheet import Worksheet

DATA_PATH = Path(
    r"C:\Users\OLMEDOGALVEZJorgeLui\OneDrive - Horse\Exchange VRAC\02_Engineering Department"
    r"\10. KPIs\DB\VTT_Differences\Overseas_Buyco\Departures-2025 and 2026.xlsx"
)
SHEET_NAME = "Hoja1"
OUTPUT_PATH = Path(__file__).parent / "Ocean_Transit_Time_Analysis.xlsx"

# Reliability tiers used to flag how much to trust a flow's percentile-based recommendation.
MIN_N_FOR_HIGH_CONFIDENCE = 15
MIN_N_FOR_MEDIUM_CONFIDENCE = 5

# Recommended planning percentile: the transit time within which X% of shipments arrive.
SERVICE_LEVEL_PERCENTILE = 90


def load_shipments() -> pd.DataFrame:
    cols = [
        "Container number", "POL", "POL - Country", "POL - Unlocode",
        "POD", "POD - Country", "POD - Unlocode", "ATD", "ATA",
    ]
    return pd.read_excel(DATA_PATH, sheet_name=SHEET_NAME, usecols=cols)


def clean_shipments(df: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    total = len(df)
    df = df.dropna(subset=["POL - Unlocode", "POD - Unlocode", "ATD", "ATA"]).copy()
    df["transit_days"] = (df["ATA"] - df["ATD"]).dt.total_seconds() / 86400

    missing_dropped = total - len(df)
    invalid_negative = int((df["transit_days"] <= 0).sum())
    df = df[df["transit_days"] > 0]

    # Ocean transits under 1 day or beyond 150 days are almost certainly data-entry errors
    # (mismatched ATD/ATA, wrong port code), not real transit performance.
    invalid_range = int((~df["transit_days"].between(1, 150)).sum())
    df = df[df["transit_days"].between(1, 150)]

    stats = {
        "total_rows": total,
        "dropped_missing": missing_dropped,
        "dropped_negative_or_zero": invalid_negative,
        "dropped_out_of_range": invalid_range,
        "clean_rows": len(df),
    }
    return df, stats


def _percentile(series: pd.Series, p: float) -> float:
    return float(np.percentile(series, p))


def summarize_flows(df: pd.DataFrame) -> pd.DataFrame:
    records = []
    for (pol, pod), g in df.groupby(["POL - Unlocode", "POD - Unlocode"]):
        tt = g["transit_days"]
        n = len(tt)

        mean = tt.mean()
        std = tt.std(ddof=1) if n > 1 else 0.0
        cv = std / mean if mean else 0.0

        p50 = _percentile(tt, 50)
        p75 = _percentile(tt, 75)
        p90 = _percentile(tt, 90)
        p95 = _percentile(tt, 95)

        # Normal-approximation cross-check for the 90th percentile (z(0.90) = 1.2816).
        normal_p90 = mean + 1.2816 * std
        skew_gap_days = p90 - normal_p90  # positive gap -> right-skewed / long-tail flow

        if n >= MIN_N_FOR_HIGH_CONFIDENCE:
            confidence = "High"
        elif n >= MIN_N_FOR_MEDIUM_CONFIDENCE:
            confidence = "Medium"
        else:
            confidence = "Low"

        # Recommended planning transit time: empirical P90 when there is enough history;
        # otherwise a conservative median + 1 std fallback for unstable small samples.
        if confidence == "Low":
            recommended_tt = p50 + std
            method = "Median + 1 Std Dev (fallback, n < 5)"
        else:
            recommended_tt = p90
            method = f"P{SERVICE_LEVEL_PERCENTILE} empirical percentile"

        records.append(
            {
                "POL": pol,
                "POD": pod,
                "Origin": g["POL"].iloc[0],
                "Origin Country": g["POL - Country"].iloc[0],
                "Destination": g["POD"].iloc[0],
                "Destination Country": g["POD - Country"].iloc[0],
                "Shipments (n)": n,
                "Confidence": confidence,
                "Mean (days)": round(mean, 1),
                "Median / P50 (days)": round(p50, 1),
                "Std Dev (days)": round(std, 1),
                "CV (Std/Mean)": round(cv, 2),
                "P75 (days)": round(p75, 1),
                "P90 (days)": round(p90, 1),
                "P95 (days)": round(p95, 1),
                "Normal-approx P90 (days)": round(normal_p90, 1),
                "Tail skew gap (P90 - Normal P90)": round(skew_gap_days, 1),
                "Recommended Transit Time (days)": round(recommended_tt, 1),
                "Method": method,
            }
        )

    return (
        pd.DataFrame(records)
        .sort_values("Shipments (n)", ascending=False)
        .reset_index(drop=True)
    )


def print_insights(flows: pd.DataFrame, clean_stats: dict) -> None:
    print("=" * 80)
    print("OCEAN TRANSIT TIME ANALYSIS - POL > POD")
    print("=" * 80)
    print(f"Total rows in file:                 {clean_stats['total_rows']}")
    print(f"Dropped (missing POL/POD/ATD/ATA):  {clean_stats['dropped_missing']}")
    print(f"Dropped (ATA <= ATD):                {clean_stats['dropped_negative_or_zero']}")
    print(f"Dropped (outside 1-150 days range):  {clean_stats['dropped_out_of_range']}")
    print(f"Clean shipments used:                {clean_stats['clean_rows']}")
    print(f"Distinct POL>POD flows:              {len(flows)}")
    print()

    print("Methodology")
    print("-" * 80)
    print(
        "1) transit_days = ATA (actual arrival at POD) - ATD (actual departure from POL).\n"
        "2) For each POL>POD flow we compute the full empirical distribution (not just the\n"
        "   mean): mean, median (P50), P75, P90, P95, std dev, and CV = std/mean.\n"
        "3) Recommended planning transit time = P90 (the value that covers 90% of historical\n"
        "   shipments on that flow), NOT the average. Ocean transit distributions are\n"
        "   right-skewed (congestion, rollings, weather add delay far more often than they\n"
        "   subtract it), so planning on the mean means ~50% of shipments would arrive later\n"
        "   than promised.\n"
        "4) We cross-check empirical P90 against a Normal-distribution estimate\n"
        "   (mean + 1.2816 * std, since z(0.90) = 1.2816). A positive 'tail skew gap'\n"
        "   (empirical P90 > normal-approx P90) means the flow has a heavier right tail than\n"
        "   a Normal model predicts - i.e. delays are more frequent/severe than 'average\n"
        "   variability' suggests, reinforcing why the empirical percentile (not mean+z*std)\n"
        "   should drive the operational commitment.\n"
        "5) Confidence tiers by sample size: High (n>=15), Medium (n>=5), Low (n<5). Low\n"
        "   confidence flows fall back to Median + 1 Std Dev because P90 is unstable on\n"
        "   very small samples."
    )
    print()

    top_volume = flows.head(15)
    print("Top 15 flows by shipment volume:")
    print(
        top_volume[
            [
                "POL", "POD", "Shipments (n)", "Confidence", "Mean (days)",
                "P90 (days)", "Recommended Transit Time (days)", "CV (Std/Mean)",
            ]
        ].to_string(index=False)
    )
    print()

    high_variability = (
        flows[flows["Confidence"] != "Low"]
        .sort_values("CV (Std/Mean)", ascending=False)
        .head(10)
    )
    print("Top 10 most volatile reliable flows (highest CV = std/mean):")
    print(
        high_variability[
            ["POL", "POD", "Shipments (n)", "Mean (days)", "Std Dev (days)", "CV (Std/Mean)", "P90 (days)"]
        ].to_string(index=False)
    )
    print()

    gap_vs_mean = (flows["Recommended Transit Time (days)"] - flows["Mean (days)"]).mean()
    print(
        f"On average across all flows, using P90 instead of the mean adds {gap_vs_mean:.1f} "
        "extra days of buffer to the planning transit time - this is the buffer needed to "
        "reach ~90% on-time reliability."
    )


# POL>POD flows to spotlight in their own sheet for the current review.
SELECTED_FLOWS = [
    ("CNSHA", "ESALG"),
    ("CNSHA", "ESVLC"),
    ("CNSHA", "PTLEI"),
    ("CNSHA", "TRIZT"),
    ("FRLEH", "ARROS"),
    ("JPNGO", "BRPNG"),
    ("JPNGO", "TRYAR"),
    ("JPYOK", "ESALG"),
    ("JPYOK", "FRLEH"),
    ("JPYOK", "PTLEI"),
    ("JPYOK", "ROCND"),
]


def filter_selected_flows(flows: pd.DataFrame) -> pd.DataFrame:
    pairs = pd.DataFrame(SELECTED_FLOWS, columns=["POL", "POD"])
    selected = flows.merge(pairs, on=["POL", "POD"], how="right")
    missing = selected[selected["Shipments (n)"].isna()][["POL", "POD"]]
    if not missing.empty:
        print("Warning: no shipments found for these requested flows:")
        print(missing.to_string(index=False))
    return selected


# --- Excel workbook with live formulas (auditable in Excel itself) -------------------

RAW_SHEET_NAME = "Raw Data"

FLOW_HEADERS = [
    "POL", "POD", "Origin", "Origin Country", "Destination", "Destination Country",
    "Shipments (n)", "Confidence", "Mean (days)", "Median / P50 (days)", "Std Dev (days)",
    "CV (Std/Mean)", "P75 (days)", "P90 (days)", "P95 (days)", "Normal-approx P90 (days)",
    "Tail skew gap (P90 - Normal P90)", "Recommended Transit Time (days)", "Method",
]


def _write_raw_data_sheet(ws: Worksheet, clean: pd.DataFrame) -> int:
    """Write cleaned shipments with a live '=ATA-ATD' formula. Returns last data row."""
    ws.append(["POL", "POD", "ATD", "ATA", "Transit Days"])
    # itertuples sanitizes column names with spaces/dashes, so use positional access instead.
    pol_col = clean.columns.get_loc("POL - Unlocode")
    pod_col = clean.columns.get_loc("POD - Unlocode")
    atd_col = clean.columns.get_loc("ATD")
    ata_col = clean.columns.get_loc("ATA")
    for i, values in enumerate(clean.itertuples(index=False, name=None), start=2):
        r = i
        ws.cell(row=r, column=1, value=values[pol_col])
        ws.cell(row=r, column=2, value=values[pod_col])
        atd_cell = ws.cell(row=r, column=3, value=values[atd_col])
        atd_cell.number_format = "yyyy-mm-dd hh:mm"
        ata_cell = ws.cell(row=r, column=4, value=values[ata_col])
        ata_cell.number_format = "yyyy-mm-dd hh:mm"
        tt_cell = ws.cell(row=r, column=5, value=f"=D{r}-C{r}")
        tt_cell.number_format = "0.00"
    return len(clean) + 1


def _write_flow_formula_sheet(ws: Worksheet, flows: pd.DataFrame, last_raw_row: int) -> None:
    """Write one row per flow with Excel formulas (COUNTIFS/AVERAGEIFS/array MEDIAN-PERCENTILE-STDEV)
    that recompute the stats directly from the 'Raw Data' sheet, instead of pre-baked values."""
    ws.append(FLOW_HEADERS)

    pol_rng = f"'{RAW_SHEET_NAME}'!$A$2:$A${last_raw_row}"
    pod_rng = f"'{RAW_SHEET_NAME}'!$B$2:$B${last_raw_row}"
    tt_rng = f"'{RAW_SHEET_NAME}'!$E$2:$E${last_raw_row}"

    label_cols = ["POL", "POD", "Origin", "Origin Country", "Destination", "Destination Country"]
    for i, values in enumerate(flows[label_cols].itertuples(index=False, name=None), start=2):
        r = i
        for col_idx, value in enumerate(values, start=1):
            ws.cell(row=r, column=col_idx, value=value)

        n_formula = f"=COUNTIFS({pol_rng},A{r},{pod_rng},B{r})"
        ws.cell(row=r, column=7, value=n_formula)

        confidence_formula = f'=IF(G{r}>=15,"High",IF(G{r}>=5,"Medium","Low"))'
        ws.cell(row=r, column=8, value=confidence_formula)

        mean_formula = f"=IFERROR(AVERAGEIFS({tt_rng},{pol_rng},A{r},{pod_rng},B{r}),0)"
        ws.cell(row=r, column=9, value=mean_formula).number_format = "0.0"

        crit = f"({pol_rng}=A{r})*({pod_rng}=B{r})"
        median_formula = f"=IFERROR(MEDIAN(IF({crit},{tt_rng})),0)"
        ws.cell(row=r, column=10, value=ArrayFormula(f"J{r}", median_formula)).number_format = "0.0"

        std_formula = f"=IFERROR(STDEV(IF({crit},{tt_rng})),0)"
        ws.cell(row=r, column=11, value=ArrayFormula(f"K{r}", std_formula)).number_format = "0.0"

        ws.cell(row=r, column=12, value=f"=IF(I{r}=0,0,K{r}/I{r})").number_format = "0.00"

        p75_formula = f"=IFERROR(PERCENTILE(IF({crit},{tt_rng}),0.75),0)"
        ws.cell(row=r, column=13, value=ArrayFormula(f"M{r}", p75_formula)).number_format = "0.0"

        p90_formula = f"=IFERROR(PERCENTILE(IF({crit},{tt_rng}),0.9),0)"
        ws.cell(row=r, column=14, value=ArrayFormula(f"N{r}", p90_formula)).number_format = "0.0"

        p95_formula = f"=IFERROR(PERCENTILE(IF({crit},{tt_rng}),0.95),0)"
        ws.cell(row=r, column=15, value=ArrayFormula(f"O{r}", p95_formula)).number_format = "0.0"

        ws.cell(row=r, column=16, value=f"=I{r}+1.2816*K{r}").number_format = "0.0"
        ws.cell(row=r, column=17, value=f"=N{r}-P{r}").number_format = "0.0"
        ws.cell(row=r, column=18, value=f"=IF(G{r}<5,J{r}+K{r},N{r})").number_format = "0.0"
        ws.cell(
            row=r, column=19,
            value=f'=IF(G{r}<5,"Median + 1 Std Dev (fallback, n<5)","P90 empirical percentile")',
        )


def write_formula_workbook(clean: pd.DataFrame, flows: pd.DataFrame, selected_flows: pd.DataFrame) -> None:
    wb = Workbook()
    raw_ws = wb.active
    raw_ws.title = RAW_SHEET_NAME
    last_raw_row = _write_raw_data_sheet(raw_ws, clean)

    all_ws = wb.create_sheet("All Flows")
    _write_flow_formula_sheet(all_ws, flows, last_raw_row)

    selected_ws = wb.create_sheet("Selected Flows")
    _write_flow_formula_sheet(selected_ws, selected_flows, last_raw_row)

    wb.save(OUTPUT_PATH)


def main() -> None:
    raw = load_shipments()
    clean, clean_stats = clean_shipments(raw)
    flows = summarize_flows(clean)
    selected_flows = filter_selected_flows(flows)

    write_formula_workbook(clean, flows, selected_flows)

    print_insights(flows, clean_stats)
    print(f"\nFull per-flow table exported to: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
