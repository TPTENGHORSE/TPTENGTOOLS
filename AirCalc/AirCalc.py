"""AirCalc - Airfreight Rates Calculator (Streamlit port of Calculator_v2 tab).

Reproduces the logic of the "Calculator_v2" sheet found in
"Airfreights Rates Calculator v2 Tender 2026.xlsm":

- CBW (chargeable weight) = MAX(volume_m3 * 167, weight_kg)
- For the selected Origin/Destination country pair, every matching lane in the
  "1_Consolidated_AIR" table is priced using its 5 tariff legs:
    1. Per Declaration (flat fee)
    2. Pickup leg: rate/kg bounded by [Min, Max]
    3. Second leg: rate/kg bounded by a Min only
    4. Linehaul: weight-break tariff table, bounded by a Min
    5. Delivery leg: rate/kg bounded by [Min, Max]
    6. Final leg: rate/kg bounded by a Min (uses <= comparator)
  Total Rate (EUR) = sum of the legs above.
  EUR/Kg = Total Rate / CBW + DHL adder (if the carrier/lane has one).
  Total Airfreight = EUR/Kg * CBW.
"""

from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st

XLSM_PATH = Path(__file__).parent / "Airfreights Rates Calculator v2 Tender 2026.xlsm"
CONSOLIDATED_SHEET = "1_Consolidated_AIR"
DHL_ADDER_SHEET = "DHL_ADDER"
HEADER_ROW = 3  # 0-indexed row of the header inside the consolidated sheet (Excel row 4)

# Weight-break columns (in ascending breakpoint order) used for the linehaul leg.
WEIGHT_BREAK_COLUMNS = [
    "<45 kg",
    ">45 kg",
    ">100 kg",
    ">300 kg",
    ">500 kg",
    ">1000 kg",
    ">2500 kg",
    ">5000 kg",
]
# Upper bound (inclusive) of each break, aligned with WEIGHT_BREAK_COLUMNS[:-1].
WEIGHT_BREAKPOINTS = [45, 100, 300, 500, 1000, 2500, 5000]

DISPLAY_COLUMNS = {
    "Shipper/ILN": "ILN/Supplier",
    "Type of Flow": "Lane Type",
    "Service Level (Normal, Economy and Express)": "Service Level",
    "Origin Airport": "Origin Airport",
    "Destination Airport": "Destin Airport",
    "Provider_Name": "Airfreight Carrier",
    "Total Airfreight": "Total Airfreight",
    "€/Kg": "Eur/Kg",
    "Nº Transshipment": "#TS",
    "Total TT (days)": "Transit Time (Days)",
}

SERVICE_LEVEL_ORDER = {"Economy": 0, "Normal": 1, "Express": 2}
LANE_TYPE_ORDER = {"Main": 0, "Backup": 1, "Back-Up": 1}


@st.cache_data(show_spinner="Loading rate tables...")
def load_rate_table() -> pd.DataFrame:
    df = pd.read_excel(XLSM_PATH, sheet_name=CONSOLIDATED_SHEET, header=HEADER_ROW)
    numeric_cols = [
        "Per Declaration",
        "Min",
        "Pickup Flat per kg",
        "Max",
        "Min4",
        "Per KG",
        "Min7",
        *WEIGHT_BREAK_COLUMNS,
        "Min8",
        "Delivery Flat per kg",
        "Max9",
        "Min10",
        "Per KG11",
    ]
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
    df = df.dropna(subset=["Country Code", "Country Code Destin"])
    return df


@st.cache_data(show_spinner=False)
def load_dhl_adder() -> dict:
    """Map (Provider_Name, Origin Country, Destin Country) -> Eur/kg adder."""
    df = pd.read_excel(XLSM_PATH, sheet_name=DHL_ADDER_SHEET, header=None, skiprows=3)
    df.columns = ["Provider_Name", "Key", "Origin", "Destin", "EurKg"] + list(df.columns[5:])
    df = df.dropna(subset=["Provider_Name", "Origin", "Destin"])
    return {
        (row.Provider_Name, row.Origin, row.Destin): float(row.EurKg)
        for row in df.itertuples()
    }


def linehaul_tariff(cbw: float, row: pd.Series) -> float:
    """Pick the per-kg tariff for the linehaul leg based on the CBW weight break."""
    for col, upper in zip(WEIGHT_BREAK_COLUMNS[:-1], WEIGHT_BREAKPOINTS):
        if cbw <= upper:
            return row[col]
    return row[WEIGHT_BREAK_COLUMNS[-1]]


def compute_rates(lanes: pd.DataFrame, cbw: float, dhl_adder: dict) -> pd.DataFrame:
    lanes = lanes.copy()

    pickup = np.clip(cbw * lanes["Pickup Flat per kg"], lanes["Min"], lanes["Max"])
    leg2 = np.maximum(cbw * lanes["Per KG"], lanes["Min4"])

    tarifa = lanes.apply(lambda row: linehaul_tariff(cbw, row), axis=1)
    linehaul = np.maximum(tarifa * cbw, lanes["Min7"])

    leg4 = np.clip(cbw * lanes["Delivery Flat per kg"], lanes["Min8"], lanes["Max9"])

    leg5_raw = cbw * lanes["Per KG11"]
    leg5 = np.where(leg5_raw <= lanes["Min10"], lanes["Min10"], leg5_raw)

    total_rate = lanes["Per Declaration"] + pickup + leg2 + linehaul + leg4 + leg5

    adder = lanes.apply(
        lambda row: dhl_adder.get(
            (row["Provider_Name"], row["Country Code"], row["Country Code Destin"]), 0.0
        ),
        axis=1,
    )

    lanes["€/Kg"] = np.where(cbw > 0, total_rate / cbw, 0.0) + adder
    lanes["Total Airfreight"] = lanes["€/Kg"] * cbw
    return lanes


def main() -> None:
    st.set_page_config(page_title="AirCalc - Airfreight Rates Calculator", page_icon="✈️", layout="wide")
    st.title("✈️ AirCalc - Airfreight Rates Calculator")
    st.caption(
        "All-In V2.0 · Based on the latest air freight tender 2026. Costs are estimated and "
        "may vary in case of additional inspections, operational requirements or geopolitical factors."
    )

    if not XLSM_PATH.exists():
        st.error(f"Rate file not found: {XLSM_PATH.name}")
        return

    rate_table = load_rate_table()
    dhl_adder = load_dhl_adder()

    origin_options = sorted(rate_table["Country Code"].dropna().unique())

    st.subheader("1. Select Origin - Destination")
    col1, col2 = st.columns(2)
    with col1:
        origin = st.selectbox("Origin Country", origin_options)
    destin_options = sorted(
        rate_table.loc[rate_table["Country Code"] == origin, "Country Code Destin"].dropna().unique()
    )
    with col2:
        destin = st.selectbox("Destin Country", destin_options)

    st.subheader("2. Enter shipment data")
    col3, col4 = st.columns(2)
    with col3:
        volume = st.number_input("Volume (m³)", min_value=0.0, value=0.0, step=0.1, format="%.3f")
    with col4:
        weight = st.number_input("Weight (kg)", min_value=0.0, value=0.0, step=1.0)

    cbw = max(volume * 167, weight)
    st.metric("CBW (kg) - Chargeable Weight", f"{cbw:,.2f}")

    st.subheader("3. Review results")
    lanes = rate_table[
        (rate_table["Country Code"] == origin) & (rate_table["Country Code Destin"] == destin)
    ]

    if lanes.empty:
        st.warning("No lane found for this Origin-Destination pair.")
        return

    if cbw <= 0:
        st.info("Enter a volume and/or weight to calculate rates.")
        return

    results = compute_rates(lanes, cbw, dhl_adder)
    # Select by original (unique) column names first, then rename, to avoid
    # colliding with the raw sheet's unrelated "Service Level" (BA) column.
    results = results[list(DISPLAY_COLUMNS.keys())].rename(columns=DISPLAY_COLUMNS)
    results["Service Level"] = results["Service Level"].str.title()
    results["Lane Type"] = results["Lane Type"].str.title()
    service_level_rank = results["Service Level"].map(SERVICE_LEVEL_ORDER).fillna(len(SERVICE_LEVEL_ORDER))
    lane_type_rank = results["Lane Type"].map(LANE_TYPE_ORDER).fillna(len(LANE_TYPE_ORDER))
    results = (
        results.assign(_service_level_rank=service_level_rank, _lane_type_rank=lane_type_rank)
        .sort_values(
            ["ILN/Supplier", "_lane_type_rank", "Origin Airport", "Destin Airport", "_service_level_rank"]
        )
        .drop(columns=["_service_level_rank", "_lane_type_rank"])
        .reset_index(drop=True)
    )

    st.dataframe(
        results.style.format(
            {
                "Total Airfreight": "{:,.0f} €",
                "Eur/Kg": "{:,.2f}",
                "Transit Time (Days)": "{:,.0f}",
                "#TS": "{:,.0f}",
            }
        ),
        use_container_width=True,
    )

    with st.expander("How it works"):
        st.markdown(
            "- **CBW (kg)**: Chargeable weight = MAX(Volume × 167, Weight)\n"
            "- **ILN/Supplier**, if applicable\n"
            "- **Lane Type**: some lanes have back-ups\n"
            "- **Service Level**: Economy, Normal or Express\n"
            "- **Origin-Destination Airports**\n"
            "- **Airfreight Carrier**\n"
            "- **Total Airfreight**: door-to-door cost\n"
            "- **Eur/Kg**\n"
            "- **#TS**: Transshipment numbers\n"
            "- **Transit Time**: door-to-door total"
        )


if __name__ == "__main__":
    main()
