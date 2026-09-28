"""Horse Global Logistics - Active Flows Cartography.

Streamlit app to look up active transport flows by origin/destination
(no cost data). Airfreight is sourced live from AirCalc's rate table
(1_Consolidated_AIR). Overseas and Inland are wired as placeholders until
their own databases are connected.
"""

from pathlib import Path

import pandas as pd
import streamlit as st

AIRCALC_XLSM = Path(__file__).parent.parent / "AirCalc" / "Airfreights Rates Calculator v2 Tender 2026.xlsm"
CONSOLIDATED_SHEET = "1_Consolidated_AIR"
HEADER_ROW = 3  # 0-indexed row of the header inside the consolidated sheet (Excel row 4).

OVERSEAS_XLSX = Path(__file__).parent / "Overseas.xlsx"
OVERSEAS_SHEET = "budget 2026-2027"

# Destination country -> recipient Horse plant. Extend as new lanes/plants are onboarded.
RECIPIENT_BY_COUNTRY = {
    "CL": "900095 - Horse Chile",
    "RO": "900125 - Horse Romania",
    "TR": "900144 - Horse Turkey",
    "BR": "900186 - Horse Brasil",
    "ES": "910177 - Horse Valladolid",
}

CARTOGRAPHY_COLUMNS = [
    "Type Flow", "Flow", "POL", "POL Country", "POD", "POD Country",
    "Origin / ILN-Supplier", "Recipient", "Flow Type", "Commodity",
    "Carrier", "Transit Time (days)", "Lane Type",
]


@st.cache_data(show_spinner="Loading airfreight flows...")
def load_airfreight_flows() -> pd.DataFrame:
    if not AIRCALC_XLSM.exists():
        return pd.DataFrame(columns=CARTOGRAPHY_COLUMNS)

    df = pd.read_excel(AIRCALC_XLSM, sheet_name=CONSOLIDATED_SHEET, header=HEADER_ROW)
    df = df.dropna(subset=["Country Code", "Country Code Destin"])

    # One row per lane, ignoring the Economy/Normal/Express duplication used for pricing.
    group_cols = [
        "Shipper/ILN", "Type of Flow", "Country Code", "Origin Airport",
        "Country Code Destin", "Destination Airport", "Provider_Name",
    ]

    records = []
    for keys, g in df.groupby(group_cols, dropna=False):
        shipper, lane_type, pol_country, pol, pod_country, pod, carrier = keys

        # Prefer "Normal" service level as the representative transit time, else the fastest one.
        normal = g[g["Service Level (Normal, Economy and Express)"].str.title() == "Normal"]
        transit_time = normal["Total TT (days)"].iloc[0] if not normal.empty else g["Total TT (days)"].min()

        records.append(
            {
                "Type Flow": "Airfreight",
                "Flow": "Inbound",  # AirCalc currently only carries supplier -> plant parts lanes.
                "POL": pol,
                "POL Country": pol_country,
                "POD": pod,
                "POD Country": pod_country,
                "Origin / ILN-Supplier": shipper,
                "Recipient": RECIPIENT_BY_COUNTRY.get(pod_country, f"N/A - Unmapped ({pod_country})"),
                "Flow Type": "Emergency",
                "Commodity": "Parts",
                "Carrier": carrier,
                "Transit Time (days)": transit_time,
                "Lane Type": lane_type,
            }
        )

    return pd.DataFrame(records, columns=CARTOGRAPHY_COLUMNS)


@st.cache_data(show_spinner="Loading overseas flows...")
def load_overseas_flows() -> pd.DataFrame:
    if not OVERSEAS_XLSX.exists():
        return pd.DataFrame(columns=CARTOGRAPHY_COLUMNS)

    df = pd.read_excel(OVERSEAS_XLSX, sheet_name=OVERSEAS_SHEET)
    # The sheet has a continent-summary block below the real data; keep only actual flow rows.
    df = df.dropna(subset=["Flow", "POL", "POD"])
    df = df.rename(columns={"Origin": "Origin / ILN-Supplier", "Transit time": "Transit Time (days)"})
    df["Type Flow"] = "Overseas"
    df["Lane Type"] = None
    return df[CARTOGRAPHY_COLUMNS]


def load_inland_flows() -> pd.DataFrame:
    # TODO: wire up to the Inland flows database once available. POL/POD are N/A for Inland.
    return pd.DataFrame(columns=CARTOGRAPHY_COLUMNS)


def main() -> None:
    st.set_page_config(page_title="Horse GL - Active Flows", page_icon="🌍", layout="wide")
    st.title("🌍 Horse Global Logistics - Active Flows Cartography")
    st.caption("Look up active transport flows by origin/destination. No cost data shown.")

    type_flow = st.sidebar.selectbox("Type Flow", ["Airfreight", "Overseas", "Inland"])

    if type_flow == "Airfreight":
        flows = load_airfreight_flows()
    elif type_flow == "Overseas":
        flows = load_overseas_flows()
    else:
        flows = load_inland_flows()

    if flows.empty:
        st.info(
            f"No {type_flow.lower()} database is wired up yet. "
            "Airfreight is live from AirCalc; Overseas/Inland will follow."
        )
        return

    st.subheader("Filters")
    col1, col2, col3 = st.columns(3)
    with col1:
        pol_countries = sorted(flows["POL Country"].dropna().unique())
        pol_country = st.selectbox("POL Country (Origin)", ["All"] + pol_countries)
    with col2:
        pod_countries = sorted(flows["POD Country"].dropna().unique())
        pod_country = st.selectbox("POD Country (Destination)", ["All"] + pod_countries)
    with col3:
        suppliers = sorted(flows["Origin / ILN-Supplier"].dropna().unique())
        supplier = st.selectbox("Origin / ILN-Supplier", ["All"] + suppliers)

    filtered = flows.copy()
    if pol_country != "All":
        filtered = filtered[filtered["POL Country"] == pol_country]
    if pod_country != "All":
        filtered = filtered[filtered["POD Country"] == pod_country]
    if supplier != "All":
        filtered = filtered[filtered["Origin / ILN-Supplier"] == supplier]

    st.subheader(f"Active Flows ({len(filtered)})")
    st.dataframe(
        filtered.sort_values(["POL Country", "POD Country", "Origin / ILN-Supplier"]).reset_index(drop=True),
        width="stretch",
    )


if __name__ == "__main__":
    main()
