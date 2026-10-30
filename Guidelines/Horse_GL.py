"""Horse Global Logistics - Active Flows Cartography.

Streamlit app to look up active transport flows by origin/destination
(no cost data). Airfreight is sourced from AirCalc and enriched with
cartography links. Overseas is loaded from its workbook; Inland is not
connected yet.
"""

import heapq
import math
from pathlib import Path
from urllib.parse import quote

import pandas as pd
import pydeck as pdk
import streamlit as st

AIRCALC_XLSM = Path(__file__).parent.parent / "AirCalc" / "Airfreights Rates Calculator v2 Tender 2026.xlsm"
CONSOLIDATED_SHEET = "1_Consolidated_AIR"
HEADER_ROW = 3  # 0-indexed row of the header inside the consolidated sheet (Excel row 4).

FLOWS_CARTOGRAPHY_XLSX = Path(__file__).parent / "Flows Cartography.xlsx"
CARTOGRAPHY_SHEET = "AIRFREIGHTS Invoice flow"

OVERSEAS_XLSX = Path(__file__).parent / "Overseas.xlsx"
OVERSEAS_SHEET = "budget 2026-2027"

VTT_DATA_XLSX = Path(__file__).parent.parent / "VTT Tool" / "VTT DATA.xlsx"
VTT_APP_URL = "https://horsetptengtools.streamlit.app/"

# Destination country -> recipient Horse plant. Extend as new lanes/plants are onboarded.
RECIPIENT_BY_COUNTRY = {
    "CL": "900095 - Horse Chile",
    "RO": "900125 - Horse Romania",
    "TR": "900144 - Horse Turkey",
    "BR": "900186 - Horse Brasil",
    "ES": "910177 - Horse Valladolid",
}

TYPE_COLORS = {"Airfreight": [0, 120, 255], "Overseas": [255, 140, 0]}

# (longitude, latitude) of every airport/port code used by the flows.
PORT_COORDINATES = {
    "ARROS": (-60.64, -32.95), "BOM": (72.87, 19.09), "BRPNG": (-48.51, -25.52),
    "CDG": (2.55, 49.01), "CLSAI": (-71.61, -33.59), "CNSHA": (121.80, 31.23),
    "CWB": (-49.17, -25.53), "ESALG": (-5.45, 36.13), "ESVLC": (-0.32, 39.45),
    "FRLEH": (0.11, 49.48), "GRU": (-46.47, -23.43), "HND": (139.78, 35.55),
    "ICN": (126.45, 37.46), "INENR": (80.33, 13.25), "INKAT": (80.34, 13.31),
    "INMAA": (80.30, 13.08), "INNSA": (72.95, 18.95), "IST": (28.75, 41.26),
    "JPNGO": (136.88, 35.05), "JPYOK": (139.65, 35.45), "KRPUS": (129.04, 35.10),
    "MAD": (-3.57, 40.49), "MAPTM": (-5.50, 35.89), "MEX": (-99.07, 19.44),
    "MXVER": (-96.13, 19.20), "NGO": (136.81, 34.86), "NRT": (140.39, 35.77),
    "OPO": (-8.68, 41.24), "OTP": (26.10, 44.57), "PTLEI": (-8.70, 41.19),
    "PVG": (121.81, 31.14), "ROCND": (28.65, 44.17), "SCL": (-70.79, -33.39),
    "TRGEM": (29.15, 40.43), "TRIZT": (29.92, 40.77), "YEI": (29.55, 40.25),
}

# Simplified sea-lane network (lon, lat) so Overseas routes follow the water instead of cutting over land.
SEA_NODES = {
    "Channel": (-1, 49.8), "ChannelW": (-5.5, 48.5), "Biscay": (-8, 46), "IbW": (-11, 40),
    "IbS": (-9.5, 36.2), "Gib": (-5.6, 35.95), "MedW": (1, 38.3), "MedC": (14, 36.5),
    "MedE": (28, 34.5), "Dardanelles": (26, 40), "Marmara": (28.5, 40.8),
    "Bosphorus": (29.1, 41.3), "BlackSea": (30, 43.5), "Suez": (32.6, 29.9),
    "RedSea": (38, 20), "BabEl": (43.4, 12.6), "Aden": (50, 12), "Arabian": (65, 12),
    "ArabianN": (70, 15), "SriLanka": (80, 5.5), "BoB": (82, 12), "Andaman": (94, 6),
    "Malacca": (101, 2.5), "Singapore": (104, 1.2), "SCS": (110, 8), "Taiwan": (119, 22),
    "ECS": (124, 28), "ECSEast": (130, 30), "JapanS": (138, 33), "ShanghaiOff": (123.5, 30),
    "KoreaS": (129.5, 33), "Canary": (-20, 28), "AtlM": (-30, 10), "BrNE": (-33, -6),
    "BrE": (-37, -20), "BrSE": (-40, -25), "BrS": (-46, -27.5), "RdlP": (-54, -35.5),
    "SAtl": (-30, -42), "Patagonia": (-58, -47), "CapeHorn": (-67, -57.5),
    "ChilePac": (-76, -50), "ChileN": (-76, -40), "Cape": (18, -37), "IndS": (60, -25),
    "GulfMex": (-94, 21), "Yucatan": (-86, 22), "FlStraits": (-80, 24),
    "Caribbean": (-65, 20), "NAtlEq": (-45, 5),
}
SEA_EDGES = [
    ("Channel", "ChannelW"), ("ChannelW", "Biscay"), ("Biscay", "IbW"), ("IbW", "IbS"),
    ("IbS", "Gib"), ("Gib", "MedW"), ("MedW", "MedC"), ("MedC", "MedE"), ("MedE", "Dardanelles"),
    ("Dardanelles", "Marmara"), ("Marmara", "Bosphorus"), ("Bosphorus", "BlackSea"),
    ("MedE", "Suez"), ("Suez", "RedSea"), ("RedSea", "BabEl"), ("BabEl", "Aden"),
    ("Aden", "Arabian"), ("Arabian", "ArabianN"), ("Arabian", "SriLanka"), ("SriLanka", "BoB"),
    ("SriLanka", "Andaman"), ("Andaman", "Malacca"), ("Malacca", "Singapore"),
    ("Singapore", "SCS"), ("SCS", "Taiwan"), ("Taiwan", "ECS"), ("ECS", "ShanghaiOff"),
    ("ECS", "ECSEast"), ("ECSEast", "KoreaS"), ("ECSEast", "JapanS"), ("IbS", "Canary"),
    ("Canary", "AtlM"), ("AtlM", "BrNE"), ("BrNE", "BrE"), ("BrE", "BrSE"), ("BrSE", "BrS"),
    ("BrS", "RdlP"), ("RdlP", "Patagonia"), ("Patagonia", "CapeHorn"), ("CapeHorn", "ChilePac"),
    ("ChilePac", "ChileN"), ("RdlP", "SAtl"), ("CapeHorn", "SAtl"), ("SAtl", "Cape"),
    ("Cape", "IndS"), ("IndS", "SriLanka"), ("GulfMex", "Yucatan"), ("Yucatan", "FlStraits"),
    ("FlStraits", "Caribbean"), ("Caribbean", "NAtlEq"), ("NAtlEq", "BrNE"),
]
# Port code -> first sea node reached from the port.
PORT_SEA_ACCESS = {
    "FRLEH": "Channel", "PTLEI": "IbW", "ESALG": "Gib", "MAPTM": "Gib", "ESVLC": "MedW",
    "TRIZT": "Marmara", "TRGEM": "Marmara", "ROCND": "BlackSea", "INNSA": "ArabianN",
    "INMAA": "BoB", "INENR": "BoB", "INKAT": "BoB", "CNSHA": "ShanghaiOff",
    "KRPUS": "KoreaS", "JPYOK": "JapanS", "JPNGO": "JapanS", "BRPNG": "BrS",
    "ARROS": "RdlP", "CLSAI": "ChileN", "MXVER": "GulfMex",
}


def sea_path(pol: str, pod: str) -> list[list[float]]:
    # Dijkstra over the simplified sea network; falls back to a straight line if a port is unmapped.
    start, end = PORT_SEA_ACCESS.get(pol), PORT_SEA_ACCESS.get(pod)
    origin, target = list(PORT_COORDINATES[pol]), list(PORT_COORDINATES[pod])
    if start is None or end is None:
        return [origin, target]

    graph: dict[str, list[tuple[str, float]]] = {name: [] for name in SEA_NODES}
    for a, b in SEA_EDGES:
        distance = math.dist(SEA_NODES[a], SEA_NODES[b])
        graph[a].append((b, distance))
        graph[b].append((a, distance))

    best = {start: 0.0}
    previous: dict[str, str] = {}
    queue = [(0.0, start)]
    while queue:
        cost, node = heapq.heappop(queue)
        if node == end:
            break
        if cost > best.get(node, float("inf")):
            continue
        for neighbour, distance in graph[node]:
            new_cost = cost + distance
            if new_cost < best.get(neighbour, float("inf")):
                best[neighbour] = new_cost
                previous[neighbour] = node
                heapq.heappush(queue, (new_cost, neighbour))
    if end not in best:
        return [origin, target]

    nodes = [end]
    while nodes[-1] != start:
        nodes.append(previous[nodes[-1]])
    return [origin] + [list(SEA_NODES[name]) for name in reversed(nodes)] + [target]


def dashed_segments(src: list[float], dst: list[float], dash_degrees: float = 1.5) -> list[list[list[float]]]:
    # Splits a straight line into short on/off pieces to draw it dotted.
    steps = max(int(math.dist(src, dst) / dash_degrees), 2)
    points = [[src[0] + (dst[0] - src[0]) * i / steps, src[1] + (dst[1] - src[1]) * i / steps] for i in range(steps + 1)]
    return [[points[i], points[i + 1]] for i in range(0, steps, 2)]


COUNTRY_NAME_BY_CODE = {
    "AR": "Argentina",
    "BR": "Brazil",
    "CL": "Chile",
    "CN": "China",
    "ES": "Spain",
    "FR": "France",
    "IN": "India",
    "JP": "Japan",
    "KR": "Korea",
    "MX": "Mexico",
    "PT": "Portugal",
    "RO": "Romania",
    "TR": "Turkey",
}

CARTOGRAPHY_COLUMNS = [
    "Type Flow", "Flow", "POL", "POL Country", "POD", "POD Country",
    "Origin / ILN-Supplier", "Recipient", "Flow Type", "Commodity",
    "Carrier", "Transit Time (days)", "Customer Lead Time (days)", "Lane Type",
    "Axis", "Carto No.", "Cartography Link", "VTT ID",
]
COUNTRY_CODE_BY_NAME = {name.casefold(): code for code, name in COUNTRY_NAME_BY_CODE.items()}


@st.cache_data(show_spinner="Loading VTT customer lead times...")
def load_vtt_customer_lead_times() -> pd.DataFrame:
    """CLT per VTT row, computed as in VTT2.py: '14 Rounding' - '1 Day Customer Order' + 1."""
    if not VTT_DATA_XLSX.exists():
        return pd.DataFrame(columns=["ID", "POL", "POD", "Shipper", "CLT"])

    df = pd.read_excel(VTT_DATA_XLSX)
    df.columns = [str(column).strip() for column in df.columns]
    df["CLT"] = (
        pd.to_numeric(df["14 Rounding"], errors="coerce")
        - pd.to_numeric(df["1 Day Customer Order"], errors="coerce")
        + 1
    )
    for column in ("POL", "POD"):
        df[column] = df[column].astype(str).str.strip().str.upper()
    df["Shipper"] = df["Shipper"].astype(str).str.strip().str.casefold()
    df["ID"] = df["ID"].astype(str).str.strip()
    return df.dropna(subset=["CLT"])[["ID", "POL", "POD", "Shipper", "CLT"]]


def resolve_customer_lead_time(vtt: pd.DataFrame, pol, pod, origin) -> tuple[str | None, str | None]:
    """POL + POD + Shipper match against the VTT.

    Returns (CLT, VTT IDs). The CLT is None when the VTT holds conflicting values; the IDs list every matching row.
    """
    origin_key = normalize_match_value(origin)
    # Horse/Renault plants appear in the VTT as "ILN <site>" shippers; named suppliers are matched by name.
    if origin_key in ("horse plant", "renault iln"):
        shipper_match = vtt["Shipper"].str.startswith("iln")
    else:
        shipper_match = vtt["Shipper"].str.contains(origin_key, regex=False)
    matches = vtt[
        (vtt["POL"] == str(pol).strip().upper())
        & (vtt["POD"] == str(pod).strip().upper())
        & shipper_match
    ]
    if matches.empty:
        return None, None
    values = matches["CLT"].unique()
    clt = str(int(values[0])) if len(values) == 1 else None
    return clt, ", ".join(matches["ID"])


def normalize_axis(value: object) -> str:
    if pd.isna(value):
        return ""
    return ">".join(part.strip().upper() for part in str(value).split(">"))


def normalize_match_value(value: object) -> str:
    if pd.isna(value):
        return ""
    return " ".join(str(value).split()).casefold()


def normalize_country(value: object) -> object:
    if pd.isna(value):
        return value
    country = str(value).strip()
    return COUNTRY_NAME_BY_CODE.get(country.upper(), country)


@st.cache_data(show_spinner="Loading cartography links...")
def load_cartography_flows() -> pd.DataFrame:
    columns = ["AXIS", "Shipper/ILN", "Destin Plant", "Carto No.", "Cartography Link"]
    if not FLOWS_CARTOGRAPHY_XLSX.exists():
        return pd.DataFrame(columns=columns)

    df = pd.read_excel(FLOWS_CARTOGRAPHY_XLSX, sheet_name=CARTOGRAPHY_SHEET)
    # The number column header contains a non-ASCII "N°"; find it by its ending.
    number_column = next(c for c in df.columns if str(c).strip().endswith("Carto"))
    df = df.rename(columns={number_column: "Carto No."})[columns].dropna(subset=["AXIS", "Cartography Link"])
    df["Carto No."] = df["Carto No."].astype(str).str.strip()
    df["AXIS"] = df["AXIS"].map(normalize_axis)
    df["Cartography Link"] = df["Cartography Link"].astype(str).str.strip()
    return df[df["Cartography Link"] != ""]


def resolve_cartography_link(
    axis: object,
    shipper: object,
    destination_plant: object,
    cartography: pd.DataFrame,
) -> tuple[str | None, str | None]:
    """Returns (Carto No., Cartography Link) for the Axis, narrowed by shipper and plant; (None, None) if ambiguous."""
    axis_key = normalize_axis(axis)
    if not axis_key or cartography.empty:
        return None, None

    matches = cartography[cartography["AXIS"] == axis_key]
    for column, value in (
        ("Shipper/ILN", shipper),
        ("Destin Plant", destination_plant),
    ):
        value_key = normalize_match_value(value)
        if value_key:
            keys = matches[column].map(normalize_match_value)
            if value_key in ("horse plant", "renault iln"):
                narrowed = matches[keys.str.startswith("iln")]
            else:
                narrowed = matches[(keys == value_key) | keys.str.contains(value_key, regex=False)]
            if not narrowed.empty:
                matches = narrowed

    matches = matches.drop_duplicates(["Carto No.", "Cartography Link"])
    return (matches["Carto No."].iloc[0], matches["Cartography Link"].iloc[0]) if len(matches) == 1 else (None, None)


@st.cache_data(show_spinner="Loading airfreight flows...")
def load_airfreight_flows() -> pd.DataFrame:
    if not AIRCALC_XLSM.exists():
        return pd.DataFrame(columns=CARTOGRAPHY_COLUMNS)

    cartography = load_cartography_flows()
    df = pd.read_excel(AIRCALC_XLSM, sheet_name=CONSOLIDATED_SHEET, header=HEADER_ROW)
    df = df.dropna(subset=["Country Code", "Country Code Destin"])

    # One row per lane, ignoring the Economy/Normal/Express duplication used for pricing.
    group_cols = [
        "Shipper/ILN", "Type of Flow", "Country Code", "Origin Airport",
        "Country Code Destin", "Destination Airport", "Provider_Name",
        "Destination Plant",
    ]

    df = df[df["Type of Flow"].astype(str).str.strip().str.title() == "Main"]

    records = []
    for keys, g in df.groupby(group_cols, dropna=False):
        shipper, lane_type, pol_country, pol, pod_country, pod, carrier, destination_plant = keys

        # Prefer "Normal" service level as the representative transit time, else the fastest one.
        normal = g[g["Service Level (Normal, Economy and Express)"].str.title() == "Normal"]
        transit_time = normal["Total TT (days)"].iloc[0] if not normal.empty else g["Total TT (days)"].min()

        carto_no, carto_link = resolve_cartography_link(
            f"{pol_country}>{pod_country}", shipper, destination_plant, cartography
        )
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
                "Customer Lead Time (days)": None,
                "Lane Type": lane_type,
                "_plant": destination_plant,
                "Axis": f"{pol_country}>{pod_country}",
                "Carto No.": carto_no,
                "Cartography Link": carto_link,
                "VTT ID": None,
            }
        )

    flows = pd.DataFrame(records, columns=[*CARTOGRAPHY_COLUMNS, "_plant"])
    # A single Airfreight line per flow (shipper + origin/destination country + plant): keep the fastest.
    flows = (
        flows.sort_values("Transit Time (days)", na_position="last")
        .drop_duplicates(["Origin / ILN-Supplier", "POL Country", "POD Country", "_plant"])
        .sort_index()
    )
    return flows[CARTOGRAPHY_COLUMNS].reset_index(drop=True)


@st.cache_data(show_spinner="Loading overseas flows...")
def load_overseas_flows() -> pd.DataFrame:
    if not OVERSEAS_XLSX.exists():
        return pd.DataFrame(columns=CARTOGRAPHY_COLUMNS)

    df = pd.read_excel(OVERSEAS_XLSX, sheet_name=OVERSEAS_SHEET)
    # The sheet has a continent-summary block below the real data; keep only actual flow rows.
    df = df.dropna(subset=["Flow", "POL", "POD"])
    df = df.rename(columns={"Origin": "Origin / ILN-Supplier", "Transit time": "Transit Time (days)"})
    df["Type Flow"] = "Overseas"
    vtt = load_vtt_customer_lead_times()
    cartography = load_cartography_flows()
    clt_and_ids = [
        resolve_customer_lead_time(vtt, pol, pod, origin)
        for pol, pod, origin in zip(df["POL"], df["POD"], df["Origin / ILN-Supplier"])
    ]
    df["Customer Lead Time (days)"] = [clt for clt, _ in clt_and_ids]
    df["VTT ID"] = [ids for _, ids in clt_and_ids]
    df["Lane Type"] = None
    # Airfreight is the emergency mode of the overseas flows, so they share the cartography of the same Axis.
    df["Axis"] = [
        f"{COUNTRY_CODE_BY_NAME.get(str(pol).strip().casefold(), '')}>{COUNTRY_CODE_BY_NAME.get(str(pod).strip().casefold(), '')}"
        for pol, pod in zip(df["POL Country"], df["POD Country"])
    ]
    carto = [
        resolve_cartography_link(axis, origin, None, cartography)
        for axis, origin in zip(df["Axis"], df["Origin / ILN-Supplier"])
    ]
    df["Carto No."] = [number for number, _ in carto]
    df["Cartography Link"] = [link for _, link in carto]
    return df[CARTOGRAPHY_COLUMNS]


def load_inland_flows() -> pd.DataFrame:
    # TODO: wire up to the Inland flows database once available. POL/POD are N/A for Inland.
    return pd.DataFrame(columns=CARTOGRAPHY_COLUMNS)


@st.cache_data(show_spinner="Loading active flows...")
def load_active_flows() -> pd.DataFrame:
    # Combines flows and guarantees every expected column exists, even with stale cached sub-results.
    flows = pd.concat(
        [load_airfreight_flows(), load_overseas_flows()],
        ignore_index=True,
    )
    for column in ("POL Country", "POD Country"):
        flows[column] = flows[column].map(normalize_country)
    for column in ("Customer Lead Time (days)", "Carto No.", "VTT ID"):
        if column not in flows.columns:
            flows[column] = None
    return flows


def main() -> None:
    st.set_page_config(page_title="Horse GL - Active Flows", page_icon="🌍", layout="wide")
    st.title("🌍 Horse Global Logistics - Active Flows Cartography")
    st.caption("Select origin and destination countries to see available flows and transit times.")

    flows = load_active_flows()
    if not FLOWS_CARTOGRAPHY_XLSX.exists():
        st.warning("The Flows Cartography file is unavailable; cartography links cannot be loaded.")

    if flows.empty:
        st.info(
            "No airfreight or overseas flows are available from the connected data files."
        )
        return

    st.subheader("Filters")
    col1, col2 = st.columns(2)
    with col1:
        pol_countries = sorted(flows["POL Country"].dropna().unique())
        pol_country = st.selectbox("Origin Country", ["All"] + pol_countries)
    with col2:
        pod_countries = sorted(flows["POD Country"].dropna().unique())
        pod_country = st.selectbox("Destination Country", ["All"] + pod_countries)

    filtered = flows.copy()
    if pol_country != "All":
        filtered = filtered[filtered["POL Country"] == pol_country]
    if pod_country != "All":
        filtered = filtered[filtered["POD Country"] == pod_country]

    without_link = filtered["Cartography Link"].isna()
    if without_link.any():
        st.caption(
            f"{without_link.sum()} flow(s) have no unambiguous Cartography Link for their Axis."
        )

    if filtered.empty:
        st.info("No flows match the selected origin and destination countries.")
        return

    filtered = filtered.sort_values(["POL Country", "POD Country", "Origin / ILN-Supplier"]).reset_index(drop=True)

    tab_air, tab_ovs, tab_inland = st.tabs(["✈️ Airfreight", "🚢 Overseas", "🚚 Inland"])
    with tab_air:
        render_flow_tab(filtered[filtered["Type Flow"] == "Airfreight"], "air")
    with tab_ovs:
        render_flow_tab(filtered[filtered["Type Flow"] == "Overseas"], "ovs")
    with tab_inland:
        st.info("Inland flows: source not connected yet.")

    with st.expander("Full table"):
        st.dataframe(
            filtered,
            width="stretch",
            column_config={
                "Cartography Link": st.column_config.LinkColumn(
                    "Cartography Link",
                    display_text="Open cartography",
                ),
            },
        )


def render_flow_tab(flows: pd.DataFrame, key: str) -> None:
    flows = flows.reset_index(drop=True)
    state_key = f"route_{key}"
    render_cards(flows, state_key)
    if flows.empty:
        return

    st.subheader("Route map")
    choice = st.session_state.get(state_key)
    if choice is None or choice >= len(flows):
        st.caption("Click “Show on map” on a flow card to see its route here.")
        return

    selected = flows.iloc[[choice]]
    flow = selected.iloc[0]
    label = f"{flow['POL']} → {flow['POD']} · {flow['Origin / ILN-Supplier']} · {na(flow['Carrier'])}"
    color = TYPE_COLORS[flow["Type Flow"]]
    style = "straight dotted line" if flow["Type Flow"] == "Airfreight" else "solid line along sea lanes"
    st.markdown(
        f"<span style='color:rgb({color[0]},{color[1]},{color[2]})'>&#9473;&#9473;</span> "
        f"**{label}** ({style})",
        unsafe_allow_html=True,
    )
    render_route_map(selected)


def render_route_map(flows: pd.DataFrame) -> None:
    routes = flows[flows["POL"].isin(PORT_COORDINATES) & flows["POD"].isin(PORT_COORDINATES)]
    if routes.empty:
        st.caption("No coordinates available to draw the routes.")
        return

    records = []
    for _, flow in routes.drop_duplicates(["POL", "POD", "Type Flow"]).iterrows():
        label = f"{flow['POL']} → {flow['POD']} ({flow['Type Flow']})"
        color = TYPE_COLORS[flow["Type Flow"]]
        if flow["Type Flow"] == "Airfreight":
            paths = dashed_segments(list(PORT_COORDINATES[flow["POL"]]), list(PORT_COORDINATES[flow["POD"]]))
        else:
            paths = [sea_path(flow["POL"], flow["POD"])]
        records.extend({"route": label, "path": path, "color": color} for path in paths)

    layer = pdk.Layer(
        "PathLayer",
        data=pd.DataFrame(records),
        get_path="path",
        get_color="color",
        get_width=1,
        width_min_pixels=3,
        pickable=True,
    )
    st.pydeck_chart(
        pdk.Deck(
            layers=[layer],
            initial_view_state=pdk.ViewState(latitude=25, longitude=10, zoom=1, pitch=0),
            tooltip={"text": "{route}"},
        ),
        height=420,
    )
    unmapped = sorted((set(flows["POL"]) | set(flows["POD"])) - set(PORT_COORDINATES))
    if len(routes) < len(flows):
        st.caption(f"{len(flows) - len(routes)} flow(s) not drawn: missing coordinates ({', '.join(unmapped)}).")


def na(value: object) -> str:
    return "Not available" if value is None or pd.isna(value) else str(value)


def render_cards(flows: pd.DataFrame, state_key: str) -> None:
    if flows.empty:
        st.info("No flows of this type for the selected countries.")
        return

    columns = st.columns(2)
    for position, (_, flow) in enumerate(flows.iterrows()):
        is_selected = st.session_state.get(state_key) == position
        with columns[position % 2].container(border=True):
            st.markdown(f"**{flow['POL']} → {flow['POD']}**  \n{flow['POL Country']} → {flow['POD Country']}")
            st.caption(f"{na(flow['Origin / ILN-Supplier'])} ➜ {na(flow['Recipient'])} · {na(flow['Carrier'])}")
            metrics = st.columns(2)
            metrics[0].metric("Transit (days)", na(flow["Transit Time (days)"]))
            if flow["Type Flow"] == "Overseas":
                metrics[1].metric("Customer lead time (days)", na(flow["Customer Lead Time (days)"]))
                ids = [i.strip() for i in str(flow["VTT ID"]).split(",")] if pd.notna(flow["VTT ID"]) else []
                if ids:
                    links = ", ".join(f"[{i}]({VTT_APP_URL}?id={quote(i)})" for i in ids)
                    st.markdown(f"VTT ID: {links}")
                else:
                    st.caption("VTT ID: Not available")
            tags = [flow["Lane Type"], flow["Flow"], flow["Flow Type"], flow["Commodity"]]
            st.markdown(" ".join(f"`{tag}`" for tag in tags if tag is not None and pd.notna(tag)))
            if pd.notna(flow["Cartography Link"]):
                st.markdown(f"Cartography: [{flow['Carto No.']}]({flow['Cartography Link']})")
            else:
                st.caption("Cartography: Not available")
            if st.button(
                "📍 Showing on map" if is_selected else "🗺️ Show on map",
                key=f"{state_key}_btn_{position}",
                type="primary" if is_selected else "secondary",
                width="stretch",
            ):
                st.session_state[state_key] = position
                st.rerun()


if __name__ == "__main__":
    main()
