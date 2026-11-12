"""Horse Global Logistics - Active Flows Cartography.

Streamlit app to look up active transport flows by origin/destination
(no cost data). Airfreight is sourced from AirCalc and enriched with
cartography links. Overseas is loaded from its workbook; Inland is not
connected yet.
"""

import heapq
from html import escape
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

# Display names for the UN/LOCODE ports and IATA airports used by active flows.
LOCATION_NAMES = {
    "ARBUE": "Buenos Aires",
    "ARROS": "Rosario",
    "BOM": "Chhatrapati Shivaji Maharaj International Airport (Mumbai)",
    "BRPNG": "Paranaguá",
    "CDG": "Paris-Charles de Gaulle Airport",
    "CLSAI": "San Antonio",
    "CNSHA": "Shanghai",
    "COBUN": "Buenaventura",
    "COCTG": "Cartagena",
    "CWB": "Afonso Pena International Airport (Curitiba)",
    "ESALG": "Algeciras",
    "ESVLC": "Valencia",
    "FRLEH": "Le Havre",
    "GRU": "São Paulo/Guarulhos International Airport",
    "HND": "Tokyo International Airport (Haneda)",
    "ICN": "Incheon International Airport (Seoul)",
    "INENR": "Ennore (Kamarajar Port)",
    "INKAT": "Kattupalli",
    "INMAA": "Chennai",
    "INNSA": "Nhava Sheva (Jawaharlal Nehru Port)",
    "IST": "Istanbul Airport",
    "JPNGO": "Nagoya",
    "JPSMZ": "Shimizu",
    "JPYOK": "Yokohama",
    "KRPUS": "Busan",
    "MACAS": "Casablanca",
    "MAD": "Adolfo Suárez Madrid-Barajas Airport",
    "MAPTM": "Tanger Med",
    "MEX": "Mexico City International Airport",
    "MXATM": "Altamira",
    "MXVER": "Veracruz",
    "NGO": "Chubu Centrair International Airport (Nagoya)",
    "NRT": "Narita International Airport (Tokyo)",
    "OPO": "Francisco Sá Carneiro Airport (Porto)",
    "OTP": "Henri Coandă International Airport (Bucharest)",
    "PTLEI": "Leixões",
    "PVG": "Shanghai Pudong International Airport",
    "ROCND": "Constanța",
    "SCL": "Arturo Merino Benítez International Airport (Santiago)",
    "SIKOP": "Koper",
    "TRGEM": "Gemlik",
    "TRIZT": "İzmit",
    "TRYAR": "Yarımca",
    "YEI": "Bursa Yenişehir Airport",
}

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
    "CO": "Colombia",
    "MA": "Morocco",
    "MX": "Mexico",
    "PT": "Portugal",
    "RO": "Romania",
    "SI": "Slovenia",
    "TR": "Turkey",
}

CARTOGRAPHY_COLUMNS = [
    "Type Flow", "Flow", "POL", "POL Country", "POD", "POD Country", "Destin Plant",
    "Origin / ILN-Supplier", "Recipient", "Flow Type", "Commodity",
    "Carrier", "Transit Time (days)", "Customer Lead Time (days)", "Lane Type",
    "Axis", "Carto No.", "Cartography Link", "VTT ID", "Purchase Incoterm",
    "Operations Contact", "Transport Engineering Contact",
]
COUNTRY_CODE_BY_NAME = {
    name.casefold(): code for code, name in COUNTRY_NAME_BY_CODE.items()
}
COUNTRY_CODE_BY_NAME.update({"south korea": "KR", "thailand": "TH"})


@st.cache_data(show_spinner="Loading VTT flows...")
def load_vtt_data() -> pd.DataFrame:
    """Load overseas flows and customer lead times from VTT DATA."""
    if not VTT_DATA_XLSX.exists():
        return pd.DataFrame(
            columns=[
                "ID", "POL", "POD", "Carrier", "Destin Country", "Destin Plant",
                "Shipper", "CLT", "Flow Type", "Commodity", "Transit Time (days)",
            ]
        )

    df = pd.read_excel(VTT_DATA_XLSX)
    df.columns = [str(column).strip() for column in df.columns]
    df["CLT"] = (
        pd.to_numeric(df["14 Rounding"], errors="coerce")
        - pd.to_numeric(df["1 Day Customer Order"], errors="coerce")
        + 1
    )
    for column in ("POL", "POD"):
        df[column] = df[column].astype(str).str.strip().str.upper()
    df["ID"] = df["ID"].map(lambda value: str(value).strip() if pd.notna(value) else None)
    df["Shipper"] = df["Shipper"].map(
        lambda value: str(value).strip() if pd.notna(value) else None
    )
    df["Destin Plant"] = df["Name Destin Site"].map(
        lambda value: normalize_destination_site(value)
    )
    df["Destin Country"] = df["Destin Country"].map(
        lambda value: str(value).strip().title() if pd.notna(value) else None
    )
    df["Flow Type"] = df["Flow Type"].map(
        lambda value: str(value).strip().title() if pd.notna(value) else None
    )
    df["Commodity"] = df["Comodity"].map(
        lambda value: str(value).strip().title() if pd.notna(value) else None
    )
    df["Transit Time (days)"] = pd.to_numeric(df["Transit time"], errors="coerce")
    return df.dropna(subset=["POL", "POD", "Destin Plant"])[
        [
            "ID", "POL", "POD", "Carrier", "Destin Country", "Destin Plant",
            "Shipper", "CLT", "Flow Type", "Commodity", "Transit Time (days)",
        ]
    ].reset_index(drop=True)


def normalize_destination_site(value: object) -> str | None:
    if pd.isna(value):
        return None
    site = " ".join(str(value).split())
    if not site:
        return None
    return {
        "horse brasil": "Horse Brazil",
        "horse valladolid": "Horse Motores",
    }.get(site.casefold(), site)


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
    columns = [
        "AXIS", "Origin Country", "Flow Type", "Shipper/ILN", "Destin Plant",
        "Carto No.", "Cartography Link", "Purchase Incoterm",
    ]
    if not FLOWS_CARTOGRAPHY_XLSX.exists():
        return pd.DataFrame(columns=columns)

    df = pd.read_excel(FLOWS_CARTOGRAPHY_XLSX, sheet_name=CARTOGRAPHY_SHEET)
    df.columns = df.columns.map(lambda column: str(column).strip())
    # The number column header contains a non-ASCII "N°"; find it by its ending.
    number_column = next(c for c in df.columns if str(c).strip().endswith("Carto"))
    df = df.rename(columns={number_column: "Carto No."})[columns].dropna(subset=["AXIS", "Cartography Link"])
    df["Carto No."] = df["Carto No."].astype(str).str.strip()
    df["AXIS"] = df["AXIS"].map(normalize_axis)
    df["Destin Plant"] = df["Destin Plant"].map(normalize_destination_site)
    df["Cartography Link"] = df["Cartography Link"].astype(str).str.strip()
    return df[df["Cartography Link"] != ""]


def matching_cartography_flow(
    axis: object,
    shipper: object,
    destination_plant: object,
    cartography: pd.DataFrame,
) -> pd.DataFrame:
    matches = cartography[cartography["AXIS"] == normalize_axis(axis)]
    shipper_key = normalize_match_value(shipper)
    plant_key = normalize_match_value(normalize_destination_site(destination_plant))
    if not shipper_key or not plant_key:
        return matches.iloc[0:0]

    shipper_matches = matches["Shipper/ILN"].map(normalize_match_value) == shipper_key
    plant_matches = matches["Destin Plant"].map(
        lambda value: normalize_match_value(normalize_destination_site(value))
    ) == plant_key
    return matches[shipper_matches & plant_matches]


def cartography_reclassifies_origin(
    pol: object,
    axis: object,
    shipper: object,
    destination_plant: object,
    cartography: pd.DataFrame,
) -> bool:
    pol_prefix = str(pol).strip().upper()[:2]
    matches = matching_cartography_flow(axis, shipper, destination_plant, cartography)
    origin_codes = [
        COUNTRY_CODE_BY_NAME.get(str(country).strip().casefold())
        for country in matches["Origin Country"].dropna()
    ]
    return bool(origin_codes) and all(
        origin_code is not None and origin_code != pol_prefix
        for origin_code in origin_codes
    )


def build_origin_country_map(
    vtt_pols: pd.Series, cartography: pd.DataFrame
) -> dict[str, str]:
    """Combine VTT POL country codes with cartography origins as a fallback."""
    vtt_prefixes = {
        str(pol).strip().upper()[:2]
        for pol in vtt_pols.dropna()
        if str(pol).strip()
    }
    origin_by_prefix = {
        prefix: COUNTRY_NAME_BY_CODE[prefix]
        for prefix in vtt_prefixes
        if prefix in COUNTRY_NAME_BY_CODE
    }

    if cartography.empty:
        return origin_by_prefix

    cartography_origins = cartography[["AXIS", "Origin Country"]].dropna()
    cartography_origins = cartography_origins.assign(
        _prefix=cartography_origins["AXIS"].str.split(">").str[0]
    )
    for prefix, origins in cartography_origins.groupby("_prefix")["Origin Country"]:
        origin_names = origins.astype(str).str.strip().unique()
        if prefix not in origin_by_prefix and len(origin_names) == 1:
            origin_by_prefix[prefix] = origin_names[0]
    return origin_by_prefix


def resolve_cartography_matches(
    axis: object,
    shipper: object,
    destination_plant: object,
    cartography: pd.DataFrame,
) -> pd.DataFrame:
    axis_key = normalize_axis(axis)
    if not axis_key or cartography.empty:
        return cartography.iloc[0:0]

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
    return matches


def resolve_cartography_link(
    axis: object,
    shipper: object,
    destination_plant: object,
    cartography: pd.DataFrame,
) -> tuple[str | None, str | None]:
    """Returns (Carto No., Cartography Link) when the route resolves unambiguously."""
    matches = resolve_cartography_matches(axis, shipper, destination_plant, cartography)

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
        transit_column = "Total Transit Door-to-Door (days)"
        transit_time = (
            normal[transit_column].iloc[0]
            if not normal.empty
            else g[transit_column].min()
        )

        axis = f"{pol_country}>{pod_country}"
        route_matches = matching_cartography_flow(
            axis, shipper, destination_plant, cartography
        )
        cartography_origins = (
            route_matches["Origin Country"].dropna().astype(str).str.strip().unique()
        )
        purchase_matches = resolve_cartography_matches(
            axis, shipper, destination_plant, cartography
        )
        purchase_incoterms = (
            purchase_matches["Purchase Incoterm"]
            .dropna()
            .astype(str)
            .str.strip()
            .loc[lambda values: values.ne("")]
            .unique()
        )
        purchase_incoterm = purchase_incoterms[0] if len(purchase_incoterms) == 1 else None
        resolved_pol_country = (
            cartography_origins[0] if len(cartography_origins) == 1 else pol_country
        )
        carto_no, carto_link = resolve_cartography_link(
            axis, shipper, destination_plant, cartography
        )

        records.append(
            {
                "Type Flow": "Airfreight",
                "Flow": "Inbound",  # AirCalc currently only carries supplier -> plant parts lanes.
                "POL": pol,
                "POL Country": resolved_pol_country,
                "POD": pod,
                "POD Country": pod_country,
                "Destin Plant": normalize_destination_site(destination_plant),
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
                "Purchase Incoterm": purchase_incoterm,
                "Operations Contact": None,
                "Transport Engineering Contact": None,
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
    df = load_vtt_data()
    if df.empty:
        return pd.DataFrame(columns=CARTOGRAPHY_COLUMNS)

    cartography = load_cartography_flows()
    df["Axis"] = [
        f"{str(pol).strip().upper()[:2]}>"
        f"{COUNTRY_CODE_BY_NAME.get(str(pod_country).casefold(), '')}"
        for pol, pod_country in zip(df["POL"], df["Destin Country"])
    ]
    airfreight = load_airfreight_flows()
    airfreight_keys = {
        (
            normalize_axis(axis),
            normalize_match_value(shipper),
            normalize_match_value(normalize_destination_site(plant)),
        )
        for axis, shipper, plant in zip(
            airfreight["Axis"],
            airfreight["Origin / ILN-Supplier"],
            airfreight["Destin Plant"],
        )
    }
    covered_by_air = [
        (
            normalize_axis(axis),
            normalize_match_value(shipper),
            normalize_match_value(normalize_destination_site(plant)),
        )
        in airfreight_keys
        and matching_cartography_flow(axis, shipper, plant, cartography)["Flow Type"]
        .astype(str).str.strip().str.casefold().eq("emergency").any()
        and cartography_reclassifies_origin(
            pol, axis, shipper, plant, cartography
        )
        for pol, axis, shipper, plant in zip(
            df["POL"], df["Axis"], df["Shipper"], df["Destin Plant"]
        )
    ]
    df = df.loc[~pd.Series(covered_by_air, index=df.index)].copy()
    if df.empty:
        return pd.DataFrame(columns=CARTOGRAPHY_COLUMNS)

    df["Type Flow"] = "Overseas"
    df["Purchase Incoterm"] = None
    df["Operations Contact"] = None
    df["Transport Engineering Contact"] = None
    origin_country_by_prefix = build_origin_country_map(df["POL"], cartography)
    df["POL Country"] = df["POL"].str[:2].map(origin_country_by_prefix)
    df["POD Country"] = df["Destin Country"]
    df["Flow"] = df["Destin Plant"].map(
        lambda site: "Inbound" if site.casefold().startswith("horse ") else "Outbound"
    )
    df["Origin / ILN-Supplier"] = df["Shipper"]
    df["Recipient"] = df["Destin Plant"]
    df["VTT ID"] = df["ID"]
    df["Customer Lead Time (days)"] = df["CLT"].map(
        lambda value: str(int(value)) if pd.notna(value) else None
    )
    df["Lane Type"] = None
    carto = [
        resolve_cartography_link(axis, origin, None, cartography)
        for axis, origin in zip(df["Axis"], df["Origin / ILN-Supplier"])
    ]
    df["Carto No."] = [number for number, _ in carto]
    df["Cartography Link"] = [link for _, link in carto]
    df["Carrier"] = df["Carrier"].fillna("Not available")
    return df[CARTOGRAPHY_COLUMNS]


def load_inland_flows() -> pd.DataFrame:
    # TODO: wire up to the Inland flows database once available. POL/POD are N/A for Inland.
    return pd.DataFrame(columns=CARTOGRAPHY_COLUMNS)


def ensure_flow_schema(flows: pd.DataFrame) -> pd.DataFrame:
    """Add missing columns so partial flow sources retain the expected schema."""
    for column in CARTOGRAPHY_COLUMNS:
        if column not in flows.columns:
            flows[column] = None
    return flows


def available_destination_plants(flows: pd.DataFrame) -> list[str]:
    return sorted(
        {
            plant.strip()
            for value in flows["Destin Plant"].dropna()
            for plant in str(value).split(" | ")
            if plant.strip()
        },
        key=str.casefold,
    )


def load_active_flows() -> pd.DataFrame:
    # Recombine cached sources on each rerun so changes to either source are reflected.
    flows = pd.concat(
        [load_airfreight_flows(), load_overseas_flows()],
        ignore_index=True,
    )
    flows = ensure_flow_schema(flows)
    for column in ("POL Country", "POD Country"):
        flows[column] = flows[column].map(normalize_country)
    return flows


def main() -> None:
    st.set_page_config(page_title="Horse GL - Active Flows", page_icon="🌍", layout="wide")
    st.title("🌍 Horse Global Logistics - Active Flows Cartography")
    st.caption(
        "Select an origin country and destination plant to explore available transport flows and route details."
    )

    flows = load_active_flows()
    if not FLOWS_CARTOGRAPHY_XLSX.exists():
        st.warning("The Flows Cartography file is unavailable; cartography links cannot be loaded.")

    if flows.empty:
        st.info(
            "No airfreight or overseas flows are available from the connected data files."
        )
        return

    cartography = load_cartography_flows()
    pol_countries = set(flows["POL Country"].dropna().map(normalize_country))
    pol_countries.update(
        cartography["Origin Country"]
        .dropna()
        .map(lambda country: normalize_country(str(country).strip()))
    )
    pol_countries = sorted(
        (country for country in pol_countries if country),
        key=str.casefold,
    )
    col1, col2 = st.columns(2)
    with col1:
        pol_country = st.selectbox("Origin Country", ["All"] + pol_countries)

    origin_filtered = flows
    if pol_country != "All":
        origin_filtered = flows[flows["POL Country"] == pol_country]
    destination_sites = available_destination_plants(origin_filtered)
    with col2:
        destination_site = st.selectbox("Destin Plant", ["All"] + destination_sites)

    filtered = origin_filtered.copy()
    if destination_site != "All":
        filtered = filtered[
            filtered["Destin Plant"].fillna("").map(
                lambda value: destination_site in value.split(" | ")
            )
        ]

    if filtered.empty:
        st.info("No flows match the selected origin country and destination site.")
        return

    filtered = filtered.sort_values(["POL Country", "POD Country", "Origin / ILN-Supplier"]).reset_index(drop=True)

    flow_sections = [
        ("🚢 Overseas", filtered[filtered["Type Flow"] == "Overseas"], "ovs"),
        ("✈️ Airfreight", filtered[filtered["Type Flow"] == "Airfreight"], "air"),
        ("🚚 Inland", filtered[filtered["Type Flow"] == "Inland"], "inland"),
    ]
    available_sections = [
        (label, section_flows, key)
        for label, section_flows, key in flow_sections
        if not section_flows.empty
    ]
    tabs = st.tabs([label for label, _, _ in available_sections])
    for tab, (_, section_flows, key) in zip(tabs, available_sections):
        with tab:
            if key == "inland":
                st.info("Inland flows: source not connected yet.")
            else:
                render_flow_tab(section_flows, key)


def render_flow_tab(flows: pd.DataFrame, key: str) -> None:
    flows = flows.reset_index(drop=True)
    state_key = f"route_{key}"
    render_cards(flows, state_key)
    choice = st.session_state.get(state_key)
    if choice is not None and 0 <= choice < len(flows):
        render_route_map(flows.iloc[[choice]])


def render_route_map(flows: pd.DataFrame) -> None:
    routes = flows[flows["POL"].isin(PORT_COORDINATES) & flows["POD"].isin(PORT_COORDINATES)]
    if routes.empty:
        st.caption("No coordinates available to draw this route.")
        return

    records = []
    for _, flow in routes.iterrows():
        label = f"{format_location(flow['POL'])} → {format_location(flow['POD'])}"
        color = TYPE_COLORS[flow["Type Flow"]]
        if flow["Type Flow"] == "Airfreight":
            paths = dashed_segments(
                list(PORT_COORDINATES[flow["POL"]]),
                list(PORT_COORDINATES[flow["POD"]]),
            )
        else:
            paths = [sea_path(flow["POL"], flow["POD"])]
        records.extend(
            {"route": label, "path": path, "color": color} for path in paths
        )

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


def na(value: object) -> str:
    return "Not available" if value is None or pd.isna(value) else str(value)


def format_days(value: object) -> str:
    if value is None or pd.isna(value):
        return "Not available"
    return f"{float(value):.0f}"


def format_location(code: object) -> str:
    if code is None or pd.isna(code):
        return "Not available"
    code = str(code).strip()
    name = LOCATION_NAMES.get(code.upper())
    return f"{code} ({name})" if name else code


def render_cards(flows: pd.DataFrame, state_key: str) -> None:
    if flows.empty:
        st.info("No flows of this type for the selected countries.")
        return

    columns = st.columns(2)
    for position, (_, flow) in enumerate(flows.iterrows()):
        with columns[position % 2].container(border=True):
            st.markdown(
                f"**{format_location(flow['POL'])} → {format_location(flow['POD'])}**  \n"
                f"{flow['POL Country']} → {flow['POD Country']}"
            )
            description = (
                f"{na(flow['Origin / ILN-Supplier'])} ➜ "
                f"{na(flow['Recipient'])} · {na(flow['Carrier'])}"
            )
            st.markdown(
                f"<div style='color: #1976d2;'>{escape(description)}</div>",
                unsafe_allow_html=True,
            )
            metrics = st.columns(2)
            metrics[0].metric("Transit (days)", format_days(flow["Transit Time (days)"]))
            if flow["Type Flow"] == "Overseas":
                metrics[1].metric("Customer lead time (days)", na(flow["Customer Lead Time (days)"]))
            tags = [flow["Lane Type"], flow["Flow"], flow["Flow Type"], flow["Commodity"]]
            visible_tags = [str(tag) for tag in tags if tag is not None and pd.notna(tag)]
            if visible_tags:
                tag_markup = " ".join(
                    (
                        "<code style='color: #000; "
                        "background-color: var(--secondary-background-color, #f0f2f6); "
                        "padding: 0.2em 0.4em; border-radius: 0.25rem;'>"
                        f"{escape(tag)}</code>"
                    )
                    for tag in visible_tags
                )
                st.markdown(tag_markup, unsafe_allow_html=True)
            if flow["Type Flow"] == "Overseas":
                cartography_column, vtt_column = st.columns(2)
                with cartography_column:
                    if pd.notna(flow["Cartography Link"]):
                        st.markdown(f"Cartography: [{flow['Carto No.']}]({flow['Cartography Link']})")
                    else:
                        st.caption("Cartography: Not available")
                ids = [i.strip() for i in str(flow["VTT ID"]).split(",")] if pd.notna(flow["VTT ID"]) else []
                with vtt_column:
                    if ids:
                        links = ", ".join(f"[{i}]({VTT_APP_URL}?id={quote(i)})" for i in ids)
                        st.markdown(f"VTT ID: {links}")
                    else:
                        st.caption("VTT ID: Not available")
            st.markdown(
                "<div style='color: #000;'>"
                f"Purchase Incoterm: {escape(na(flow['Purchase Incoterm']))}<br>"
                f"Operations Contact: {escape(na(flow['Operations Contact']))}<br>"
                "Transport Engineering Contact: "
                f"{escape(na(flow['Transport Engineering Contact']))}"
                "</div>",
                unsafe_allow_html=True,
            )
            if st.button(
                "📍 Showing on map" if st.session_state.get(state_key) == position else "🗺️ Show on map",
                key=f"{state_key}_btn_{position}",
                type="primary" if st.session_state.get(state_key) == position else "secondary",
                width="stretch",
            ):
                st.session_state[state_key] = position
                st.rerun()


if __name__ == "__main__":
    main()
