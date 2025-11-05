import os
import pandas as pd


def _assert_file(path: str):
    if not os.path.exists(path):
        raise FileNotFoundError(f"Data file not found: {path}")


def load_main_ports(path: str) -> pd.DataFrame:
    _assert_file(path)
    return pd.read_excel(path, sheet_name="MAIN PORTS")


def load_transit_time(path: str) -> pd.DataFrame:
    _assert_file(path)
    return pd.read_excel(path, sheet_name="TRANSITTIME")


def load_horse_puerto(path: str) -> pd.DataFrame:
    _assert_file(path)
    return pd.read_excel(path, sheet_name="HORSE-PUERTO")


def load_cost_per_km(path: str) -> pd.DataFrame:
    _assert_file(path)
    return pd.read_excel(path, sheet_name="COSTPERKM")


def load_packaging(path: str) -> pd.DataFrame:
    _assert_file(path)
    return pd.read_excel(path, sheet_name="PACKAGING")


def map_factory_to_port(df_hp: pd.DataFrame, factory_name: str):
    if df_hp is None or df_hp.empty:
        return None
    try:
        m = df_hp[df_hp["Factory"].astype(str).str.strip().str.lower() == str(factory_name).strip().lower()]
        if m.empty:
            return None
        return m.iloc[0].to_dict()
    except Exception:
        return None


def find_port_by_country(df_hp: pd.DataFrame, country_code: str):
    if df_hp is None or df_hp.empty:
        return None
    try:
        m = df_hp[df_hp["Country Code"].astype(str).str.upper() == str(country_code).strip().upper()]
        if m.empty:
            return None
        return m.iloc[0].to_dict()
    except Exception:
        return None
