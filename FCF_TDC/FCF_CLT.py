"""Streamlit dashboard for maritime lead-time and inventory cost analysis."""

from __future__ import annotations

import math
import re
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from excel_export import build_workbook
from calculations import (
    DEFAULT_DAYS_PER_YEAR,
    ScenarioInputs,
    calculate_comparison,
    calculate_scenario,
)


st.set_page_config(
    page_title="Maritime CLT & Working Capital Optimizer",
    page_icon="🚢",
    layout="wide",
)


def money(value: float, currency: str) -> str:
    return f"{currency} {value:,.2f}"


def money_up(value: float, currency: str) -> str:
    return f"{currency} {math.ceil(value):,.0f}"


VTT_DATA_PATH = Path(__file__).resolve().parent.parent / "VTT Tool" / "VTT DATA.xlsx"


VTT_ROW_LABELS = [
    "1. Day Customer Order",
    "2. Day ILN/Supplier Order",
    "3. First Receipt Days",
    "4. Pack. prep. & load",
    "5. Transport to POL",
    "6. First Day to POL",
    "7. Cut off",
    "8. ETD",
    "9. Transit Duration (ETD>ETA)",
    "10. Days of flexibility",
    "11. Customs clearence",
    "12. Transport to plant",
    "13. Rounding",
    "14. Due Date",
]


def _vtt_int(row: pd.Series, *names: str) -> int:
    """First matching column as int; a name ending in '*' matches by prefix."""
    for name in names:
        if name.endswith("*"):
            match = next((c for c in row.index if str(c).startswith(name[:-1])), None)
        else:
            match = name if name in row.index else None
        if match is not None:
            value = pd.to_numeric(row[match], errors="coerce")
            return 0 if pd.isna(value) else int(round(float(value)))
    return 0


def vtt_gantt_steps(row: pd.Series, vtt_result: dict[str, int]) -> list[dict]:
    """Rows of the VTT timeline after applying the proposed sailing schedule."""
    green, blue = "#90ee90", "#87ceeb"

    def paint(final: int, plus: int, single: bool = False) -> list[dict]:
        if not final:
            return []
        length = 1 if single or plus <= 0 else plus
        return [{"start": max(1, final - length + 1), "end": final, "fill": green}]

    steps: list[dict] = []
    for final in (_vtt_int(row, "1 Day Customer Order"), _vtt_int(row, "2 Day ILN Order")):
        steps.append({"day": final, "plus": 0, "final": final, "segments": paint(final, 0, True)})

    plus3, final3 = _vtt_int(row, "3 .1 Time of Recept in ILN"), _vtt_int(row, "3.2 First Receipt Days")
    steps.append({"day": _vtt_int(row, "3 First Receipt Days"), "plus": plus3, "final": final3, "segments": paint(final3, plus3)})
    plus4, final4 = _vtt_int(row, "4.2 Packaging*"), _vtt_int(row, "4.3 Packaging*")
    steps.append({"day": _vtt_int(row, "4.1 Packaging*"), "plus": plus4, "final": final4, "segments": paint(final4, plus4)})
    plus5, final5 = _vtt_int(row, "5.2 Transport ILN to POL"), _vtt_int(row, "5.3 Transport ILN to POL")
    steps.append({"day": _vtt_int(row, "5.1 Transport ILN to POL"), "plus": plus5, "final": final5, "segments": paint(final5, plus5)})
    for name in ("6 First Day to POL", "7 Cutt off", "8 ETD"):
        final = (
            vtt_result["new_cutoff_day"]
            if name.startswith("7")
            else vtt_result["new_etd_day"]
            if name.startswith("8")
            else _vtt_int(row, name)
        )
        plus = _vtt_int(row, "First Day to POL") if name.startswith("6") else 0
        steps.append({"day": final, "plus": plus, "final": final, "segments": paint(final, plus, True)})

    eta = vtt_result["new_eta_day"]
    sea_days = vtt_result["new_transit_days"]
    steps.append({"day": vtt_result["new_etd_day"], "plus": sea_days, "final": eta, "segments": paint(eta, sea_days)})

    flex1, flex2 = _vtt_int(row, "Time for security"), _vtt_int(row, "Time for security2 buffer")
    flex_final = _vtt_int(row, "11 Days flexibility 2") - vtt_result["total_shift_days"]
    flex_plus = flex1 + flex2
    flex_segments: list[dict] = []
    if flex_final and flex_plus > 0:
        start = max(1, flex_final - flex_plus + 1)
        if flex1 > 0:
            flex_segments.append({"start": start, "end": min(flex_final, start + flex1 - 1), "fill": blue})
        if start + flex1 <= flex_final:
            flex_segments.append({"start": start + flex1, "end": flex_final, "fill": green})
    steps.append({"day": eta + 1, "plus": flex_plus, "final": flex_final, "segments": flex_segments})

    customs_plus = _vtt_int(row, "Cust.")
    customs_final = _vtt_int(row, "12 Customs Clearance") - vtt_result["total_shift_days"]
    steps.append({
        "day": customs_final if customs_plus <= 0 else flex_final + 1,
        "plus": customs_plus, "final": customs_final, "segments": paint(customs_final, customs_plus) if customs_plus > 0 else [],
    })

    transport_plus = _vtt_int(row, "Trpt POD/PFI vers Usine")
    transport_final = vtt_result["new_pre_rounding"]
    steps.append({
        "day": customs_final + 1, "plus": transport_plus, "final": transport_final,
        "segments": paint(transport_final, transport_plus) if transport_plus > 0 else [],
    })

    round_plus = vtt_result["new_rounding"]
    round_final = vtt_result["new_pre_rounding"] + round_plus
    steps.append({
        "day": round_final if round_plus <= 0 else transport_final + 1,
        "plus": round_plus, "final": round_final, "segments": paint(round_final, round_plus) if round_plus > 0 else [],
    })

    base_final = vtt_result["base_pre_rounding"] + vtt_result["base_rounding"]
    due = _vtt_int(row, "15 Due Date")
    due_plus = due - base_final if due >= base_final > 0 else 7
    due_final = round_final + due_plus
    steps.append({"day": round_final + 1, "plus": due_plus, "final": due_final, "segments": paint(due_final, due_plus)})
    return steps


def clt_card(label: str, value: str, background: str, border: str, text: str,
             delta: str = "", delta_color: str = "#475569") -> str:
    """Highlighted KPI box for the CLT values."""
    delta_html = (
        f"<div style='font-size:0.85rem;font-weight:600;color:{delta_color};'>{delta}</div>" if delta else ""
    )
    return (
        f"<div style='background:{background};border:2px solid {border};border-radius:10px;"
        f"padding:8px 14px;'>"
        f"<div style='font-size:0.85rem;color:{text};'>{label}</div>"
        f"<div style='font-size:1.9rem;font-weight:700;color:{text};line-height:1.2;'>{value}</div>"
        f"{delta_html}</div>"
    )


def vtt_gantt_html(steps: list[dict]) -> str:
    """Timeline table in the VTT2.py layout: Steps | Day | Day+ | Final Day + week/day Gantt grid."""
    start_date = datetime.today() - timedelta(days=datetime.today().weekday())
    num_days = max(110, max(s["final"] for s in steps) + 7)
    days = [start_date + timedelta(days=i) for i in range(num_days)]

    week_cells, current, span = [], None, 0
    for day in days:
        week = day.isocalendar()[1]
        if week == current:
            span += 1
            continue
        if current is not None:
            week_cells.append((current, span))
        current, span = week, 1
    week_cells.append((current, span))

    th_fixed = "padding:5px 7px;border:1px solid #eee;text-align:center;background:#f5f5f5;"
    html = ["<div style='overflow-x:auto;'><table style='width:100%;border-collapse:collapse;margin-top:8px;'><thead><tr>"]
    html.append("<th style='border:none;background:none;'></th>" * 4)
    for week, span in week_cells:
        html.append(
            f"<th colspan='{span}' style='padding:0 1px;border:1px solid #eee;min-width:28px;text-align:center;"
            f"background:#fffbe6;font-size:13.5px;font-weight:bold;'>W{week}</th>"
        )
    html.append("</tr><tr>")
    html.append(f"<th style='{th_fixed}min-width:200px;white-space:nowrap;'>Steps</th>")
    for header in ("Day", "Day+", "Final Day"):
        html.append(f"<th style='{th_fixed}min-width:50px;width:50px;'>{header}</th>")
    for day in days:
        color = "#ffd6d6" if day.weekday() in (5, 6) else "#e3eafc"
        html.append(
            f"<th style='padding:0 1px;border:1px solid #eee;min-width:20px;width:20px;height:50px;text-align:center;"
            f"background:{color};font-size:12px;vertical-align:bottom;'>{day.strftime('%a')[0].upper()}</th>"
        )
    html.append("</tr></thead><tbody>")

    cell = "border:1px solid #eee;height:15px;line-height:15px;padding:1px 4px;text-align:center;"
    for label, step in zip(VTT_ROW_LABELS, steps):
        html.append("<tr style='height:15px;'>")
        html.append(
            f"<td style='{cell}text-align:left;font-weight:bold;background:#f5f5f5;min-width:200px;white-space:nowrap;'>{label}</td>"
        )
        for value in (step["day"], step["plus"], step["final"]):
            html.append(f"<td style='{cell}'>{value if value or value == 0 else '-'}</td>")
        for index, day in enumerate(days, start=1):
            fill = next((s["fill"] for s in step["segments"] if s["start"] <= index <= s["end"]), None)
            if fill is None and day.weekday() in (5, 6):
                fill = "#ffd6d6"
            style = cell + (f"background-color:{fill};" if fill else "")
            html.append(f"<td style='{style}'></td>")
        html.append("</tr>")
    html.append("</tbody></table></div>")
    return "".join(html)


WACC_BY_COUNTRY = {"ES": 8.9, "RO": 12.4, "TR": 15.2, "BR": 12.3, "PT": 10.9}
SCENARIO_COLORS = {"Scenario A": "#64748B", "Scenario B": "#0F766E"}


def _id_sort_key(value: str) -> tuple:
    """IDs starting with letters first, then numeric IDs; digits compared as numbers."""
    text = str(value).strip().upper()
    parts = [int(t) if t.isdigit() else t for t in re.split(r"(\d+)", text) if t]
    return (0 if text[:1].isalpha() else 1, [(isinstance(t, int), t if isinstance(t, int) else 0, str(t)) for t in parts])


@st.cache_data(show_spinner=False)
def load_vtt_data(path: str, mtime: float) -> pd.DataFrame:
    """Read the VTT route table (same file used by VTT2.py, read-only)."""
    df = pd.read_excel(path, sheet_name=0)
    df.columns = [str(c).strip() for c in df.columns]
    needed = ["ID", "POL", "POD", "Transit time", "13 Transport to Plant", "14 Rounding", "1 Day Customer Order"]
    df = df.dropna(subset=["13 Transport to Plant", "14 Rounding", "1 Day Customer Order"])
    df = df.loc[:, ~df.columns.duplicated()]
    for col in needed:
        if col not in df.columns:
            raise KeyError(f"Column '{col}' not found in VTT DATA.xlsx")
    df["ID"] = df["ID"].astype(str).str.strip()
    df = df.sort_values("ID", key=lambda col: col.map(_id_sort_key), kind="stable")
    return df.reset_index(drop=True)


def round_to_sunday(day_before_rounding: float, sunday_offset: int) -> int:
    """First Sunday on or after day_before_rounding (rounding is 0 to 6 days).

    If the previous step already ends on a Sunday, no rounding is needed.
    In the VTT timeline Sundays fall on day numbers that are multiples of 7
    (plus a route-specific offset when the source data is not aligned).
    """
    return sunday_offset + 7 * math.ceil((day_before_rounding - sunday_offset) / 7)


def vtt_clt_after_transit_change(
    row: pd.Series,
    reduction_days: int,
    cutoff_weekday: int | None = None,
    etd_weekday: int | None = None,
) -> dict[str, int]:
    """Re-run downstream VTT steps after changing transit and sailing weekdays.

    Weekday indexes use Monday=0 through Sunday=6, with VTT day 1 anchored
    to Monday. Cut-off may occur on the first day at POL; ETD must be later.
    """
    # Same as VTT2.py: CLT = final day of "14 Rounding" - day of "1 Day Customer Order" + 1.
    start_day = int(row["1 Day Customer Order"])
    base_final_day = int(row["14 Rounding"])
    base_clt = base_final_day - start_day + 1
    base_pre_rounding = int(row["13 Transport to Plant"])
    base_rounding = base_final_day - base_pre_rounding
    first_day_to_pol = int(row["6 First Day to POL"])
    base_cutoff_day = int(row["7 Cutt off"])
    base_etd_day = int(row["8 ETD"])
    cutoff_weekday = (
        (base_cutoff_day - 1) % 7 if cutoff_weekday is None else cutoff_weekday
    )
    etd_weekday = (base_etd_day - 1) % 7 if etd_weekday is None else etd_weekday

    new_cutoff_day = first_day_to_pol + (
        cutoff_weekday - (first_day_to_pol - 1) % 7
    ) % 7
    etd_wait = (etd_weekday - (new_cutoff_day - 1) % 7) % 7
    new_etd_day = new_cutoff_day + (etd_wait or 7)

    total_shift_days = base_etd_day - new_etd_day + reduction_days
    offset = base_final_day % 7
    new_pre_rounding = base_pre_rounding - total_shift_days
    if total_shift_days == 0:
        new_final_day = base_final_day
    else:
        new_final_day = round_to_sunday(new_pre_rounding, offset)
    new_clt = new_final_day - start_day + 1
    base_eta_day = int(row["9 ETD>ETA"])
    current_transit_days = int(row["Transit time"])
    return {
        "base_clt": base_clt,
        "base_pre_rounding": base_pre_rounding,
        "base_rounding": base_rounding,
        "base_cutoff_day": base_cutoff_day,
        "base_etd_day": base_etd_day,
        "new_cutoff_day": new_cutoff_day,
        "new_etd_day": new_etd_day,
        "new_eta_day": base_eta_day - total_shift_days,
        "new_transit_days": current_transit_days - reduction_days,
        "total_shift_days": total_shift_days,
        "new_pre_rounding": new_pre_rounding,
        "new_rounding": new_final_day - new_pre_rounding,
        "new_clt": new_clt,
    }


def scenario_inputs(
    label: str, *, defaults: dict[str, float], clt_from_vtt: bool = False
) -> ScenarioInputs:
    st.subheader(label)
    cost_col1, cost_col2, cost_col3 = st.columns(3)
    with cost_col1:
        ocean_freight = st.number_input(
            "Ocean freight per container",
            min_value=0.0,
            value=defaults["ocean_freight"],
            step=50.0,
            key=f"{label}_ocean_freight",
            help="Ocean freight cost for one container.",
        )
    with cost_col2:
        inland_cost = st.number_input(
            "Port-to-plant transportation per container",
            min_value=0.0,
            value=defaults["inland_cost"],
            step=25.0,
            key=f"{label}_inland_cost",
        )
    with cost_col3:
        other_cost = st.number_input(
            "Additional logistics cost per container",
            min_value=0.0,
            value=defaults["other_cost"],
            step=25.0,
            key=f"{label}_other_cost",
            help="For example: handling, documentation or other route-specific charges.",
        )

    clt_key = f"{label}_clt_days"
    # The value is set through session state only (passing `value` too triggers a warning).
    st.session_state[clt_key] = (
        float(defaults["clt_days"])
        if clt_from_vtt
        else st.session_state.get(clt_key, float(defaults["clt_days"]))
    )
    clt_days = st.number_input(
        "Complete Customer Lead Time (days)",
        min_value=0.0,
        step=1.0,
        key=clt_key,
        disabled=clt_from_vtt,
        help=(
            "End-to-end lead time, including supplier processing, production, "
            "export handling, transportation, customs and delivery. "
            "Locked to the VTT value while 'Use VTT route' is checked."
        ),
    )
    return ScenarioInputs(
        ocean_freight_per_container=ocean_freight,
        inland_transport_per_container=inland_cost,
        additional_logistics_per_container=other_cost,
        customer_lead_time_days=clt_days,
    )


st.title("🚢 Maritime CLT & Inventory Cost Optimizer")
st.caption(
    "Model the operational and financial impact of lead-time changes, from route "
    "assumptions to an annual business case."
)

st.subheader("1. Global assumptions")
with st.container(border=True):
    global_col1, global_col2, global_col3 = st.columns(3)
    with global_col1:
        wacc_choice = st.selectbox(
            "WACC (% per year)",
            options=list(WACC_BY_COUNTRY),
            index=None,
            accept_new_options=True,
            placeholder="Select a country or type a value (e.g. 8.5)",
            format_func=lambda k: (
                f"{k} ? {WACC_BY_COUNTRY[k]:.1f}%" if k in WACC_BY_COUNTRY else k
            ),
            help="Pick a country WACC or type any percentage.",
        )
        if wacc_choice in WACC_BY_COUNTRY:
            wacc_percent = WACC_BY_COUNTRY[wacc_choice]
        elif wacc_choice is None:
            wacc_percent = 8.0
        else:
            try:
                wacc_percent = float(str(wacc_choice).replace("%", "").replace(",", ".").strip())
                if wacc_percent < 0:
                    raise ValueError
            except ValueError:
                st.error("Enter a valid non-negative percentage, e.g. 8.5")
                st.stop()
        use_carrying_rate = st.checkbox(
            "Use separate inventory carrying cost rate",
            value=False,
            help="This replaces WACC for inventory financing; it is not added to WACC.",
        )
        carrying_percent = st.number_input(
            "Inventory carrying cost rate (% per year)",
            min_value=0.0,
            value=10.0,
            step=0.25,
            disabled=not use_carrying_rate,
        )
    with global_col2:
        annual_days = st.number_input(
            "Days per year for per-container finance calculation",
            min_value=1,
            value=DEFAULT_DAYS_PER_YEAR,
            step=1,
            help="Defaults to 365. Annual pipeline cost uses pipeline value × annual rate.",
        )
        annual_volume = st.number_input(
            "Annual container volume",
            min_value=0,
            value=None,
            step=1,
            format="%d",
        )
        annual_volume = annual_volume or 0
        plant_weeks = st.number_input(
            "Plant working weeks per year",
            min_value=1,
            value=48,
            step=1,
            format="%d",
            help="Used to convert annual volume into containers per week.",
        )
        containers_per_day = annual_volume / annual_days
        st.caption(
            f"Pipeline inventory = annual volume ÷ {annual_days} × CLT days × "
            f"value per container ({containers_per_day:,.3f} containers/calendar day)."
        )
    with global_col3:
        cargo_value = st.number_input(
            "Cargo value per container",
            min_value=0.0,
            value=100_000.0,
            step=5_000.0,
        )
        currency = st.selectbox("Currency", ["EUR", "USD", "GBP"], index=0)

financing_rate = (
    carrying_percent if use_carrying_rate else wacc_percent
) / 100

vtt_result = None
st.subheader("2. Route and lead-time change")
if not VTT_DATA_PATH.exists():
    st.info(f"VTT data not found at {VTT_DATA_PATH}. Enter the CLT manually below.")
else:
    try:
        vtt_df = load_vtt_data(str(VTT_DATA_PATH), VTT_DATA_PATH.stat().st_mtime)
    except Exception as error:  # noqa: BLE001
        vtt_df = None
        st.warning(f"Could not read VTT data: {error}")
    if vtt_df is not None and not vtt_df.empty:
        use_vtt = st.checkbox(
            "Use VTT route to set Scenario A / B CLT", value=True
        )
        labels = [
            f"{r['ID']} | {r['POL']}>{r['POD']}"
            + (f" | {r['Carrier']}" if "Carrier" in vtt_df.columns and pd.notna(r["Carrier"]) else "")
            + (f" | {r['Shipper']}" if "Shipper" in vtt_df.columns and pd.notna(r["Shipper"]) else "")
            for _, r in vtt_df.iterrows()
        ]
        vtt_col1, vtt_col2 = st.columns([2, 1])
        with vtt_col1:
            chosen = st.selectbox("VTT route", labels)
        vtt_row = vtt_df.iloc[labels.index(chosen)]
        sea_days = pd.to_numeric(vtt_row.get("Transit time"), errors="coerce")
        current_pol_pod = int(0 if pd.isna(sea_days) else sea_days)
        with vtt_col2:
            new_pol_pod = st.number_input(
                "New POL>POD (days)",
                min_value=0,
                value=current_pol_pod,
                step=1,
                format="%d",
                key=f"new_pol_pod_{chosen}",
                help="Proposed POL>POD sea transit (without security time). Reduction = current POL>POD - New POL>POD.",
            )
        st.markdown("**6. First Day to POL → Cut-off → ETD**")
        weekdays = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
        base_cutoff_day = int(vtt_row["7 Cutt off"])
        base_etd_day = int(vtt_row["8 ETD"])
        schedule_col1, schedule_col2 = st.columns(2)
        with schedule_col1:
            cutoff_weekday_name = st.selectbox(
                "Cut-off weekday",
                weekdays,
                index=(base_cutoff_day - 1) % 7,
                key=f"cutoff_weekday_{chosen}",
                help="Select the weekly cut-off on or after the shipment first reaches POL.",
            )
        with schedule_col2:
            etd_weekday_name = st.selectbox(
                "ETD weekday",
                weekdays,
                index=(base_etd_day - 1) % 7,
                key=f"etd_weekday_{chosen}",
                help="ETD is the next selected weekday strictly after cut-off; if both weekdays match, ETD is the following week.",
            )
        cutoff_weekday = weekdays.index(cutoff_weekday_name)
        etd_weekday = weekdays.index(etd_weekday_name)
        reduction_days = current_pol_pod - int(new_pol_pod)
        vtt_result = vtt_clt_after_transit_change(
            vtt_row,
            reduction_days,
            cutoff_weekday,
            etd_weekday,
        )
        v1, v2, v3, v4, v5 = st.columns(5)
        v1.metric("Current POL>POD (VTT)", f"{current_pol_pod} d",
                  help="Sea transit only (Transit time in VTT), without security time.")
        clt_delta = vtt_result["new_clt"] - vtt_result["base_clt"]
        delta_color = "#15803D" if clt_delta < 0 else "#B91C1C" if clt_delta > 0 else "#475569"
        v2.markdown(
            clt_card("Current CLT (VTT)", f"{vtt_result['base_clt']} d", "#EEF2FF", "#4F46E5", "#312E81"),
            unsafe_allow_html=True,
        )
        v3.metric("Current rounding", f"{vtt_result['base_rounding']} d")
        v4.metric("New rounding to Sunday", f"{vtt_result['new_rounding']} d",
                  delta=f"{vtt_result['new_rounding'] - vtt_result['base_rounding']:+d} d vs current",
                  delta_color="off")
        v5.markdown(
            clt_card("New CLT", f"{vtt_result['new_clt']} d", "#ECFDF5", "#0F766E", "#134E4A",
                     f"{clt_delta:+d} d vs current", delta_color),
            unsafe_allow_html=True,
        )
        effective = vtt_result["base_clt"] - vtt_result["new_clt"]
        st.caption(
            f"Cut-off: day {vtt_result['new_cutoff_day']} ({cutoff_weekday_name}); "
            f"ETD: day {vtt_result['new_etd_day']} ({etd_weekday_name}). "
            f"Changing POL>POD from {current_pol_pod} to {int(new_pol_pod)} days "
            f"and these weekdays changes CLT by {-effective:+d} days "
            f"({effective / 7:.2f} weeks), including the VTT rounding step."
        )
        with st.expander("Effective CLT reduction vs. sea-transit reduction"):
            table = []
            for d in range(0, 15):
                r = vtt_clt_after_transit_change(
                    vtt_row,
                    d,
                    cutoff_weekday,
                    etd_weekday,
                )
                table.append(
                    {
                        "New POL>POD (days)": current_pol_pod - d,
                        "POL>POD reduction (days)": d,
                        "Rounding (days)": r["new_rounding"],
                        "New CLT (days)": r["new_clt"],
                        "CLT reduction (days)": r["base_clt"] - r["new_clt"],
                    }
                )
            st.dataframe(pd.DataFrame(table), hide_index=True, width="stretch")
        if not use_vtt:
            vtt_result = None

vtt_export = None
if vtt_result:
    vtt_export = {
        "route": chosen,
        "current_pol_pod": current_pol_pod,
        "new_pol_pod": int(new_pol_pod),
        "start_day": int(vtt_row["1 Day Customer Order"]),
        "base_final_day": int(vtt_row["14 Rounding"]),
        "base_pre_rounding": int(vtt_row["13 Transport to Plant"]),
        "base_cutoff_day": int(vtt_row["7 Cutt off"]),
        "base_etd_day": int(vtt_row["8 ETD"]),
        "first_day_to_pol": int(vtt_row["6 First Day to POL"]),
        "cutoff_weekday": cutoff_weekday,
        "etd_weekday": etd_weekday,
        "cutoff_name": cutoff_weekday_name,
        "etd_name": etd_weekday_name,
    }
vtt_base_clt = float(vtt_result["base_clt"]) if vtt_result else 0.0
vtt_new_clt = float(vtt_result["new_clt"]) if vtt_result else 0.0
st.subheader("3. Current and proposed Transport Cost")
scenario_col_a, scenario_col_b = st.columns(2)
with scenario_col_a:
    with st.container(border=True):
        scenario_a = scenario_inputs(
            "A — Current logistics solution",
            defaults={
                "ocean_freight": 0.0,
                "inland_cost": 0.0,
                "other_cost": 0.0,
                "clt_days": vtt_base_clt,
            },
            clt_from_vtt=vtt_result is not None,
        )
with scenario_col_b:
    with st.container(border=True):
        scenario_b = scenario_inputs(
            "B — Proposed logistics solution",
            defaults={
                "ocean_freight": 0.0,
                "inland_cost": 0.0,
                "other_cost": 0.0,
                "clt_days": vtt_new_clt,
            },
            clt_from_vtt=vtt_result is not None,
        )

try:
    metrics_a = calculate_scenario(
        scenario_a,
        cargo_value_per_container=cargo_value,
        containers_consumed_per_day=containers_per_day,
        annual_container_volume=annual_volume,
        annual_financing_rate=financing_rate,
        days_per_year=annual_days,
    )
    metrics_b = calculate_scenario(
        scenario_b,
        cargo_value_per_container=cargo_value,
        containers_consumed_per_day=containers_per_day,
        annual_container_volume=annual_volume,
        annual_financing_rate=financing_rate,
        days_per_year=annual_days,
    )
except ValueError as error:
    st.error(str(error))
    st.stop()

comparison = calculate_comparison(
    metrics_a, metrics_b, annual_container_volume=annual_volume
)

st.divider()
st.header("4. Decision summary")

clt_reduction_weeks = comparison["transit_time_reduction_days"] / 7
containers_per_week = annual_volume / plant_weeks
net_benefit = comparison["net_annual_economic_benefit"]
preferred_scenario = (
    "B"
    if net_benefit > 0
    else "A"
    if net_benefit < 0
    else "Tie"
)
with st.container(border=True):
    summary_col1, summary_col2 = st.columns([1, 2])
    summary_col1.metric("Preferred scenario", preferred_scenario)
    summary_col2.metric(
        "Annual economic benefit (Scenario B vs. A)",
        money_up(net_benefit, currency),
        delta_color="off",
    )
st.caption(
    "Working capital released is a balance-sheet/cash-flow measure; annual financing "
    "savings and net annual economic benefit are separate measures."
)

st.subheader("Operational impact")
case_col1, case_col2 = st.columns(2)
case_col1.metric(
    "CLT reduction",
    f"{math.ceil(comparison['transit_time_reduction_days']):,} days",
    f"{-math.ceil(comparison['transit_time_reduction_percent']):,}%",
    delta_color="inverse",
)
_transport_delta = comparison["annual_transportation_savings"]
case_col2.metric(
    "Transportation savings (annual)"
    if _transport_delta >= 0
    else "Additional transport cost (annual)",
    money_up(abs(_transport_delta), currency),
)

st.subheader("Financial impact")
finance_col1, finance_col2, finance_col3 = st.columns(3)
finance_col1.metric(
    "Working capital released",
    money_up(cargo_value * containers_per_week * clt_reduction_weeks, currency),
    help=(
        "Annual volume ÷ plant weeks × cargo value × weeks of CLT reduction "
        f"(CLT reduction ÷ 7 = {clt_reduction_weeks:,.2f} weeks; "
        f"{containers_per_week:,.2f} containers/week). This is not an annual P&L saving."
    ),
)
finance_col2.metric(
    "Annual financing savings",
    money_up(comparison["annual_financing_savings"], currency),
    help="Annual reduction in inventory financing cost; not the capital released.",
)
finance_col3.metric(
    "Financing-only break-even freight premium",
    money_up(comparison["maximum_freight_premium_per_container"], currency)
    + " / container",
    help=(
        "Excludes service level, safety stock, obsolescence, warehousing, stockouts "
        "and other benefits not explicitly modeled."
    ),
)

with st.expander("Additional indicator: working capital released per week"):
    st.metric(
        "Per 1 week of CLT reduction",
        money_up(cargo_value * containers_per_week, currency),
        help="Annual volume ÷ plant working weeks × cargo value per container.",
    )

actual_premium = comparison["actual_additional_transport_cost_per_container"]
break_even_premium = comparison["maximum_freight_premium_per_container"]
st.caption("Transport-premium check against the financing-only break-even threshold:")
if actual_premium < break_even_premium:
    st.success(
        f"Scenario B's additional transport cost ({money_up(actual_premium, currency)} "
        f"per container) is below its financing-only break-even premium."
    )
elif actual_premium == break_even_premium:
    st.info("Scenario B is at financing-only break-even on the transport premium.")
else:
    st.warning(
        f"Scenario B's additional transport cost ({money_up(actual_premium, currency)} "
        f"per container) exceeds its financing-only break-even premium."
    )

scenario_comparison = pd.DataFrame(
    [
        (
            "CLT (days)",
            math.ceil(metrics_a.analysis_days),
            math.ceil(metrics_b.analysis_days),
        ),
        (
            "Pipeline inventory value",
            metrics_a.pipeline_inventory_value,
            metrics_b.pipeline_inventory_value,
        ),
        (
            "Finance cost / container",
            metrics_a.financing_cost_per_container,
            metrics_b.financing_cost_per_container,
        ),
        (
            "Annual financing cost",
            metrics_a.annual_financing_cost,
            metrics_b.annual_financing_cost,
        ),
        (
            "Transport cost / container",
            metrics_a.transportation_cost_per_container,
            metrics_b.transportation_cost_per_container,
        ),
        (
            "Annual transportation cost",
            metrics_a.annual_transportation_cost,
            metrics_b.annual_transportation_cost,
        ),
        (
            "Economic cost / container",
            metrics_a.total_economic_cost_per_container,
            metrics_b.total_economic_cost_per_container,
        ),
        (
            "Annual economic logistics cost",
            metrics_a.annual_total_economic_cost,
            metrics_b.annual_total_economic_cost,
        ),
    ],
    columns=["Metric", "Scenario A", "Scenario B"],
).set_index("Metric")
scenario_comparison["Change (B − A)"] = (
    scenario_comparison["Scenario B"] - scenario_comparison["Scenario A"]
)
monetary_metrics = scenario_comparison.index.drop("CLT (days)")
scenario_comparison_style = (
    scenario_comparison.style.format(
        lambda value: money_up(value, currency),
        subset=pd.IndexSlice[
            monetary_metrics, ["Scenario A", "Scenario B"]
        ],
    )
    .format(
        "{:,.0f} days",
        subset=pd.IndexSlice[["CLT (days)"], ["Scenario A", "Scenario B"]],
    )
    .format(
        lambda value: f"{currency} {math.ceil(value):+,.0f}",
        subset=pd.IndexSlice[monetary_metrics, ["Change (B − A)"]],
    )
    .format(
        "{:+,.0f} days",
        subset=pd.IndexSlice[["CLT (days)"], ["Change (B − A)"]],
    )
    .set_properties(
        subset=["Change (B − A)"],
        **{"background-color": "#F1F5F9", "font-weight": "bold"},
    )
)
with st.expander("Detailed scenario comparison"):
    st.dataframe(scenario_comparison_style, width="stretch")

with st.expander("Calculation breakdown & assumptions"):
    st.markdown(
        f"""
        - **Time basis:** Complete Customer Lead Time (CLT) of each scenario.
        - **Pipeline inventory value:** cargo value × containers consumed per day × CLT days.
        - **Working capital released:** annual volume ÷ plant working weeks × cargo value × (CLT reduction ÷ 7).
        - **Per-container financing cost:** cargo value × selected annual rate × CLT days ÷ {annual_days}.
        - **Annual financing cost:** pipeline inventory value × selected annual rate.
        - **Annual financing savings:** (annual container volume ÷ days per year) × cargo value per container × CLT reduction (days) × selected annual rate.
        - **Transportation cost per container:** ocean freight + inland transportation + additional logistics.
        - **Net annual economic benefit:** Scenario A annual economic logistics cost − Scenario B annual economic logistics cost.
        - **Finance rate used:** {financing_rate:.2%} ({'inventory carrying cost' if use_carrying_rate else 'WACC'}).

        Inventory value and released working capital are balance-sheet/cash-flow
        measures, not P&L savings. Safety stock, service-level benefits, obsolescence,
        warehousing, demurrage, detention and stockout costs are not modeled.
        """
    )

st.divider()
st.header("5. VTT & Charts")

st.subheader("VTT with New POL>POD and schedule")
if vtt_result is None:
    st.info("Select a VTT route and enable 'Use VTT route' to see the VTT timeline.")
else:
    st.caption(
        f"{chosen} · POL>POD: {current_pol_pod} d → {int(new_pol_pod)} d · "
        f"Cut-off: {weekdays[(vtt_result['new_cutoff_day'] - 1) % 7]} · "
        f"ETD: {weekdays[(vtt_result['new_etd_day'] - 1) % 7]} · "
        f"CLT: {vtt_result['base_clt']} d → {vtt_result['new_clt']} d · "
        f"Rounding: {vtt_result['base_rounding']} d → {vtt_result['new_rounding']} d."
    )
    st.markdown(
        vtt_gantt_html(vtt_gantt_steps(vtt_row, vtt_result)),
        unsafe_allow_html=True,
    )

st.subheader("Charts")

annual_cost_data = pd.DataFrame(
    [
        {"Scenario": scenario, "Cost type": cost_type, "Annual cost": amount}
        for scenario, metrics in [("Scenario A", metrics_a), ("Scenario B", metrics_b)]
        for cost_type, amount in [
            ("Transportation", metrics.annual_transportation_cost),
            ("Inventory financing", metrics.annual_financing_cost),
        ]
    ]
)
chart_col1, chart_col2 = st.columns(2)
with chart_col1:
    inventory_start = metrics_a.pipeline_inventory_value
    inventory_end = metrics_b.pipeline_inventory_value
    inventory_change = inventory_end - inventory_start
    currency_symbol = {"EUR": "€", "USD": "$", "GBP": "£"}[currency]
    pipeline_chart = go.Figure(
        go.Waterfall(
            x=["Scenario A", "Change to Scenario B", "Scenario B"],
            y=[inventory_start, inventory_change, inventory_end],
            measure=["absolute", "relative", "total"],
            text=[
                f"{currency_symbol}{inventory_start / 1_000:,.1f}K",
                f"{currency_symbol}{inventory_change / 1_000:+,.1f}K",
                f"{currency_symbol}{inventory_end / 1_000:,.1f}K",
            ],
            customdata=[
                inventory_start / 1_000,
                inventory_change / 1_000,
                inventory_end / 1_000,
            ],
            textposition="outside",
            connector={"line": {"color": "#CBD5E1"}},
            increasing={"marker": {"color": "#0F766E"}},
            decreasing={"marker": {"color": "#DC2626"}},
            totals={"marker": {"color": "#1D4ED8"}},
            hovertemplate=f"%{{customdata:,.1f}}K {currency_symbol}<extra></extra>",
        )
    )
    pipeline_chart.update_layout(
        title="Pipeline inventory value (capital tied up)",
        yaxis_title=f"Value ({currency_symbol})",
        yaxis={"range": [0, max(inventory_start, inventory_end, 1) * 1.2]},
        xaxis={"showgrid": False},
        template="plotly_white",
        font={"color": "#334155"},
        margin={"t": 80, "l": 20, "r": 20, "b": 20},
    )
    st.plotly_chart(pipeline_chart, width="stretch")
with chart_col2:
    annual_cost_chart = px.bar(
        annual_cost_data,
        x="Scenario",
        y="Annual cost",
        color="Cost type",
        barmode="group",
        title="Annual transportation and financing costs",
        labels={"Annual cost": f"Annual cost ({currency})"},
        color_discrete_map={
            "Transportation": "#1D4ED8",
            "Inventory financing": "#F59E0B",
        },
    )
    annual_cost_chart.update_traces(
        texttemplate="%{y:,.0f}", textposition="outside", cliponaxis=False
    )
    annual_cost_chart.update_layout(
        template="plotly_white",
        font={"color": "#334155"},
        xaxis={"showgrid": False},
        yaxis={"gridcolor": "#E2E8F0"},
        margin={"t": 70, "l": 20, "r": 20, "b": 20},
    )
    st.plotly_chart(annual_cost_chart, width="stretch")

annual_cost_start = metrics_a.annual_total_economic_cost
annual_cost_end = metrics_b.annual_total_economic_cost
annual_cost_change = annual_cost_end - annual_cost_start
total_cost_chart = go.Figure(
    go.Waterfall(
        x=["Scenario A", "Change to Scenario B", "Scenario B"],
        y=[annual_cost_start, annual_cost_change, annual_cost_end],
        measure=["absolute", "relative", "total"],
        text=[
            f"{currency_symbol}{annual_cost_start / 1_000:,.1f}K",
            f"{currency_symbol}{annual_cost_change / 1_000:+,.1f}K",
            f"{currency_symbol}{annual_cost_end / 1_000:,.1f}K",
        ],
        customdata=[
            annual_cost_start / 1_000,
            annual_cost_change / 1_000,
            annual_cost_end / 1_000,
        ],
        textposition="outside",
        connector={"line": {"color": "#CBD5E1"}},
        increasing={"marker": {"color": "#0F766E"}},
        decreasing={"marker": {"color": "#DC2626"}},
        totals={"marker": {"color": "#1D4ED8"}},
        hovertemplate=f"%{{customdata:,.1f}}K {currency_symbol}<extra></extra>",
    )
)
total_cost_chart.update_layout(
    title="Annual total economic logistics cost",
    yaxis_title=f"Annual cost ({currency_symbol})",
    yaxis={
        "range": [
            0,
            max(annual_cost_start, annual_cost_end, 1) * 1.2,
        ],
        "gridcolor": "#E2E8F0",
    },
    xaxis={"showgrid": False},
    template="plotly_white",
    font={"color": "#334155"},
    margin={"t": 70, "l": 20, "r": 20, "b": 20},
)
st.plotly_chart(total_cost_chart, width="stretch")

st.divider()
st.header("6. Export")
st.download_button(
    "📥 Download Excel",
    data=build_workbook(
        wacc_percent=wacc_percent,
        use_carrying_rate=use_carrying_rate,
        carrying_percent=carrying_percent,
        annual_days=annual_days,
        annual_volume=annual_volume,
        plant_weeks=plant_weeks,
        cargo_value=cargo_value,
        currency=currency,
        scenario_a=scenario_a,
        scenario_b=scenario_b,
        vtt=vtt_export,
        gantt_steps=vtt_gantt_steps(vtt_row, vtt_result) if vtt_result else None,
        gantt_labels=VTT_ROW_LABELS,
    ),
    file_name=f"CLT_analysis_{datetime.now():%Y%m%d_%H%M}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
