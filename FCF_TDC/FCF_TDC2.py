"""Streamlit dashboard for maritime lead-time and inventory cost analysis."""

from __future__ import annotations

import math
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from calculations import (
    DEFAULT_DAYS_PER_YEAR,
    ScenarioInputs,
    calculate_comparison,
    calculate_scenario,
    calculate_sensitivity,
)


st.set_page_config(
    page_title="Maritime CLT & Inventory Cost Optimizer",
    page_icon="🚢",
    layout="wide",
)


def money(value: float, currency: str) -> str:
    return f"{currency} {value:,.2f}"


def money_up(value: float, currency: str) -> str:
    return f"{currency} {math.ceil(value):,.0f}"


VTT_DATA_PATH = Path(__file__).resolve().parent.parent / "VTT Tool" / "VTT DATA.xlsx"


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
    return df.reset_index(drop=True)


def round_to_sunday(day_before_rounding: float, sunday_offset: int) -> int:
    """First day >= day_before_rounding that lands on a Sunday.

    In the VTT timeline Sundays fall on day numbers that are multiples of 7
    (plus a route-specific offset when the source data is not aligned).
    """
    return sunday_offset + 7 * math.ceil((day_before_rounding - sunday_offset) / 7)


def vtt_clt_after_transit_change(row: pd.Series, reduction_days: int) -> dict[str, int]:
    """Re-run VTT steps 13 -> 14 after changing only POL>POD transit.

    Every cumulative step after the sea leg moves by the same number of days,
    so stage 13 (Transport to Plant) shifts, and Rounding is recomputed.
    """
    # Same as VTT2.py: CLT = final day of "14 Rounding" - day of "1 Day Customer Order" + 1.
    start_day = int(row["1 Day Customer Order"])
    base_final_day = int(row["14 Rounding"])
    base_clt = base_final_day - start_day + 1
    base_pre_rounding = int(row["13 Transport to Plant"])
    base_rounding = base_final_day - base_pre_rounding
    offset = base_final_day % 7
    new_pre_rounding = base_pre_rounding - reduction_days
    new_final_day = round_to_sunday(new_pre_rounding, offset)
    new_clt = new_final_day - start_day + 1
    return {
        "base_clt": base_clt,
        "base_pre_rounding": base_pre_rounding,
        "base_rounding": base_rounding,
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

    clt_days = st.number_input(
        "Complete Customer Lead Time (days)",
        min_value=0.0,
        value=defaults["clt_days"],
        step=1.0,
        key=f"{label}_clt_days_{int(defaults['clt_days'])}" if clt_from_vtt else f"{label}_clt_days",
        help=(
            "End-to-end lead time, including supplier processing, production, "
            "export handling, transportation, customs and delivery."
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
    "Compare route economics without confusing inventory value, working capital "
    "and annual P&L savings."
)

with st.expander("Global assumptions", expanded=True):
    global_col1, global_col2, global_col3 = st.columns(3)
    with global_col1:
        wacc_percent = st.number_input(
            "WACC (% per year)", min_value=0.0, value=8.0, step=0.25
        )
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
st.subheader("VTT link ? CLT impact of a POL>POD transit change")
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
            for _, r in vtt_df.iterrows()
        ]
        vtt_col1, vtt_col2 = st.columns([2, 1])
        with vtt_col1:
            chosen = st.selectbox("VTT route", labels)
        with vtt_col2:
            reduction_days = st.number_input(
                "POL>POD transit reduction (days)",
                value=0,
                step=1,
                format="%d",
                help="Negative values simulate a longer sea transit.",
            )
        vtt_row = vtt_df.iloc[labels.index(chosen)]
        vtt_result = vtt_clt_after_transit_change(vtt_row, int(reduction_days))
        sea_days = vtt_row["Transit time"]
        security_days = vtt_row.get("Time for security", 0)
        security_days = 0 if pd.isna(security_days) else security_days
        v1, v2, v3, v4, v5 = st.columns(5)
        v1.metric("POL>POD transit (VTT)", f"{int(sea_days) if pd.notna(sea_days) else 0} d",
                  help="'Transit time' column in VTT; Time for security is shown separately.")
        v2.metric("Current CLT (VTT)", f"{vtt_result['base_clt']} d")
        v3.metric("New day before rounding", f"{vtt_result['new_pre_rounding']} d",
                  delta=f"-{int(reduction_days)} d" if reduction_days else None, delta_color="off")
        v4.metric("New rounding to Sunday", f"{vtt_result['new_rounding']} d",
                  delta=f"{vtt_result['new_rounding'] - vtt_result['base_rounding']:+d} d vs current",
                  delta_color="off")
        v5.metric("New CLT", f"{vtt_result['new_clt']} d",
                  delta=f"{vtt_result['new_clt'] - vtt_result['base_clt']:+d} d")
        effective = vtt_result["base_clt"] - vtt_result["new_clt"]
        st.caption(
            f"A {int(reduction_days)}-day sea-transit change moves the CLT by {-effective:+d} days "
            f"({effective / 7:.2f} weeks): the rounding stage re-completes the week up to Sunday, "
            "so part of the saving can be absorbed."
        )
        with st.expander("Effective CLT reduction vs. sea-transit reduction"):
            table = []
            for d in range(0, 15):
                r = vtt_clt_after_transit_change(vtt_row, d)
                table.append(
                    {
                        "Transit reduction (days)": d,
                        "Rounding (days)": r["new_rounding"],
                        "New CLT (days)": r["new_clt"],
                        "CLT reduction (days)": r["base_clt"] - r["new_clt"],
                    }
                )
            st.dataframe(pd.DataFrame(table), hide_index=True, width="stretch")
        if not use_vtt:
            vtt_result = None

vtt_base_clt = float(vtt_result["base_clt"]) if vtt_result else 0.0
vtt_new_clt = float(vtt_result["new_clt"]) if vtt_result else 0.0
scenario_col_a, scenario_col_b = st.columns(2)
with scenario_col_a:
    scenario_a = scenario_inputs(
        "Scenario A — Current logistics solution",
        defaults={
            "ocean_freight": 0.0,
            "inland_cost": 0.0,
            "other_cost": 0.0,
            "clt_days": vtt_base_clt,
        },
        clt_from_vtt=vtt_result is not None,
    )
with scenario_col_b:
    scenario_b = scenario_inputs(
        "Scenario B — Proposed logistics solution",
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
st.header("Financial comparison")

kpi1, kpi1_week, kpi2, kpi3 = st.columns(4)
clt_reduction_weeks = comparison["transit_time_reduction_days"] / 7
containers_per_week = annual_volume / plant_weeks
kpi1.metric(
    "Working capital released (yearly, total CLT reduction)",
    money_up(cargo_value * containers_per_week * clt_reduction_weeks, currency),
    help=(
        "Annual volume ÷ plant weeks × cargo value × weeks of CLT reduction "
        f"(CLT reduction ÷ 7 = {clt_reduction_weeks:,.2f} weeks; "
        f"{containers_per_week:,.2f} containers/week). "
        "This is capital released, not an annual P&L saving."
    ),
)
kpi1_week.metric(
    "Working capital released per 1 week of CLT reduction",
    money_up(cargo_value * containers_per_week, currency),
    help=(
        "Annual volume ÷ plant weeks × cargo value. Capital released for each "
        "week (7 days) that the CLT is reduced."
    ),
)
kpi2.metric(
    "Annual financing savings",
    money_up(comparison["annual_financing_savings"], currency),
    help="Annual reduction in inventory financing cost; not the capital released.",
)
kpi3.metric(
    "Net annual economic benefit",
    money_up(comparison["net_annual_economic_benefit"], currency),
    delta=(
        "Scenario B preferable"
        if comparison["net_annual_economic_benefit"] > 0
        else "Scenario A preferable"
        if comparison["net_annual_economic_benefit"] < 0
        else "Break-even"
    ),
)

st.subheader("Route business case")
case_col1, case_col2, case_col3, case_col4 = st.columns(4)
case_col1.metric(
    "CLT reduction",
    f"{math.ceil(comparison['transit_time_reduction_days']):,} days",
    f"{math.ceil(comparison['transit_time_reduction_percent']):,}%",
)
case_col2.metric(
    "Transportation savings (annual)",
    money_up(comparison["annual_transportation_savings"], currency),
)
case_col3.metric(
    "Additional transport cost (annual)",
    money_up(-comparison["annual_transportation_savings"], currency),
)
case_col4.metric(
    "Financing-only break-even freight premium",
    money_up(comparison["maximum_freight_premium_per_container"], currency)
    + " / container",
    help=(
        "Excludes service level, safety stock, obsolescence, warehousing, stockouts "
        "and other benefits not explicitly modeled."
    ),
)

actual_premium = comparison["actual_additional_transport_cost_per_container"]
break_even_premium = comparison["maximum_freight_premium_per_container"]
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

st.subheader("Scenario KPIs")
scenario_rows = []
for name, inputs, metrics in [
    ("Scenario A", scenario_a, metrics_a),
    ("Scenario B", scenario_b, metrics_b),
]:
    scenario_rows.append(
        {
            "Metric": name,
            "CLT (days)": math.ceil(metrics.analysis_days),
            "Pipeline inventory value": money_up(metrics.pipeline_inventory_value, currency),
            "Finance cost / container": money_up(
                metrics.financing_cost_per_container, currency
            ),
            "Annual financing cost": money_up(metrics.annual_financing_cost, currency),
            "Transport cost / container": money_up(
                metrics.transportation_cost_per_container, currency
            ),
            "Annual transportation cost": money_up(
                metrics.annual_transportation_cost, currency
            ),
            "Economic cost / container": money_up(
                metrics.total_economic_cost_per_container, currency
            ),
            "Annual economic logistics cost": money_up(
                metrics.annual_total_economic_cost, currency
            ),
        }
    )
st.dataframe(pd.DataFrame(scenario_rows).set_index("Metric"), width="stretch")

with st.expander("Calculation breakdown & assumptions"):
    st.markdown(
        f"""
        - **Time basis:** Complete Customer Lead Time (CLT) of each scenario.
        - **Pipeline inventory value:** cargo value × containers consumed per day × selected days.
        - **Working capital released:** annual volume ÷ plant working weeks × cargo value × (CLT reduction ÷ 7).
        - **Per-container financing cost:** cargo value × selected annual rate × selected days ÷ {annual_days}.
        - **Annual financing cost:** pipeline inventory value × selected annual rate.
        - **Transportation cost per container:** ocean freight + inland transportation + additional logistics.
        - **Net annual economic benefit:** Scenario A annual economic logistics cost − Scenario B annual economic logistics cost.
        - **Finance rate used:** {financing_rate:.2%} ({'inventory carrying cost' if use_carrying_rate else 'WACC'}).

        Inventory value and released working capital are balance-sheet/cash-flow
        measures, not P&L savings. Safety stock, service-level benefits, obsolescence,
        warehousing, demurrage, detention and stockout costs are not modeled.
        """
    )

st.divider()
st.header("Charts & sensitivity analysis")

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
    annual_cost_chart = px.bar(
        annual_cost_data,
        x="Scenario",
        y="Annual cost",
        color="Cost type",
        barmode="group",
        title="Annual transportation and financing costs",
        labels={"Annual cost": f"Annual cost ({currency})"},
    )
    annual_cost_chart.update_traces(
        texttemplate="%{y:,.0f}", textposition="outside", cliponaxis=False
    )
    st.plotly_chart(annual_cost_chart, width="stretch")
with chart_col2:
    pipeline_chart = px.bar(
        pd.DataFrame(
            {
                "Scenario": ["Scenario A", "Scenario B"],
                "Pipeline inventory value": [
                    metrics_a.pipeline_inventory_value,
                    metrics_b.pipeline_inventory_value,
                ],
            }
        ),
        x="Scenario",
        y="Pipeline inventory value",
        title="Pipeline inventory value (capital tied up)",
        labels={"Pipeline inventory value": f"Value ({currency})"},
    )
    pipeline_chart.update_traces(
        texttemplate="%{y:,.0f}", textposition="outside", cliponaxis=False
    )
    st.plotly_chart(pipeline_chart, width="stretch")

total_cost_chart = px.bar(
    pd.DataFrame(
        {
            "Scenario": ["Scenario A", "Scenario B"],
            "Annual economic logistics cost": [
                metrics_a.annual_total_economic_cost,
                metrics_b.annual_total_economic_cost,
            ],
        }
    ),
    x="Scenario",
    y="Annual economic logistics cost",
    title="Annual total economic logistics cost",
    labels={"Annual economic logistics cost": f"Annual cost ({currency})"},
)
total_cost_chart.update_traces(
    texttemplate="%{y:,.0f}", textposition="outside", cliponaxis=False
)
st.plotly_chart(total_cost_chart, width="stretch")

sensitivity_col1, sensitivity_col2 = st.columns(2)
with sensitivity_col1:
    max_days = max(metrics_a.analysis_days, metrics_b.analysis_days, 1)
    sensitivity_days = st.slider(
        "CLT range for financing sensitivity (days)",
        min_value=1,
        max_value=max(30, int(max_days * 1.5)),
        value=(1, max(30, int(max_days * 1.5))),
    )
    time_range = list(range(sensitivity_days[0], sensitivity_days[1] + 1))
    financing_sensitivity = pd.DataFrame(
        {
            "Days": time_range * 2,
            "Scenario": ["Scenario A"] * len(time_range)
            + ["Scenario B"] * len(time_range),
            "Annual financing cost": [
                cargo_value
                * containers_per_day
                * day
                * financing_rate
                for day in time_range
            ]
            * 2,
        }
    )
    st.plotly_chart(
        px.line(
            financing_sensitivity,
            x="Days",
            y="Annual financing cost",
            color="Scenario",
            title="Annual financing cost vs. CLT",
            labels={"Annual financing cost": f"Annual cost ({currency})"},
        ),
        width="stretch",
    )

with sensitivity_col2:
    custom_reduction = st.number_input(
        "Custom lead-time reduction (days)",
        min_value=0.0,
        value=10.0,
        step=1.0,
    )
    premium_for_sensitivity = st.number_input(
        "Additional freight premium to test per container",
        min_value=0.0,
        value=100.0,
        step=25.0,
        help="Used only in this sensitivity view to calculate net annual benefit.",
    )
    reductions = sorted(set([1.0, 5.0, 10.0, 15.0, 20.0, custom_reduction]))
    sensitivity = calculate_sensitivity(
        reductions_days=reductions,
        cargo_value_per_container=cargo_value,
        containers_consumed_per_day=containers_per_day,
        annual_container_volume=annual_volume,
        annual_financing_rate=financing_rate,
        additional_freight_premium_per_container=premium_for_sensitivity,
        days_per_year=annual_days,
    )
    sensitivity_df = pd.DataFrame(sensitivity)
    sensitivity_chart = go.Figure()
    sensitivity_chart.add_trace(
        go.Scatter(
            x=sensitivity_df["reduction_days"],
            y=sensitivity_df["net_annual_benefit_after_premium"],
            mode="lines+markers",
            name="Net annual benefit",
        )
    )
    sensitivity_chart.add_hline(y=0, line_dash="dash", line_color="gray")
    sensitivity_chart.update_layout(
        title="Net annual benefit by lead-time reduction",
        xaxis_title="Lead-time reduction (days)",
        yaxis_title=f"Net annual benefit ({currency})",
    )
    st.plotly_chart(sensitivity_chart, width="stretch")
    st.dataframe(
        sensitivity_df.rename(
            columns={
                "reduction_days": "Reduction (days)",
                "working_capital_released": "Working capital released",
                "annual_financing_savings": "Annual financing savings",
                "maximum_freight_premium_per_container": (
                    "Maximum freight premium / container"
                ),
                "net_annual_benefit_after_premium": (
                    "Net annual benefit after premium"
                ),
            }
        ).set_index("Reduction (days)").style.format(
            {
                "Working capital released": lambda value: money(value, currency),
                "Annual financing savings": lambda value: money(value, currency),
                "Maximum freight premium / container": (
                    lambda value: money(value, currency)
                ),
                "Net annual benefit after premium": (
                    lambda value: money(value, currency)
                ),
            }
        ),
        width="stretch",
    )
