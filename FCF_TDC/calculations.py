"""Financial calculations for the maritime lead-time comparison."""

from dataclasses import dataclass
from math import isfinite


DEFAULT_DAYS_PER_YEAR = 365


@dataclass(frozen=True)
class ScenarioInputs:
    ocean_freight_per_container: float
    inland_transport_per_container: float
    additional_logistics_per_container: float
    customer_lead_time_days: float

    @property
    def transportation_cost_per_container(self) -> float:
        return (
            self.ocean_freight_per_container
            + self.inland_transport_per_container
            + self.additional_logistics_per_container
        )

    @property
    def analysis_days(self) -> float:
        return self.customer_lead_time_days


@dataclass(frozen=True)
class ScenarioMetrics:
    analysis_days: float
    pipeline_inventory_value: float
    financing_cost_per_container: float
    annual_financing_cost: float
    transportation_cost_per_container: float
    annual_transportation_cost: float
    total_economic_cost_per_container: float
    annual_total_economic_cost: float


def calculate_scenario(
    scenario: ScenarioInputs,
    *,
    cargo_value_per_container: float,
    containers_consumed_per_day: float,
    annual_container_volume: float,
    annual_financing_rate: float,
    days_per_year: int = DEFAULT_DAYS_PER_YEAR,
) -> ScenarioMetrics:
    """Calculate pipeline capital and annual logistics costs for one scenario."""
    nonnegative_values = {
        "cargo value per container": cargo_value_per_container,
        "containers consumed per day": containers_consumed_per_day,
        "annual container volume": annual_container_volume,
        "annual financing rate": annual_financing_rate,
        "ocean freight": scenario.ocean_freight_per_container,
        "inland transportation": scenario.inland_transport_per_container,
        "additional logistics cost": scenario.additional_logistics_per_container,
        "customer lead time": scenario.customer_lead_time_days,
    }
    invalid_values = [
        name
        for name, value in nonnegative_values.items()
        if not isfinite(value) or value < 0
    ]
    if invalid_values:
        raise ValueError(
            f"Values must be finite and nonnegative: {', '.join(invalid_values)}."
        )
    if annual_financing_rate > 1:
        raise ValueError("Annual financing rate must be entered as a decimal from 0 to 1.")
    if days_per_year <= 0:
        raise ValueError("Days per year must be greater than zero.")

    analysis_days = scenario.analysis_days
    pipeline_inventory_value = (
        cargo_value_per_container
        * containers_consumed_per_day
        * analysis_days
    )
    financing_cost_per_container = (
        cargo_value_per_container
        * annual_financing_rate
        * analysis_days
        / days_per_year
    )
    annual_financing_cost = pipeline_inventory_value * annual_financing_rate
    transportation_cost_per_container = scenario.transportation_cost_per_container
    annual_transportation_cost = (
        transportation_cost_per_container * annual_container_volume
    )

    return ScenarioMetrics(
        analysis_days=analysis_days,
        pipeline_inventory_value=pipeline_inventory_value,
        financing_cost_per_container=financing_cost_per_container,
        annual_financing_cost=annual_financing_cost,
        transportation_cost_per_container=transportation_cost_per_container,
        annual_transportation_cost=annual_transportation_cost,
        total_economic_cost_per_container=(
            transportation_cost_per_container + financing_cost_per_container
        ),
        annual_total_economic_cost=(
            annual_transportation_cost + annual_financing_cost
        ),
    )


def calculate_comparison(
    scenario_a: ScenarioMetrics,
    scenario_b: ScenarioMetrics,
    *,
    annual_container_volume: float,
) -> dict[str, float]:
    """Return B-vs-A savings and the financing-only freight break-even point."""
    working_capital_released = (
        scenario_a.pipeline_inventory_value - scenario_b.pipeline_inventory_value
    )
    annual_financing_savings = (
        scenario_a.annual_financing_cost - scenario_b.annual_financing_cost
    )
    transportation_savings_per_container = (
        scenario_a.transportation_cost_per_container
        - scenario_b.transportation_cost_per_container
    )
    annual_transportation_savings = (
        scenario_a.annual_transportation_cost
        - scenario_b.annual_transportation_cost
    )
    maximum_freight_premium_per_container = (
        scenario_a.financing_cost_per_container
        - scenario_b.financing_cost_per_container
    )
    return {
        "transit_time_reduction_days": (
            scenario_a.analysis_days - scenario_b.analysis_days
        ),
        "transit_time_reduction_percent": (
            (scenario_a.analysis_days - scenario_b.analysis_days)
            / scenario_a.analysis_days
            * 100
            if scenario_a.analysis_days
            else 0.0
        ),
        "working_capital_released": working_capital_released,
        "annual_financing_savings": annual_financing_savings,
        "transportation_savings_per_container": transportation_savings_per_container,
        "annual_transportation_savings": annual_transportation_savings,
        "net_annual_economic_benefit": (
            scenario_a.annual_total_economic_cost
            - scenario_b.annual_total_economic_cost
        ),
        "maximum_freight_premium_per_container": maximum_freight_premium_per_container,
        "annual_maximum_freight_budget": (
            maximum_freight_premium_per_container * annual_container_volume
        ),
        "actual_additional_transport_cost_per_container": (
            scenario_b.transportation_cost_per_container
            - scenario_a.transportation_cost_per_container
        ),
        "actual_additional_annual_transport_cost": (
            scenario_b.annual_transportation_cost
            - scenario_a.annual_transportation_cost
        ),
    }


def calculate_sensitivity(
    *,
    reductions_days: list[float],
    cargo_value_per_container: float,
    containers_consumed_per_day: float,
    annual_container_volume: float,
    annual_financing_rate: float,
    additional_freight_premium_per_container: float,
    days_per_year: int = DEFAULT_DAYS_PER_YEAR,
) -> list[dict[str, float]]:
    """Estimate capital release and annual economics for each lead-time reduction."""
    sensitivity_inputs = (
        cargo_value_per_container,
        containers_consumed_per_day,
        annual_container_volume,
        annual_financing_rate,
        additional_freight_premium_per_container,
        *reductions_days,
    )
    if any(not isfinite(value) or value < 0 for value in sensitivity_inputs):
        raise ValueError("Sensitivity inputs must be finite and nonnegative.")
    if annual_financing_rate > 1:
        raise ValueError("Annual financing rate must be entered as a decimal from 0 to 1.")
    if days_per_year <= 0:
        raise ValueError("Days per year must be greater than zero.")

    results = []
    for reduction in reductions_days:
        if reduction < 0:
            raise ValueError("Lead-time reductions cannot be negative.")
        capital_released = (
            cargo_value_per_container * containers_consumed_per_day * reduction
        )
        financing_savings = capital_released * annual_financing_rate
        maximum_premium = (
            cargo_value_per_container
            * annual_financing_rate
            * reduction
            / days_per_year
        )
        results.append(
            {
                "reduction_days": reduction,
                "working_capital_released": capital_released,
                "annual_financing_savings": financing_savings,
                "maximum_freight_premium_per_container": maximum_premium,
                "net_annual_benefit_after_premium": (
                    financing_savings
                    - additional_freight_premium_per_container
                    * annual_container_volume
                ),
            }
        )
    return results
