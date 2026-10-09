import unittest

from calculations import (
    ScenarioInputs,
    calculate_comparison,
    calculate_scenario,
    calculate_sensitivity,
)


class MaritimeCalculationTests(unittest.TestCase):
    def setUp(self):
        self.scenario_a = ScenarioInputs(
            ocean_freight_per_container=1500,
            inland_transport_per_container=400,
            additional_logistics_per_container=0,
            customer_lead_time_days=50,
        )
        self.scenario_b = ScenarioInputs(
            ocean_freight_per_container=1600,
            inland_transport_per_container=400,
            additional_logistics_per_container=0,
            customer_lead_time_days=40,
        )
        self.common = {
            "cargo_value_per_container": 100_000,
            "containers_consumed_per_day": 1,
            "annual_container_volume": 365,
            "annual_financing_rate": 0.08,
        }

    def test_validation_example_scenario_metrics(self):
        result_a = calculate_scenario(self.scenario_a, **self.common)
        result_b = calculate_scenario(self.scenario_b, **self.common)

        self.assertEqual(result_a.pipeline_inventory_value, 5_000_000)
        self.assertAlmostEqual(result_a.annual_financing_cost, 400_000)
        self.assertAlmostEqual(result_a.financing_cost_per_container, 1095.8904109589)
        self.assertEqual(result_a.transportation_cost_per_container, 1900)
        self.assertAlmostEqual(result_a.total_economic_cost_per_container, 2995.8904109589)
        self.assertEqual(result_a.annual_total_economic_cost, 1_093_500)

        self.assertEqual(result_b.pipeline_inventory_value, 4_000_000)
        self.assertAlmostEqual(result_b.annual_financing_cost, 320_000)
        self.assertAlmostEqual(result_b.financing_cost_per_container, 876.7123287671)
        self.assertEqual(result_b.transportation_cost_per_container, 2000)
        self.assertAlmostEqual(result_b.total_economic_cost_per_container, 2876.7123287671)
        self.assertEqual(result_b.annual_total_economic_cost, 1_050_000)

    def test_validation_example_comparison(self):
        result_a = calculate_scenario(self.scenario_a, **self.common)
        result_b = calculate_scenario(self.scenario_b, **self.common)
        comparison = calculate_comparison(
            result_a, result_b, annual_container_volume=365
        )

        self.assertEqual(comparison["working_capital_released"], 1_000_000)
        self.assertEqual(comparison["annual_financing_savings"], 80_000)
        self.assertEqual(comparison["actual_additional_transport_cost_per_container"], 100)
        self.assertEqual(comparison["actual_additional_annual_transport_cost"], 36_500)
        self.assertEqual(comparison["net_annual_economic_benefit"], 43_500)
        self.assertAlmostEqual(
            comparison["maximum_freight_premium_per_container"], 219.1780821918
        )

    def test_analysis_days_is_clt(self):
        result = calculate_scenario(self.scenario_a, **self.common)
        self.assertEqual(result.analysis_days, 50)

    def test_negative_inputs_are_rejected(self):
        with self.assertRaisesRegex(ValueError, "finite and nonnegative"):
            calculate_scenario(
                self.scenario_a,
                **{**self.common, "cargo_value_per_container": -1},
            )

    def test_non_finite_inputs_are_rejected(self):
        with self.assertRaisesRegex(ValueError, "finite and nonnegative"):
            calculate_scenario(
                self.scenario_a,
                **{**self.common, "cargo_value_per_container": float("nan")},
            )

    def test_sensitivity_results(self):
        results = calculate_sensitivity(
            reductions_days=[10],
            cargo_value_per_container=100_000,
            containers_consumed_per_day=1,
            annual_container_volume=365,
            annual_financing_rate=0.08,
            additional_freight_premium_per_container=100,
        )
        self.assertEqual(results[0]["working_capital_released"], 1_000_000)
        self.assertEqual(results[0]["annual_financing_savings"], 80_000)
        self.assertAlmostEqual(
            results[0]["maximum_freight_premium_per_container"], 219.1780821918
        )
        self.assertEqual(results[0]["net_annual_benefit_after_premium"], 43_500)


if __name__ == "__main__":
    unittest.main()
