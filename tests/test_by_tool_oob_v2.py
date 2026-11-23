import inspect
import unittest

import numpy as np
import pandas as pd

from by_tool_oob_analysis import (
    apply_whole_baseline_exemption,
    build_by_tool_plot_styles,
    by_tool_plot_alpha,
    characteristic_profile,
    default_by_tool_percentile_profile,
    evaluate_by_tool_new_rules,
    evaluate_tool_new_rule,
    filter_by_tool_plot_data,
    format_by_tool_shift_label,
    has_eligible_by_tool_shift,
    normalize_by_tool_percentile_profile,
)


def result_for(k=3.0, rules=("P50",), signed_p50=3.0):
    result = {}
    for percentile in ("P95", "P50", "P05"):
        result[f"{percentile}_k"] = k
        result[f"{percentile}_k_ori"] = signed_p50 if percentile == "P50" else k
        result[f"{percentile}_shift"] = "HIGHLIGHT" if percentile in rules else "NO_HIGHLIGHT"
    return result


def stub(result=None):
    value = result or result_for(k=0, rules=())

    def calculator(base, data, characteristic, resolution, ucl, lcl, k_threshold=1.67):
        if k_threshold != 2.0:
            raise AssertionError("By Tool must use K > 2")
        return pd.Series(value)
    return calculator


class CurrentByToolRuleTests(unittest.TestCase):
    def setUp(self):
        self.chart = {"Characteristics": "Nominal", "Resolution": 0.01, "UCL": 10, "LCL": -10}
        self.bounds = {"P95": 2.0, "P05": -2.0}

    def evaluate(self, weekly_n, baseline_n):
        return evaluate_tool_new_rule(
            "A", np.arange(baseline_n, dtype=float), np.ones(weekly_n),
            self.bounds, self.chart, stub(),
        )

    def test_weekly_4_borrows_one_and_leaves_at_least_10(self):
        row = self.evaluate(4, 11)
        self.assertTrue(row["Eligible"])
        self.assertTrue(row["Rolling Used"])
        self.assertEqual(row["Rolling Points Borrowed"], 1)
        self.assertEqual(row["Calculation Weekly Count"], 5)
        self.assertEqual(row["Calculation Baseline Count"], 10)

    def test_borrowing_that_leaves_9_is_not_eligible(self):
        row = self.evaluate(4, 10)
        self.assertFalse(row["Eligible"])
        self.assertEqual(row["Highlight"], "NOT_ELIGIBLE")
        self.assertIn("9<10", row["Ineligible Reason"])

    def test_weekly_1_needs_14_baseline_points(self):
        self.assertFalse(self.evaluate(1, 13)["Eligible"])
        self.assertTrue(self.evaluate(1, 14)["Eligible"])

    def test_weekly_5_baseline_10_is_eligible_without_rolling(self):
        row = self.evaluate(5, 10)
        self.assertTrue(row["Eligible"])
        self.assertFalse(row["Rolling Used"])

    def test_weekly_zero_is_not_eligible(self):
        self.assertFalse(self.evaluate(0, 100)["Eligible"])

    def test_latest_baseline_points_are_borrowed_and_removed(self):
        calls = []

        def capture(base, data, characteristic, resolution, ucl, lcl, k_threshold=1.67):
            calls.append((base["values"].copy(), data["values"].copy()))
            return pd.Series(result_for(k=0, rules=()))

        row = evaluate_tool_new_rule(
            "A", np.arange(12, dtype=float), np.array([20.0, 21.0, 22.0]),
            self.bounds, self.chart, capture,
        )
        self.assertTrue(row["Eligible"])
        np.testing.assert_array_equal(calls[0][0], np.arange(10, dtype=float))
        np.testing.assert_array_equal(calls[0][1], np.array([10.0, 11.0, 20.0, 21.0, 22.0]))

    def test_confirmatory_rolling_does_not_promote_weekly_3_old_no_hl(self):
        def changes_after_rolling(base, data, characteristic, resolution, ucl, lcl, k_threshold=1.67):
            rules = ("P50",) if len(data["values"]) == 5 else ()
            return pd.Series(result_for(k=3, rules=rules))

        row = evaluate_tool_new_rule(
            "A", np.arange(12, dtype=float), np.array([20.0, 21.0, 22.0]),
            {"P95": 5.0, "P05": -5.0}, self.chart, changes_after_rolling,
        )
        self.assertEqual(row["P50 HL"], "NO_HIGHLIGHT")

    def test_k_boundaries(self):
        weekly = {"P95": 1.0, "P50": 1.0, "P05": -1.0}
        for k, expected in ((2.0, "NO_HIGHLIGHT"), (3.0, "NO_HIGHLIGHT"),
                            (5.0, "NO_HIGHLIGHT"), (5.01, "HIGHLIGHT")):
            final, _, _ = apply_whole_baseline_exemption(
                result_for(k=k, rules=("P50",)), weekly, self.bounds
            )
            self.assertEqual(final["P50_shift"], expected)

    def test_custom_alarm_and_peer_bypass_k_thresholds(self):
        weekly = {"P95": 1.0, "P50": 1.0, "P05": -1.0}
        peer_blocked, _, _ = apply_whole_baseline_exemption(
            result_for(k=4.0, rules=("P50",)), weekly, self.bounds,
            k_threshold=3.0, peer_bypass_k=5.0,
        )
        bypassed, _, bypassed_rules = apply_whole_baseline_exemption(
            result_for(k=4.0, rules=("P50",)), weekly, self.bounds,
            k_threshold=3.0, peer_bypass_k=3.5,
        )
        below_alarm, _, _ = apply_whole_baseline_exemption(
            result_for(k=2.5, rules=("P50",)), weekly, self.bounds,
            k_threshold=3.0, peer_bypass_k=5.0,
        )

        self.assertEqual(peer_blocked["P50_shift"], "NO_HIGHLIGHT")
        self.assertEqual(bypassed["P50_shift"], "HIGHLIGHT")
        self.assertEqual(bypassed_rules, ["P50_shift"])
        self.assertEqual(below_alarm["P50_shift"], "NO_HIGHLIGHT")

    def test_chart_alarm_k_threshold_is_forwarded_to_calculator(self):
        received = []

        def capture_threshold(base, data, characteristic, resolution, ucl, lcl,
                              k_threshold=1.67, rule_profile=None):
            received.append(k_threshold)
            return pd.Series(result_for(k=0, rules=()))

        chart = dict(self.chart, by_tool_median_shift_k_threshold=3.25)
        evaluate_tool_new_rule(
            "A", np.arange(10, dtype=float), np.arange(5, dtype=float),
            self.bounds, chart, capture_threshold,
        )
        self.assertEqual(received, [3.25])

    def test_missing_peer_baseline_blocks_k_up_to_5(self):
        final, exemptions, _ = apply_whole_baseline_exemption(
            result_for(k=3.0, rules=("P50",)),
            {"P95": 10.0, "P50": 10.0, "P05": 10.0},
            {"P95": np.nan, "P05": np.nan},
        )
        self.assertEqual(final["P50_shift"], "NO_HIGHLIGHT")
        self.assertIn("fewer than 10 points", exemptions[0])

    def test_market_switch_off_skips_peer_gate(self):
        final, exemptions, bypassed = apply_whole_baseline_exemption(
            result_for(k=3.0, rules=("P50",)),
            {"P95": 1.0, "P50": 1.0, "P05": -1.0},
            self.bounds,
            use_peer_gate=False,
        )
        self.assertEqual(final["P50_shift"], "HIGHLIGHT")
        self.assertEqual(exemptions, [])
        self.assertEqual(bypassed, [])

    def test_market_switch_off_skips_peer_baseline_calculation(self):
        baseline = pd.DataFrame({
            "Matching": ["A"] * 10 + ["B"] * 10,
            "point_val": [0.0] * 10 + [10.0] * 10,
        })
        weekly = pd.DataFrame({"Matching": ["A"] * 5, "point_val": [1.0] * 5})
        chart = dict(self.chart, by_tool_use_peer_baseline=False)
        detail = evaluate_by_tool_new_rules(
            baseline, weekly, chart, stub(result_for(k=3.0, rules=("P50",)))
        )
        row = detail.loc[detail["Tool"] == "A"].iloc[0]
        self.assertEqual(row["P50 HL"], "HIGHLIGHT")
        self.assertEqual(row["Whole Baseline Count"], 0)
        self.assertIn("Peer:Off", row["Rule Formulas"]["P50"])

    def test_no_alarm_reason_identifies_k_gate(self):
        row = evaluate_tool_new_rule(
            "A", np.arange(10, dtype=float), np.ones(5),
            self.bounds, self.chart, stub(result_for(k=1.5, rules=())),
        )
        self.assertIn("K=1.5", row["P50 No Alarm Reason"])
        self.assertIn("未超過門檻", row["P50 No Alarm Reason"])

    def test_no_alarm_reason_identifies_missing_market_data(self):
        row = evaluate_tool_new_rule(
            "A", np.arange(10, dtype=float), np.ones(5),
            {"P95": np.nan, "P05": np.nan, "Count": 0},
            self.chart, stub(result_for(k=3.0, rules=("P50",))),
        )
        self.assertIn("大盤資料不足", row["P50 No Alarm Reason"])

    def test_alarm_rule_has_no_no_alarm_reason(self):
        chart = dict(self.chart, by_tool_use_peer_baseline=False)
        row = evaluate_tool_new_rule(
            "A", np.arange(10, dtype=float), np.ones(5),
            self.bounds, chart, stub(result_for(k=3.0, rules=("P50",))),
        )
        self.assertEqual(row["P50 HL"], "HIGHLIGHT")
        self.assertEqual(row["P50 No Alarm Reason"], "")

    def test_excel_export_writes_readable_diagnostic_column(self):
        import tempfile
        from pathlib import Path
        from openpyxl import load_workbook
        from oob_module_NGK_nostatic import save_results_to_excel

        diagnostic = "ToolA | P95: K=1.2 未超過門檻 2; P50: ALARM; P05: 大盤資料不足"
        frame = pd.DataFrame([{
            "group_name": "G", "chart_name": "C",
            "chart_path": "N/A", "weekly_chart_path": "N/A",
            "by_tool_color_path": "N/A", "by_tool_group_path": "N/A",
            "by_tool_shift_diagnostics": diagnostic,
        }])
        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "diagnostic.xlsx"
            save_results_to_excel(frame, output_path=str(output))
            sheet = load_workbook(output, read_only=False).active
            headers = [cell.value for cell in sheet[1]]
            column = headers.index("By Tool P95/P50/P05 No Alarm Reasons") + 1
            self.assertEqual(sheet.cell(2, column).value, diagnostic)
            self.assertTrue(sheet.cell(2, column).alignment.wrap_text)

    def test_whole_baseline_excludes_current_tool(self):
        baseline = pd.DataFrame({"Matching": ["A"] * 10 + ["B"] * 10, "point_val": [0] * 10 + [10] * 10})
        weekly = pd.DataFrame({"Matching": ["A"] * 5 + ["B"] * 5, "point_val": [1] * 5 + [9] * 5})
        detail = evaluate_by_tool_new_rules(baseline, weekly, self.chart, stub())
        self.assertGreater(detail.loc[detail["Tool"] == "A", "Whole Baseline P95"].iloc[0], 9)
        self.assertLess(detail.loc[detail["Tool"] == "B", "Whole Baseline P95"].iloc[0], 1)
        self.assertTrue((detail["Whole Baseline Count"] == 10).all())

    def test_peer_whole_baseline_requires_at_least_10_points(self):
        baseline = pd.DataFrame({"Matching": ["A"] * 14 + ["B"] * 9, "point_val": [0] * 14 + [10] * 9})
        weekly = pd.DataFrame({"Matching": ["A"] * 5, "point_val": [1] * 5})
        detail = evaluate_by_tool_new_rules(baseline, weekly, self.chart, stub())
        row = detail.loc[detail["Tool"] == "A"].iloc[0]
        self.assertEqual(row["Whole Baseline Count"], 9)
        self.assertTrue(np.isnan(row["Whole Baseline P95"]))

    def test_single_tool_has_no_peer_whole_baseline(self):
        baseline = pd.DataFrame({"Matching": ["A"] * 14, "point_val": np.arange(14)})
        weekly = pd.DataFrame({"Matching": ["A"] * 5, "point_val": np.arange(5)})
        detail = evaluate_by_tool_new_rules(baseline, weekly, self.chart, stub())
        row = detail.iloc[0]
        self.assertEqual(row["Whole Baseline Count"], 0)
        self.assertTrue(np.isnan(row["Whole Baseline P95"]))

    def test_production_calculator_accepts_k_threshold(self):
        from oob_module_NGK_nostatic import kshift_sigma_ratio_calculator
        self.assertIn("k_threshold", inspect.signature(kshift_sigma_ratio_calculator).parameters)
        self.assertIn("rule_profile", inspect.signature(kshift_sigma_ratio_calculator).parameters)

    def test_current_default_profile_uses_dual_tail_gates_for_p95_and_p05(self):
        profile = default_by_tool_percentile_profile()
        nominal = profile["characteristics"]["Nominal"]
        bigger = profile["characteristics"]["Bigger"]
        smaller = profile["characteristics"]["Smaller"]

        self.assertEqual((nominal["P50"]["own_low"], nominal["P50"]["own_high"]), (25, 75))
        for rules in (nominal, bigger, smaller):
            for slot in ("P95", "P05"):
                self.assertTrue(rules[slot]["own_enabled"])
                self.assertEqual((rules[slot]["own_low"], rules[slot]["own_high"]), (5, 95))
                self.assertEqual((rules[slot]["peer_low"], rules[slot]["peer_high"]), (5, 95))
        self.assertEqual(characteristic_profile(profile, "Sigma"), smaller)

    def test_invalid_or_future_profile_resets_to_current_default(self):
        expected = default_by_tool_percentile_profile()
        self.assertEqual(normalize_by_tool_percentile_profile("not json"), expected)
        self.assertEqual(normalize_by_tool_percentile_profile({"version": 999}), expected)

    def test_custom_peer_percentile_changes_whole_baseline_gate(self):
        profile = default_by_tool_percentile_profile()
        rule = profile["characteristics"]["Nominal"]["P50"]
        rule["peer_high"] = 75
        weekly = {"P95": 1.0, "P50": 1.0, "P05": -1.0}
        whole = {
            "P95": 10.0,
            "P05": -10.0,
            "percentiles": {"5": -10.0, "75": 0.5, "95": 10.0},
        }

        default_result, _, _ = apply_whole_baseline_exemption(
            result_for(k=3.0, rules=("P50",)), weekly, whole
        )
        custom_result, _, _ = apply_whole_baseline_exemption(
            result_for(k=3.0, rules=("P50",)), weekly, whole,
            characteristic_profile(profile, "Nominal"),
        )

        self.assertEqual(default_result["P50_shift"], "NO_HIGHLIGHT")
        self.assertEqual(custom_result["P50_shift"], "HIGHLIGHT")

    def test_custom_own_gate_percentiles_are_used(self):
        from oob_module_NGK_nostatic import review_kshift_results

        profile = default_by_tool_percentile_profile()
        rule = profile["characteristics"]["Nominal"]["P50"]
        rule.update({"weekly": 80, "k_base": 60, "own_low": 20, "own_high": 70})
        results = result_for(k=3.0, rules=("P50",), signed_p50=3.0)
        weekly_percentiles = {"Q80": 90.0}
        baseline_percentiles = {"Q60": 50.0, "Q20": 20.0, "Q70": 80.0}

        highlights = review_kshift_results(
            results, 0.01, "Nominal", weekly_percentiles, baseline_percentiles,
            k_threshold=2.0, rule_profile=characteristic_profile(profile, "Nominal"),
        )

        self.assertEqual(highlights["P50_shift"], "HIGHLIGHT")

    def test_p95_and_p05_each_alarm_outside_either_baseline_tail(self):
        from oob_module_NGK_nostatic import review_kshift_results

        baseline = {
            "P05": 5.0, "P25": 25.0, "P50": 50.0, "P75": 75.0, "P95": 95.0,
            "Q5": 5.0, "Q25": 25.0, "Q50": 50.0, "Q75": 75.0, "Q95": 95.0,
        }
        scenarios = (("P95", 100.0, 3.0), ("P95", 0.0, -3.0),
                     ("P05", 100.0, 3.0), ("P05", 0.0, -3.0))
        for target_slot, target_value, signed_value in scenarios:
            weekly = {"P05": 50.0, "P50": 50.0, "P95": 50.0}
            weekly[target_slot] = target_value
            weekly_with_q = dict(weekly)
            weekly_with_q.update({f"Q{int(key[1:])}": value for key, value in weekly.items()})
            results = {}
            for slot in ("P95", "P50", "P05"):
                signed = signed_value if slot == target_slot else 0.0
                results[f"{slot}_k"] = abs(signed)
                results[f"{slot}_k_ori"] = signed
            with self.subTest(slot=target_slot, value=target_value):
                configured = review_kshift_results(
                    results, 0.01, "Nominal", weekly_with_q, baseline,
                    k_threshold=2.0,
                    rule_profile=characteristic_profile(None, "Nominal"),
                )
                self.assertEqual(configured[f"{target_slot}_shift"], "HIGHLIGHT")

    def test_peer_gate_for_p95_and_p05_follows_signed_direction(self):
        profile = characteristic_profile(None, "Nominal")
        bounds = {"P95": 2.0, "P05": -2.0}
        for slot, signed, weekly_value in (("P95", -3.0, -3.0), ("P05", 3.0, 3.0)):
            result = result_for(k=0, rules=())
            result[f"{slot}_shift"] = "HIGHLIGHT"
            result[f"{slot}_k"] = 3.0
            result[f"{slot}_k_ori"] = signed
            weekly = {"P95": 0.0, "P50": 0.0, "P05": 0.0}
            weekly[slot] = weekly_value
            final, _, _ = apply_whole_baseline_exemption(
                result, weekly, bounds, profile, k_threshold=2.0, peer_bypass_k=5.0,
            )
            with self.subTest(slot=slot, signed=signed):
                self.assertEqual(final[f"{slot}_shift"], "HIGHLIGHT")

    def test_by_tool_plot_excludes_ineligible_tools(self):
        frame = pd.DataFrame({
            "Matching": ["A", "A", "B", "B", "C"],
            "point_val": [1, 2, 3, 4, 5],
        })
        details = """[
            {"tool": "A", "eligible": true},
            {"tool": "B", "eligible": false},
            {"tool": "C", "eligible": true}
        ]"""

        filtered = filter_by_tool_plot_data(frame, details)

        self.assertEqual(set(filtered["Matching"]), {"A", "C"})
        self.assertEqual(len(filtered), 3)

    def test_by_tool_plot_without_shift_details_keeps_existing_behavior(self):
        frame = pd.DataFrame({"Matching": ["A", "B"], "point_val": [1, 2]})
        filtered = filter_by_tool_plot_data(frame, "[]")
        pd.testing.assert_frame_equal(filtered, frame)

    def test_by_tool_plot_with_no_eligible_tools_falls_back_to_raw_view(self):
        frame = pd.DataFrame({"Matching": ["A", "B"], "point_val": [1, 2]})
        details = """[
            {"tool": "A", "eligible": false},
            {"tool": "B", "eligible": false}
        ]"""

        filtered = filter_by_tool_plot_data(frame, details)

        pd.testing.assert_frame_equal(filtered, frame)
        self.assertFalse(has_eligible_by_tool_shift(details))

    def test_by_tool_plot_with_only_one_eligible_tool_falls_back_to_raw_view(self):
        frame = pd.DataFrame({"Matching": ["A", "B"], "point_val": [1, 2]})
        details = """[
            {"tool": "A", "eligible": true},
            {"tool": "B", "eligible": false}
        ]"""

        filtered = filter_by_tool_plot_data(frame, details)

        pd.testing.assert_frame_equal(filtered, frame)
        self.assertFalse(has_eligible_by_tool_shift(details))

    def test_multi_rule_legend_shows_each_rule_k(self):
        label = format_by_tool_shift_label("Tool01", {
            "triggered_rules": ["P50_shift", "P05_shift"],
            "p05_k": 5.45,
            "p50_k": 3.20,
            "k": 5.45,
        })
        self.assertEqual(label, "Tool01 (P05=5.45σ, P50=3.20σ)")
        self.assertNotIn("Max", label)

    def test_legend_shows_configured_formula_without_changing_rule_id(self):
        label = format_by_tool_shift_label("Tool01", {
            "triggered_rules": ["P50_shift"],
            "p50_k": 3.2,
            "rule_formulas": {"P50": "W:P80 vs K-Base:P60; outside Base P20-P70"},
        })
        self.assertIn("P50[W:P80 vs K-Base:P60; outside Base P20-P70]", label)

    def test_non_alarm_diagnostic_k_does_not_show_max_label(self):
        label = format_by_tool_shift_label("Tool01", {
            "triggered_rules": [],
            "max_percentile": "P05",
            "k": 2.34,
        })
        self.assertEqual(label, "Tool01")

    def test_large_tool_plot_styles_do_not_repeat_colors(self):
        styles = build_by_tool_plot_styles([f"Tool{i:02d}" for i in range(30)])
        colors = [style["color"] for style in styles.values()]
        self.assertEqual(len(colors), 30)
        self.assertEqual(len(set(colors)), 30)
        self.assertGreater(len({style["marker"] for style in styles.values()}), 1)

    def test_non_alarm_tool_is_transparent(self):
        self.assertEqual(by_tool_plot_alpha({"highlight": True}, True), 0.95)
        self.assertEqual(by_tool_plot_alpha({"highlight": True}, False), 0.95)
        self.assertEqual(by_tool_plot_alpha({"highlight": False}, True), 0.45)
        self.assertEqual(by_tool_plot_alpha(None, False), 0.8)


class OneSidedKDenominatorTests(unittest.TestCase):
    def calculate(self, characteristic, ucl, lcl, by_tool=True):
        from oob_module_NGK_nostatic import kshift_sigma_ratio_calculator

        baseline = np.linspace(0.0, 10.0, 101)
        weekly = np.full(5, 20.0)
        return kshift_sigma_ratio_calculator(
            {"values": baseline}, {"values": weekly},
            characteristic, 0.01, ucl, lcl,
            rule_profile=characteristic_profile(None, characteristic) if by_tool else None,
        )

    def assert_all_slots_use_scale_ten(self, result):
        self.assertAlmostEqual(result["P95_k"], 1.05)
        self.assertAlmostEqual(result["P50_k"], 1.50)
        self.assertAlmostEqual(result["P05_k"], 1.95)

    def test_bigger_uses_lcl_scale_for_all_denominators(self):
        # Baseline P50=5 and LCL=-55, so the common scale is (5-(-55))/6=10.
        self.assert_all_slots_use_scale_ten(self.calculate("Bigger", np.nan, -55.0))

    def test_smaller_uses_ucl_scale_for_all_denominators(self):
        # Baseline P50=5 and UCL=65, so the common scale is (65-5)/6=10.
        self.assert_all_slots_use_scale_ten(self.calculate("Smaller", 65.0, np.nan))

    def test_sigma_uses_smaller_ucl_behavior(self):
        self.assert_all_slots_use_scale_ten(self.calculate("Sigma", 65.0, np.nan))

    def test_general_shift_also_uses_one_sided_denominator(self):
        result = self.calculate("Bigger", np.nan, -55.0, by_tool=False)
        self.assert_all_slots_use_scale_ten(result)


if __name__ == "__main__":
    unittest.main()
