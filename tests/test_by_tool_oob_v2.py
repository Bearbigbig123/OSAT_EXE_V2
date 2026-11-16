import inspect
import unittest

import numpy as np
import pandas as pd

from by_tool_oob_analysis import (
    apply_whole_baseline_exemption,
    build_by_tool_plot_styles,
    by_tool_plot_alpha,
    evaluate_by_tool_new_rules,
    evaluate_tool_new_rule,
    filter_by_tool_plot_data,
    format_by_tool_shift_label,
    has_eligible_by_tool_shift,
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

    def test_missing_peer_baseline_blocks_k_up_to_5(self):
        final, exemptions, _ = apply_whole_baseline_exemption(
            result_for(k=3.0, rules=("P50",)),
            {"P95": 10.0, "P50": 10.0, "P05": 10.0},
            {"P95": np.nan, "P05": np.nan},
        )
        self.assertEqual(final["P50_shift"], "NO_HIGHLIGHT")
        self.assertIn("fewer than 10 points", exemptions[0])

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


if __name__ == "__main__":
    unittest.main()
