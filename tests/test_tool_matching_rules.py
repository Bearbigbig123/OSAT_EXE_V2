import unittest

import numpy as np
import pandas as pd

from tool_matching_widget import (
    ToolMatchingWidget,
    filter_plot_eligible_tools,
    format_matching_value,
    matching_alarm_result,
    matching_axis_layout,
    matching_value_tooltip,
    normalize_matching_groups,
    remove_tool_outliers,
)


class ToolMatchingRuleTests(unittest.TestCase):
    def test_excel_images_use_filter_safe_move_and_size_anchors(self):
        import tempfile
        import zipfile
        from pathlib import Path
        from xml.etree import ElementTree
        from matplotlib.figure import Figure
        from openpyxl import load_workbook

        scatter = Figure()
        scatter.add_subplot(111).plot([0, 1], [0, 1])
        box = Figure()
        box.add_subplot(111).boxplot([[1, 2, 3]])

        class ExportContext:
            chart_figures = {("G", "C"): {"scatter": scatter, "box": box}}

            def __init__(self, folder):
                self.folder = Path(folder)

            def _get_resource_path(self, name):
                return str(self.folder / name)

        row = {
            "Need_matching": False, "AbnormalType": "", "GroupName": "G", "ChartName": "C",
            "matching_group": "Tool01", "mean_matching_index": 0.1,
            "sigma_matching_index": 0.2, "K": 1.0, "mean": 10.0, "sigma": 1.0,
            "mean_median": 10.0, "sigma_median": 1.0, "samplesize": 10,
            "characteristic": "Nominal",
        }
        with tempfile.TemporaryDirectory() as folder:
            output = ToolMatchingWidget._export_to_excel(ExportContext(folder), [row, row])
            workbook = load_workbook(output, read_only=False)
            self.assertEqual(workbook["Tool Matching Results"].auto_filter.ref, "A1:P3")

            with zipfile.ZipFile(output) as archive:
                drawing = ElementTree.fromstring(archive.read("xl/drawings/drawing1.xml"))
            namespace = {"xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"}
            anchors = drawing.findall("xdr:twoCellAnchor", namespace)
            self.assertEqual(len(anchors), 4)
            self.assertTrue(all(anchor.get("editAs") is None for anchor in anchors))

    def test_outliers_are_removed_per_tool_not_across_tools(self):
        frame = pd.DataFrame({
            "matching_group": ["A"] * 11 + ["B"] * 10,
            "point_val": [0.0] * 10 + [100.0] + [1000.0] * 10,
        })

        cleaned, outliers = remove_tool_outliers(frame)

        self.assertEqual(len(outliers), 1)
        self.assertEqual(outliers.iloc[0]["matching_group"], "A")
        self.assertEqual(outliers.iloc[0]["point_val"], 100.0)
        self.assertEqual(len(cleaned[cleaned["matching_group"] == "B"]), 10)

    def test_zero_sigma_and_single_point_groups_are_kept(self):
        frame = pd.DataFrame({
            "matching_group": ["A", "A", "A", "B"],
            "point_val": [5.0, 5.0, 5.0, 999.0],
        })

        cleaned, outliers = remove_tool_outliers(frame)

        self.assertTrue(outliers.empty)
        pd.testing.assert_frame_equal(cleaned, frame)

    def test_missing_and_blank_matching_groups_become_unknown(self):
        frame = pd.DataFrame({
            "matching_group": [np.nan, None, "", "   ", "Tool01"],
            "point_val": [1, 2, 3, 4, 5],
        })

        normalized = normalize_matching_groups(frame)

        self.assertEqual(
            normalized["matching_group"].tolist(),
            ["Unknown", "Unknown", "Unknown", "Unknown", "Tool01"],
        )

    def test_plot_excludes_tools_with_fewer_than_five_usable_points(self):
        frame = pd.DataFrame({
            "matching_group": ["A"] * 4 + ["B"] * 5,
            "point_val": list(range(4)) + list(range(5)),
        })

        plot_data = filter_plot_eligible_tools(frame)

        self.assertEqual(set(plot_data["matching_group"]), {"B"})
        self.assertEqual(len(plot_data), 5)

    def test_three_plus_tool_directional_mean_rules(self):
        bigger_low = ToolMatchingWidget.calculate_mean_index(None, 8.0, 10.0, 2.0, "Bigger")
        bigger_high = ToolMatchingWidget.calculate_mean_index(None, 12.0, 10.0, 2.0, "Bigger")
        smaller_high = ToolMatchingWidget.calculate_mean_index(None, 12.0, 10.0, 2.0, "Smaller")
        smaller_low = ToolMatchingWidget.calculate_mean_index(None, 8.0, 10.0, 2.0, "Smaller")

        self.assertGreaterEqual(bigger_low, 1.0)
        self.assertLess(bigger_high, 1.0)
        self.assertGreaterEqual(smaller_high, 1.0)
        self.assertLess(smaller_low, 1.0)

    def test_zero_sigma_mean_rule_keeps_directionality(self):
        bigger_low = ToolMatchingWidget.calculate_mean_index(None, 8.0, 10.0, 0.0, "Bigger")
        bigger_high = ToolMatchingWidget.calculate_mean_index(None, 12.0, 10.0, 0.0, "Bigger")
        smaller_high = ToolMatchingWidget.calculate_mean_index(None, 12.0, 10.0, 0.0, "Smaller")
        smaller_low = ToolMatchingWidget.calculate_mean_index(None, 8.0, 10.0, 0.0, "Smaller")
        nominal_equal = ToolMatchingWidget.calculate_mean_index(None, 10.0, 10.0, 0.0, "Nominal")

        self.assertEqual(bigger_low, float("inf"))
        self.assertEqual(bigger_high, 0.0)
        self.assertEqual(smaller_high, float("inf"))
        self.assertEqual(smaller_low, 0.0)
        self.assertEqual(nominal_equal, 0.0)

    def test_two_zero_sigma_groups_never_raise_sigma_index(self):
        class Harness:
            get_k_value = ToolMatchingWidget.get_k_value
            calculate_mean_index = ToolMatchingWidget.calculate_mean_index

        stats = pd.DataFrame({
            "matching_group": ["Low", "High"],
            "mean": [8.0, 10.0],
            "std": [0.0, 0.0],
            "count": [5, 5],
        })
        results = []

        ToolMatchingWidget._analyze_two_groups(
            Harness(), stats, "G", "C", "Bigger", results
        )

        self.assertEqual(results[0][4], float("inf"))
        self.assertEqual(results[1][4], 0.0)
        self.assertEqual(results[0][5], 0.0)
        self.assertEqual(results[1][5], 0.0)

    def test_nonzero_sigma_is_mismatch_against_zero_sigma_reference(self):
        class Harness:
            get_k_value = ToolMatchingWidget.get_k_value
            calculate_mean_index = ToolMatchingWidget.calculate_mean_index

        stats = pd.DataFrame({
            "matching_group": ["Stable", "Variable"],
            "mean": [10.0, 10.0],
            "std": [0.0, 1.0],
            "count": [5, 5],
        })
        results = []

        ToolMatchingWidget._analyze_two_groups(
            Harness(), stats, "G", "C", "Nominal", results
        )

        self.assertEqual(results[0][5], 0.0)
        self.assertEqual(results[1][5], float("inf"))

    def test_matching_values_use_compact_high_precision_display(self):
        self.assertEqual(format_matching_value(1.41421356, index_value=True), "1.414")
        self.assertEqual(format_matching_value(1.41421356, index_value=True, detailed=True), "1.414214")
        self.assertEqual(format_matching_value(0.000012345678), "1.2346e-05")
        self.assertEqual(format_matching_value(123.456789), "123.46")
        self.assertEqual(matching_value_tooltip(1.41421356), "Full precision: 1.41421356")

    def test_analysis_keeps_raw_precision_before_ui_formatting(self):
        class Harness:
            get_k_value = ToolMatchingWidget.get_k_value
            calculate_mean_index = ToolMatchingWidget.calculate_mean_index

        stats = pd.DataFrame({
            "matching_group": ["A", "B"],
            "mean": [1.123456789, 1.223456789],
            "std": [0.123456789, 0.223456789],
            "count": [11, 11],
        })
        results = []
        ToolMatchingWidget._analyze_two_groups(
            Harness(), stats, "G", "C", "Nominal", results
        )

        self.assertEqual(results[0][6], 1.414)
        self.assertEqual(results[0][7], 1.123456789)
        self.assertEqual(results[0][8], 0.123456789)

    def test_matching_plot_alarm_and_twenty_tool_layout(self):
        settings = {
            "mean_index_enabled": False,
            "sigma_index_enabled": False,
        }
        self.assertEqual(
            matching_alarm_result(1.2, 1.0, 1.73, 0.5, 0.5, settings),
            (True, "Mean"),
        )
        self.assertEqual(
            matching_alarm_result(0.2, 1.0, 1.73, 0.5, 0.5, settings),
            (False, ""),
        )
        layout = matching_axis_layout(20)
        self.assertEqual(layout["rotation"], 45)
        self.assertEqual(matching_axis_layout(5)["rotation"], 45)
        self.assertEqual(matching_axis_layout(30)["rotation"], 45)
        self.assertGreaterEqual(layout["figure_width"], 10.0)
        self.assertEqual(layout["legend_columns"], 2)

    def test_alarm_tool_is_emphasized_without_background_shading(self):
        class PlotContext:
            chart_figures = {}
            tool_matching_settings = {
                "mean_index_enabled": False,
                "sigma_index_enabled": False,
            }

        frame = pd.DataFrame({
            "matching_group": ["AlarmTool"] * 5 + ["NormalTool"] * 5,
            "point_val": [10, 11, 12, 13, 14, 10, 10.1, 9.9, 10.2, 9.8],
            "point_time": pd.date_range("2026-01-01", periods=10),
        })
        results = [
            ["G", "C", "AlarmTool", "group_all", 1.2, 1.0, 1.73, 12.0, 0.5, 10.0, 0.5, 5, "Nominal"],
            ["G", "C", "NormalTool", "group_all", 0.2, 1.0, 1.73, 10.0, 0.5, 10.0, 0.5, 5, "Nominal"],
        ]

        context = PlotContext()
        ToolMatchingWidget._create_boxplot_single(context, frame, "G", "C", results)
        scatter_figure = context.chart_figures[("G", "C")]["scatter"]
        box_figure = context.chart_figures[("G", "C")]["box"]
        scatter_axis = scatter_figure.axes[0]
        box_axis = box_figure.axes[0]

        self.assertEqual(len(scatter_axis.patches), 0)
        self.assertEqual(len(box_axis.patches), 2)
        self.assertEqual([collection.get_alpha() for collection in scatter_axis.collections], [0.95, 0.45])
        self.assertTrue(all(label.get_rotation() == 45 for label in scatter_axis.get_xticklabels()))
        self.assertTrue(all(label.get_rotation() == 45 for label in box_axis.get_xticklabels()))


if __name__ == "__main__":
    unittest.main()
