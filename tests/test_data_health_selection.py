import os
import unittest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PyQt6 import QtCore, QtWidgets

from analysis_selection import get_selected_analysis_rows
from data_health_check import DataHealthCheckWidget


class DataHealthSelectionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])

    def setUp(self):
        self.widget = DataHealthCheckWidget()
        self.widget.table.setSortingEnabled(False)
        self.widget.table.setRowCount(500)
        self.widget.analysis_chart_rows = []
        self.widget.analysis_chart_rows_by_excel_row = {}
        self.widget.analysis_log_row_states = {}
        self.widget._loading_analysis_table = True

        for index in range(500):
            excel_row = index + 2
            row_data = {
                "excel_row": excel_row,
                "group": "G",
                "chart": f"C{index}",
                "characteristics": "Nominal",
                "selected": True,
            }
            self.widget.analysis_chart_rows.append(row_data)
            self.widget.analysis_chart_rows_by_excel_row[excel_row] = row_data
            self.widget.analysis_log_row_states[excel_row] = True
            item = QtWidgets.QTableWidgetItem()
            item.setData(QtCore.Qt.ItemDataRole.UserRole, excel_row)
            self.widget.table.setItem(index, 0, item)

        self.widget._loading_analysis_table = False
        self.widget.update_analysis_checkboxes_in_result_table()
        self.widget.sync_analysis_selection_to_owner()

    def tearDown(self):
        self.widget.deleteLater()

    def test_bulk_clear_updates_all_500_rows(self):
        self.widget.set_visible_analysis_selection(False)

        self.assertFalse(any(row["selected"] for row in self.widget.analysis_chart_rows))
        self.assertEqual(get_selected_analysis_rows(self.widget), set())
        self.assertTrue(all(
            self.widget.table.item(index, 0).checkState() == QtCore.Qt.CheckState.Unchecked
            for index in range(500)
        ))

    def test_single_click_refreshes_only_the_matching_result_row(self):
        calls = []
        original = self.widget.populate_analysis_checkbox_for_table_row

        def capture(table_row):
            calls.append(table_row)
            original(table_row)

        self.widget.populate_analysis_checkbox_for_table_row = capture
        self.widget.table.item(250, 0).setCheckState(QtCore.Qt.CheckState.Unchecked)

        self.assertEqual(calls, [250])
        self.assertFalse(self.widget.analysis_chart_rows_by_excel_row[252]["selected"])
        self.assertEqual(len(get_selected_analysis_rows(self.widget)), 499)


if __name__ == "__main__":
    unittest.main()
