import pandas as pd
from PyQt6 import QtWidgets

from translations import tr


SELECTION_ATTR = "selected_analysis_chart_rows"


def find_selection_owner(widget):
    current = widget
    while current is not None:
        if hasattr(current, SELECTION_ATTR):
            return current
        current = current.parent() if hasattr(current, "parent") else None
    return None


def get_selected_analysis_rows(widget):
    owner = find_selection_owner(widget)
    if owner is None:
        return None
    return getattr(owner, SELECTION_ATTR, None)


def set_selected_analysis_rows(widget, selected_rows):
    owner = find_selection_owner(widget)
    if owner is None:
        owner = widget.window() if hasattr(widget, "window") else widget
    setattr(owner, SELECTION_ATTR, set(selected_rows))
    return owner


def clear_selected_analysis_rows(widget):
    owner = find_selection_owner(widget)
    if owner is None:
        owner = widget.window() if hasattr(widget, "window") else widget
    setattr(owner, SELECTION_ATTR, None)
    return owner


def filter_chart_info_for_analysis(widget, chart_info_df):
    selected_rows = get_selected_analysis_rows(widget)
    if selected_rows is None:
        return chart_info_df

    selected_rows = set(selected_rows)
    if not selected_rows:
        return chart_info_df.iloc[0:0].copy()

    excel_rows = pd.Series(chart_info_df.index, index=chart_info_df.index).astype(int) + 2
    return chart_info_df[excel_rows.isin(selected_rows)].copy()


def warn_if_no_analysis_selection(widget):
    selected_rows = get_selected_analysis_rows(widget)
    if selected_rows is not None and not selected_rows:
        QtWidgets.QMessageBox.warning(
            widget,
            tr("warning", "Warning"),
            tr("select_at_least_one_analysis_chart", "Please select at least one chart for analysis."),
        )
        return True
    return False
