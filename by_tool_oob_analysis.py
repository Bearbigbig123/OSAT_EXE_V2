"""Production-only By Tool OOB v2 rules.

This module intentionally contains no replay, simulation, audit, or Old/New
reporting features.  It only prepares each Tool for the approved current rule
and delegates every K-shift decision to the production calculator.
"""

from __future__ import annotations

from typing import Any, Callable, Iterable, Mapping

import numpy as np
import pandas as pd


NEW_WEEKLY_MIN = 5
NEW_BASELINE_MIN = 10
BY_TOOL_K_THRESHOLD = 2.0
WHOLE_BASELINE_K_MAX = 5.0
PERCENTILES = ("P95", "P50", "P05")
KShiftCalculator = Callable[..., Any]


def _as_float(value: Any) -> float:
    try:
        number = float(value)
    except (TypeError, ValueError):
        return np.nan
    return number if np.isfinite(number) else np.nan


def _values(frame: pd.DataFrame) -> np.ndarray:
    if frame is None or frame.empty or "point_val" not in frame.columns:
        return np.array([], dtype=float)
    return pd.to_numeric(frame["point_val"], errors="coerce").dropna().to_numpy(dtype=float)


def percentile_values(values: Iterable[float]) -> dict[str, float]:
    array = np.asarray(list(values) if not isinstance(values, np.ndarray) else values, dtype=float)
    array = array[np.isfinite(array)]
    if not len(array):
        return {name: np.nan for name in PERCENTILES}
    p05, p50, p95 = np.percentile(array, [5, 50, 95])
    return {"P95": float(p95), "P50": float(p50), "P05": float(p05)}


def whole_baseline_bounds(baseline_data: pd.DataFrame) -> dict[str, float]:
    """Calculate Whole Baseline bounds before Matching filtering or borrowing."""
    stats = percentile_values(_values(baseline_data))
    return {"P95": stats["P95"], "P05": stats["P05"]}


def _result_dict(result: Any) -> dict[str, Any]:
    if isinstance(result, pd.Series):
        return result.to_dict()
    if isinstance(result, Mapping):
        return dict(result)
    raise TypeError("K-shift calculator must return a mapping or pandas Series")


def call_production_kshift(
    baseline_values: np.ndarray,
    weekly_values: np.ndarray,
    chart_info: Mapping[str, Any],
    kshift_calculator: KShiftCalculator,
) -> dict[str, Any]:
    baseline_values = np.asarray(baseline_values, dtype=float)
    weekly_values = np.asarray(weekly_values, dtype=float)
    base_stats = {
        "values": baseline_values,
        "cnt": int(len(baseline_values)),
        "mean": float(np.mean(baseline_values)),
        "sigma": float(np.std(baseline_values, ddof=1)) if len(baseline_values) > 1 else 0.0,
    }
    weekly_stats = {
        "values": weekly_values,
        "cnt": int(len(weekly_values)),
        "mean": float(np.mean(weekly_values)),
        "sigma": float(np.std(weekly_values, ddof=1)) if len(weekly_values) > 1 else 0.0,
    }
    return _result_dict(
        kshift_calculator(
            base_stats,
            weekly_stats,
            chart_info.get("Characteristics"),
            chart_info.get("Resolution"),
            chart_info.get("UCL"),
            chart_info.get("LCL"),
            k_threshold=BY_TOOL_K_THRESHOLD,
        )
    )


def apply_whole_baseline_exemption(
    result: Mapping[str, Any],
    weekly_percentiles: Mapping[str, float],
    whole_bounds: Mapping[str, float],
) -> tuple[dict[str, Any], list[str], list[str]]:
    """Apply the current per-percentile Whole Baseline gate."""
    final = dict(result)
    exemptions: list[str] = []
    kept_over_5: list[str] = []
    whole_p95 = _as_float(whole_bounds.get("P95"))
    whole_p05 = _as_float(whole_bounds.get("P05"))

    for percentile in PERCENTILES:
        shift_key = f"{percentile}_shift"
        if result.get(shift_key) != "HIGHLIGHT":
            final[shift_key] = "NO_HIGHLIGHT"
            continue
        k_value = abs(_as_float(result.get(f"{percentile}_k")))
        weekly_value = _as_float(weekly_percentiles.get(percentile))
        if pd.isna(k_value) or k_value <= BY_TOOL_K_THRESHOLD:
            final[shift_key] = "NO_HIGHLIGHT"
            continue
        if k_value > WHOLE_BASELINE_K_MAX:
            kept_over_5.append(shift_key)
            continue

        if percentile == "P95":
            keep = pd.notna(weekly_value) and pd.notna(whole_p95) and weekly_value > whole_p95
            boundary = f"Whole P95={whole_p95:.8g}"
        elif percentile == "P05":
            keep = pd.notna(weekly_value) and pd.notna(whole_p05) and weekly_value < whole_p05
            boundary = f"Whole P05={whole_p05:.8g}"
        else:
            signed_k = _as_float(result.get("P50_k_ori"))
            if signed_k > 0:
                keep = pd.notna(weekly_value) and pd.notna(whole_p95) and weekly_value > whole_p95
                boundary = f"Whole P95={whole_p95:.8g}"
            else:
                keep = pd.notna(weekly_value) and pd.notna(whole_p05) and weekly_value < whole_p05
                boundary = f"Whole P05={whole_p05:.8g}"
        if not keep:
            final[shift_key] = "NO_HIGHLIGHT"
            exemptions.append(
                f"{percentile}: 2<K<=5, weekly={weekly_value:.8g} did not cross {boundary}"
            )
    return final, exemptions, kept_over_5


def evaluate_tool_new_rule(
    tool: str,
    baseline_values: np.ndarray,
    weekly_values: np.ndarray,
    whole_bounds: Mapping[str, float],
    chart_info: Mapping[str, Any],
    kshift_calculator: KShiftCalculator,
) -> dict[str, Any]:
    """Evaluate one Tool using only the current production rule."""
    baseline_values = np.asarray(baseline_values, dtype=float)
    weekly_values = np.asarray(weekly_values, dtype=float)
    baseline_values = baseline_values[np.isfinite(baseline_values)]
    weekly_values = weekly_values[np.isfinite(weekly_values)]
    baseline_count, weekly_count = len(baseline_values), len(weekly_values)

    rolling_required = max(0, NEW_WEEKLY_MIN - weekly_count) if weekly_count > 0 else 0
    remaining_baseline = baseline_count - rolling_required
    eligible = weekly_count > 0 and remaining_baseline >= NEW_BASELINE_MIN

    if eligible and rolling_required:
        borrowed = baseline_values[-rolling_required:]
        calculation_baseline = baseline_values[:-rolling_required]
        calculation_weekly = np.concatenate([borrowed, weekly_values])
    else:
        calculation_baseline = baseline_values
        calculation_weekly = weekly_values

    baseline_pct = percentile_values(calculation_baseline)
    weekly_pct = percentile_values(calculation_weekly)
    row: dict[str, Any] = {
        "Tool": str(tool),
        "Characteristic": chart_info.get("Characteristics", "N/A"),
        "Weekly Count": int(weekly_count),
        "Baseline Count": int(baseline_count),
        "Eligible": bool(eligible),
        "Highlight": "NOT_ELIGIBLE" if not eligible else "NO_HIGHLIGHT",
        "Rolling Used": bool(eligible and rolling_required > 0),
        "Rolling Points Borrowed": int(rolling_required if eligible else 0),
        "Calculation Weekly Count": int(len(calculation_weekly)),
        "Calculation Baseline Count": int(len(calculation_baseline)),
        "Whole Baseline P95": _as_float(whole_bounds.get("P95")),
        "Whole Baseline P05": _as_float(whole_bounds.get("P05")),
        "Exemption Reasons": "",
        "K > 5 Kept Rules": "",
    }
    for percentile in PERCENTILES:
        row[f"Baseline {percentile}"] = baseline_pct[percentile]
        row[f"Weekly {percentile}"] = weekly_pct[percentile]
        row[f"{percentile} K"] = np.nan
        row[f"{percentile} Signed K"] = np.nan
        row[f"{percentile} HL"] = "NOT_ELIGIBLE" if not eligible else "NO_HIGHLIGHT"

    if not eligible:
        if weekly_count == 0:
            row["Ineligible Reason"] = "Weekly=0"
        else:
            row["Ineligible Reason"] = (
                f"Borrowing {rolling_required} point(s) leaves Baseline "
                f"{remaining_baseline}<{NEW_BASELINE_MIN}"
            )
        return row

    rolled_result = call_production_kshift(
        calculation_baseline, calculation_weekly, chart_info, kshift_calculator
    )

    # Weekly 3-4 was already calculable by the former production threshold.
    # Rolling is confirmatory and cannot promote an original NO_HIGHLIGHT.
    current_result = None
    if rolling_required and weekly_count >= 3:
        current_result = call_production_kshift(
            baseline_values, weekly_values, chart_info, kshift_calculator
        )

    final, exemptions, kept_over_5 = apply_whole_baseline_exemption(
        rolled_result, weekly_pct, whole_bounds
    )
    triggered = []
    for percentile in PERCENTILES:
        shift_key = f"{percentile}_shift"
        if current_result is not None and current_result.get(shift_key) != "HIGHLIGHT":
            final[shift_key] = "NO_HIGHLIGHT"
        row[f"{percentile} K"] = _as_float(rolled_result.get(f"{percentile}_k"))
        row[f"{percentile} Signed K"] = _as_float(rolled_result.get(f"{percentile}_k_ori"))
        row[f"{percentile} HL"] = final.get(shift_key, "NO_HIGHLIGHT")
        if row[f"{percentile} HL"] == "HIGHLIGHT":
            triggered.append(shift_key)
    row["Highlight"] = "HIGHLIGHT" if triggered else "NO_HIGHLIGHT"
    row["Triggered Rules"] = ", ".join(triggered)
    row["Exemption Reasons"] = " | ".join(exemptions)
    row["K > 5 Kept Rules"] = ", ".join(kept_over_5)
    return row


def evaluate_by_tool_new_rules(
    baseline_data: pd.DataFrame,
    weekly_data: pd.DataFrame,
    chart_info: Mapping[str, Any],
    kshift_calculator: KShiftCalculator,
    matching_column: str = "Matching",
) -> pd.DataFrame:
    """Evaluate all Tools using only the current production rule."""
    required = {matching_column, "point_val"}
    if baseline_data is None or weekly_data is None:
        return pd.DataFrame()
    if not required.issubset(baseline_data.columns) or not required.issubset(weekly_data.columns):
        return pd.DataFrame()

    baseline = baseline_data.copy()
    weekly = weekly_data.copy()
    for frame in (baseline, weekly):
        frame["point_val"] = pd.to_numeric(frame["point_val"], errors="coerce")
        frame[matching_column] = frame[matching_column].fillna("Unknown").astype(str)
        frame.dropna(subset=["point_val"], inplace=True)
        if "point_time" in frame.columns:
            frame["point_time"] = pd.to_datetime(frame["point_time"], errors="coerce")
            frame.sort_values("point_time", inplace=True)

    bounds = whole_baseline_bounds(baseline)
    baseline_groups = {name: group for name, group in baseline.groupby(matching_column)}
    weekly_groups = {name: group for name, group in weekly.groupby(matching_column)}
    records = []
    for tool in sorted(set(baseline_groups) | set(weekly_groups)):
        base_values = _values(baseline_groups.get(tool, pd.DataFrame(columns=["point_val"])))
        week_values = _values(weekly_groups.get(tool, pd.DataFrame(columns=["point_val"])))
        records.append(
            evaluate_tool_new_rule(
                tool, base_values, week_values, bounds, chart_info, kshift_calculator
            )
        )
    return pd.DataFrame(records)
