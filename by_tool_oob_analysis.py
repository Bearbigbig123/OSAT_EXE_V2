"""Production-only By Tool OOB v2 rules.

This module intentionally contains no replay, simulation, audit, or Old/New
reporting features.  It only prepares each Tool for the approved current rule
and delegates every K-shift decision to the production calculator.
"""

from __future__ import annotations

import json
import colorsys
import copy
import inspect
from typing import Any, Callable, Iterable, Mapping

import numpy as np
import pandas as pd


NEW_WEEKLY_MIN = 5
NEW_BASELINE_MIN = 10
BY_TOOL_K_THRESHOLD = 2.0
WHOLE_BASELINE_K_MAX = 5.0
PERCENTILES = ("P95", "P50", "P05")
PROFILE_VERSION = 1
MIN_BY_TOOL_PLOT_COMPARISON_TOOLS = 2
KShiftCalculator = Callable[..., Any]


CURRENT_BY_TOOL_PERCENTILE_PROFILE = {
    "version": PROFILE_VERSION,
    "characteristics": {
        "Nominal": {
            "P95": {"weekly": 95, "k_base": 95, "own_enabled": True, "own_low": 5, "own_high": 95, "peer_low": 5, "peer_high": 95},
            "P50": {"weekly": 50, "k_base": 50, "own_enabled": True, "own_low": 25, "own_high": 75, "peer_low": 5, "peer_high": 95},
            "P05": {"weekly": 5, "k_base": 5, "own_enabled": True, "own_low": 5, "own_high": 95, "peer_low": 5, "peer_high": 95},
        },
        "Bigger": {
            "P95": {"weekly": 95, "k_base": 95, "own_enabled": True, "own_low": 5, "own_high": 95, "peer_low": 5, "peer_high": 95},
            "P50": {"weekly": 50, "k_base": 50, "own_enabled": True, "own_low": 25, "own_high": None, "peer_low": 5, "peer_high": 95},
            "P05": {"weekly": 5, "k_base": 5, "own_enabled": True, "own_low": 5, "own_high": 95, "peer_low": 5, "peer_high": 95},
        },
        "Smaller": {
            "P95": {"weekly": 95, "k_base": 95, "own_enabled": True, "own_low": 5, "own_high": 95, "peer_low": 5, "peer_high": 95},
            "P50": {"weekly": 50, "k_base": 50, "own_enabled": True, "own_low": None, "own_high": 75, "peer_low": 5, "peer_high": 95},
            "P05": {"weekly": 5, "k_base": 5, "own_enabled": True, "own_low": 5, "own_high": 95, "peer_low": 5, "peer_high": 95},
        },
    },
}


def default_by_tool_percentile_profile() -> dict[str, Any]:
    return copy.deepcopy(CURRENT_BY_TOOL_PERCENTILE_PROFILE)


def _valid_percentile(value: Any, fallback: Any) -> int | None:
    if value is None and fallback is None:
        return None
    try:
        number = int(value)
    except (TypeError, ValueError):
        return fallback
    return number if 0 <= number <= 100 else fallback


def normalize_by_tool_percentile_profile(profile: Any) -> dict[str, Any]:
    """Return a complete v1 profile, replacing invalid/missing fields with current defaults."""
    normalized = default_by_tool_percentile_profile()
    if isinstance(profile, str):
        try:
            profile = json.loads(profile)
        except (TypeError, ValueError, json.JSONDecodeError):
            return normalized
    if not isinstance(profile, Mapping) or profile.get("version") != PROFILE_VERSION:
        return normalized
    source_characteristics = profile.get("characteristics")
    if not isinstance(source_characteristics, Mapping):
        return normalized
    for characteristic in ("Nominal", "Bigger", "Smaller"):
        source_rules = source_characteristics.get(characteristic)
        if not isinstance(source_rules, Mapping):
            continue
        for slot in PERCENTILES:
            source_rule = source_rules.get(slot)
            if not isinstance(source_rule, Mapping):
                continue
            target = normalized["characteristics"][characteristic][slot]
            for key in ("weekly", "k_base", "own_low", "own_high", "peer_low", "peer_high"):
                target[key] = _valid_percentile(source_rule.get(key), target[key])
            target["own_enabled"] = bool(source_rule.get("own_enabled", target["own_enabled"]))
    return normalized


def characteristic_profile(profile: Any, characteristic: Any) -> dict[str, Any]:
    normalized = normalize_by_tool_percentile_profile(profile)
    name = str(characteristic or "Nominal").strip().title()
    if name == "Sigma":
        name = "Smaller"
    if name not in normalized["characteristics"]:
        name = "Nominal"
    return normalized["characteristics"][name]


def percentile_label(value: Any) -> str:
    return f"P{int(value):02d}" if value is not None else "Off"


def _number_label(value: Any) -> str:
    number = float(value)
    return str(int(number)) if number.is_integer() else f"{number:g}"


def by_tool_rule_formula(
    slot: str,
    rule: Mapping[str, Any],
    characteristic: str,
    k_threshold: float = BY_TOOL_K_THRESHOLD,
    peer_bypass_k: float = WHOLE_BASELINE_K_MAX,
    use_peer_gate: bool = True,
) -> str:
    weekly = percentile_label(rule.get("weekly"))
    k_base = percentile_label(rule.get("k_base"))
    if not rule.get("own_enabled", True):
        gate = "Gate Off"
    elif characteristic == "Nominal":
        gate = f"outside Base {percentile_label(rule.get('own_low'))}-{percentile_label(rule.get('own_high'))}"
    elif characteristic == "Bigger":
        gate = f"< Base {percentile_label(rule.get('own_low'))}"
    elif characteristic in {"Smaller", "Sigma"}:
        gate = f"> Base {percentile_label(rule.get('own_high'))}"
    elif slot == "P95":
        gate = f"> Base {percentile_label(rule.get('own_high'))}"
    else:
        gate = f"< Base {percentile_label(rule.get('own_low'))}"
    if use_peer_gate:
        peer_gate = (
            f"outside Whole {percentile_label(rule.get('peer_low'))}-"
            f"{percentile_label(rule.get('peer_high'))}"
        )
        peer_text = f"Peer(K<={_number_label(peer_bypass_k)}):{peer_gate}"
    else:
        peer_text = "Peer:Off"
    return (
        f"W:{weekly} vs K-Base:{k_base}; K>{_number_label(k_threshold)}; Own:{gate}; "
        f"{peer_text}"
    )


def percentile_at(values: Iterable[float], percentile: Any) -> float:
    if percentile is None:
        return np.nan
    array = np.asarray(list(values) if not isinstance(values, np.ndarray) else values, dtype=float)
    array = array[np.isfinite(array)]
    return float(np.percentile(array, int(percentile))) if len(array) else np.nan


def filter_by_tool_plot_data(
    raw_data: pd.DataFrame,
    all_tools_json: str | None,
    matching_column: str = "Matching",
) -> pd.DataFrame:
    """Keep only Tools eligible for By Tool Shift when rule details exist."""
    frame = raw_data.copy()
    if matching_column not in frame.columns:
        frame[matching_column] = "Unknown"
    frame[matching_column] = frame[matching_column].fillna("Unknown").astype(str)

    try:
        records = json.loads(all_tools_json) if all_tools_json else []
    except (TypeError, ValueError, json.JSONDecodeError):
        return frame
    if not isinstance(records, list) or not records:
        return frame

    eligible_tools = {
        str(item.get("tool", ""))
        for item in records
        if isinstance(item, Mapping) and bool(item.get("eligible", False))
    }
    # A comparison view needs at least two eligible Tools. With zero or one,
    # keep the raw By Tool view instead of producing an empty/single-Tool chart.
    if len(eligible_tools) < MIN_BY_TOOL_PLOT_COMPARISON_TOOLS:
        return frame
    return frame[frame[matching_column].isin(eligible_tools)].copy()


def has_eligible_by_tool_shift(all_tools_json: str | None) -> bool:
    """Whether at least two Tools can participate in the plot comparison."""
    try:
        records = json.loads(all_tools_json) if all_tools_json else []
    except (TypeError, ValueError, json.JSONDecodeError):
        return False
    if not isinstance(records, list):
        return False
    eligible_count = sum(
        1 for item in records
        if isinstance(item, Mapping) and bool(item.get("eligible", False))
    )
    return eligible_count >= MIN_BY_TOOL_PLOT_COMPARISON_TOOLS


def _as_float(value: Any) -> float:
    try:
        number = float(value)
    except (TypeError, ValueError):
        return np.nan
    return number if np.isfinite(number) else np.nan


def format_by_tool_shift_label(tool: str, tool_info: Mapping[str, Any]) -> str:
    """Build an unambiguous legend label with a K value for each triggered rule."""
    percentile_rank = {"P05": 0, "P50": 1, "P95": 2}
    triggered = [
        str(rule).replace("_shift", "").upper()
        for rule in tool_info.get("triggered_rules", [])
        if rule
    ]
    triggered.sort(key=lambda name: percentile_rank.get(name, 99))
    if not triggered:
        return str(tool)

    parts = []
    fallback_k = _as_float(tool_info.get("k"))
    formulas = tool_info.get("rule_formulas", {})
    for percentile in triggered:
        k_value = _as_float(tool_info.get(f"{percentile.lower()}_k"))
        if pd.isna(k_value) and len(triggered) == 1:
            k_value = fallback_k
        formula = formulas.get(percentile) if isinstance(formulas, Mapping) else None
        name = f"{percentile}[{formula}]" if formula else percentile
        parts.append(f"{name}={k_value:.2f}σ" if pd.notna(k_value) else name)
    return f"{tool} ({', '.join(parts)})"


def build_by_tool_plot_styles(tools: Iterable[Any]) -> dict[str, dict[str, str]]:
    """Create stable, non-repeating colors plus markers for large Tool sets."""
    marker_cycle = ("o", "s", "^", "D", "v", "P", "X", "<", ">", "h")
    golden_ratio = 0.618033988749895
    styles = {}
    for index, tool in enumerate(tools):
        hue = (0.08 + index * golden_ratio) % 1.0
        red, green, blue = colorsys.hsv_to_rgb(hue, 0.72, 0.82)
        color = "#{:02x}{:02x}{:02x}".format(
            round(red * 255), round(green * 255), round(blue * 255)
        )
        styles[str(tool)] = {
            "color": color,
            "marker": marker_cycle[index % len(marker_cycle)],
        }
    return styles


def by_tool_plot_alpha(tool_info: Mapping[str, Any] | None, has_shift_details: bool) -> float:
    """De-emphasize eligible Tools that did not trigger a By Tool Shift alarm."""
    highlighted = tool_info.get("highlight", False) if tool_info is not None else False
    if isinstance(highlighted, str):
        highlighted = highlighted.strip().lower() in {"true", "1", "yes", "highlight"}
    if bool(highlighted):
        return 0.95
    return 0.45 if has_shift_details else 0.8


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


def whole_baseline_bounds(
    baseline_data: pd.DataFrame,
    min_points: int = 1,
    requested_percentiles: Iterable[int] = (5, 95),
) -> dict[str, float | int]:
    """Calculate Whole Baseline bounds when enough comparison points exist."""
    values = _values(baseline_data)
    if len(values) < min_points:
        percentiles = {str(int(q)): np.nan for q in requested_percentiles if q is not None}
        return {"P95": np.nan, "P05": np.nan, "Count": int(len(values)), "percentiles": percentiles}
    requested = {5, 95, *(int(q) for q in requested_percentiles if q is not None)}
    percentiles = {str(q): percentile_at(values, q) for q in sorted(requested)}
    return {
        "P95": percentiles["95"],
        "P05": percentiles["5"],
        "Count": int(len(values)),
        "percentiles": percentiles,
    }


def _result_dict(result: Any) -> dict[str, Any]:
    if isinstance(result, pd.Series):
        return result.to_dict()
    if isinstance(result, Mapping):
        return dict(result)
    raise TypeError("K-shift calculator must return a mapping or pandas Series")


def _diagnostic_number(value: Any) -> str:
    number = _as_float(value)
    return "N/A" if pd.isna(number) else f"{number:.6g}"


def explain_tool_rule_no_alarm(
    slot: str,
    result: Mapping[str, Any],
    baseline_values: np.ndarray,
    weekly_values: np.ndarray,
    chart_info: Mapping[str, Any],
    rule: Mapping[str, Any],
    k_threshold: float,
) -> str:
    """Return the first production gate that blocked one By Tool rule."""
    weekly_value = percentile_at(weekly_values, rule.get("weekly"))
    base_value = percentile_at(baseline_values, rule.get("k_base"))
    k_value = _as_float(result.get(f"{slot}_k"))
    signed_k = _as_float(result.get(f"{slot}_k_ori"))
    if pd.isna(weekly_value) or pd.isna(base_value) or pd.isna(k_value):
        return "資料不足，無法計算 K"

    difference = abs(weekly_value - base_value)
    resolution = _as_float(chart_info.get("Resolution"))
    if pd.notna(resolution) and resolution > 0 and difference < resolution:
        return (
            f"偏移量 {_diagnostic_number(difference)} 小於 Resolution "
            f"{_diagnostic_number(resolution)}"
        )
    if k_value <= k_threshold:
        return f"K={_diagnostic_number(k_value)} 未超過門檻 {_diagnostic_number(k_threshold)}"

    characteristic = str(chart_info.get("Characteristics", "Nominal") or "Nominal").title()
    if characteristic == "Sigma":
        characteristic = "Smaller"
    if characteristic == "Bigger" and (pd.isna(signed_k) or signed_k >= 0):
        return "偏移方向不符合 Bigger（只檢查向下偏移）"
    if characteristic == "Smaller" and (pd.isna(signed_k) or signed_k <= 0):
        return "偏移方向不符合 Smaller（只檢查向上偏移）"

    if rule.get("own_enabled", True):
        low_q = rule.get("own_low")
        high_q = rule.get("own_high")
        low = percentile_at(baseline_values, low_q)
        high = percentile_at(baseline_values, high_q)
        if characteristic == "Bigger":
            if pd.isna(low):
                return "Tool 自身下限資料不足"
            if weekly_value >= low:
                return (
                    f"Weekly {percentile_label(rule.get('weekly'))}={_diagnostic_number(weekly_value)} "
                    f"未低於 Tool Baseline {percentile_label(low_q)}={_diagnostic_number(low)}"
                )
        elif characteristic == "Smaller":
            if pd.isna(high):
                return "Tool 自身上限資料不足"
            if weekly_value <= high:
                return (
                    f"Weekly {percentile_label(rule.get('weekly'))}={_diagnostic_number(weekly_value)} "
                    f"未高於 Tool Baseline {percentile_label(high_q)}={_diagnostic_number(high)}"
                )
        else:
            if pd.isna(low) or pd.isna(high):
                return "Tool 自身範圍資料不足"
            if low <= weekly_value <= high:
                return (
                    f"Weekly {percentile_label(rule.get('weekly'))}={_diagnostic_number(weekly_value)} "
                    f"仍在 Tool Baseline {percentile_label(low_q)}-{percentile_label(high_q)} "
                    f"({_diagnostic_number(low)}~{_diagnostic_number(high)}) 內"
                )
    return "未通過 Tool 自身規則"


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
    k_threshold = float(chart_info.get("by_tool_median_shift_k_threshold", BY_TOOL_K_THRESHOLD))
    kwargs = {"k_threshold": k_threshold}
    try:
        signature = inspect.signature(kshift_calculator)
        accepts_profile = (
            "rule_profile" in signature.parameters
            or any(p.kind == inspect.Parameter.VAR_KEYWORD for p in signature.parameters.values())
        )
    except (TypeError, ValueError):
        accepts_profile = False
    if accepts_profile:
        kwargs["rule_profile"] = characteristic_profile(
            chart_info.get("by_tool_percentile_profile"),
            chart_info.get("Characteristics"),
        )
    return _result_dict(kshift_calculator(
        base_stats,
        weekly_stats,
        chart_info.get("Characteristics"),
        chart_info.get("Resolution"),
        chart_info.get("UCL"),
        chart_info.get("LCL"),
        **kwargs,
    ))


def apply_whole_baseline_exemption(
    result: Mapping[str, Any],
    weekly_percentiles: Mapping[str, float],
    whole_bounds: Mapping[str, float],
    rule_profile: Mapping[str, Any] | None = None,
    k_threshold: float = BY_TOOL_K_THRESHOLD,
    peer_bypass_k: float = WHOLE_BASELINE_K_MAX,
    use_peer_gate: bool = True,
) -> tuple[dict[str, Any], list[str], list[str]]:
    """Apply the current per-percentile Whole Baseline gate."""
    final = dict(result)
    exemptions: list[str] = []
    kept_over_5: list[str] = []
    whole_p95 = _as_float(whole_bounds.get("P95"))
    whole_p05 = _as_float(whole_bounds.get("P05"))
    peer_percentiles = whole_bounds.get("percentiles", {})
    active_profile = rule_profile or CURRENT_BY_TOOL_PERCENTILE_PROFILE["characteristics"]["Nominal"]

    def peer_value(percentile: Any, legacy_value: float) -> float:
        if percentile is None:
            return np.nan
        if isinstance(peer_percentiles, Mapping):
            value = _as_float(peer_percentiles.get(str(int(percentile))))
            if pd.notna(value):
                return value
        return legacy_value if int(percentile) in {5, 95} else np.nan

    for percentile in PERCENTILES:
        shift_key = f"{percentile}_shift"
        if result.get(shift_key) != "HIGHLIGHT":
            final[shift_key] = "NO_HIGHLIGHT"
            continue
        k_value = abs(_as_float(result.get(f"{percentile}_k")))
        weekly_value = _as_float(weekly_percentiles.get(percentile))
        if pd.isna(k_value) or k_value <= k_threshold:
            final[shift_key] = "NO_HIGHLIGHT"
            continue
        if not use_peer_gate:
            continue
        if k_value > peer_bypass_k:
            kept_over_5.append(shift_key)
            continue

        rule = active_profile[percentile]
        signed_k = _as_float(result.get(f"{percentile}_k_ori"))
        if signed_k > 0:
            boundary_q = rule.get("peer_high")
            boundary_value = peer_value(boundary_q, whole_p95)
            keep = pd.notna(weekly_value) and pd.notna(boundary_value) and weekly_value > boundary_value
            boundary = f"Whole {percentile_label(boundary_q)}={boundary_value:.8g}"
        else:
            boundary_q = rule.get("peer_low")
            boundary_value = peer_value(boundary_q, whole_p05)
            keep = pd.notna(weekly_value) and pd.notna(boundary_value) and weekly_value < boundary_value
            boundary = f"Whole {percentile_label(boundary_q)}={boundary_value:.8g}"
        if not keep:
            final[shift_key] = "NO_HIGHLIGHT"
            if pd.isna(boundary_value):
                exemptions.append(
                    f"{percentile}: {_number_label(k_threshold)}<K<={_number_label(peer_bypass_k)}, "
                    f"peer Whole Baseline has fewer than "
                    f"{NEW_BASELINE_MIN} points"
                )
            else:
                exemptions.append(
                    f"{percentile}: {_number_label(k_threshold)}<K<={_number_label(peer_bypass_k)}, "
                    f"weekly={weekly_value:.8g} did not cross {boundary}"
                )
    return final, exemptions, kept_over_5


def explain_peer_no_alarm(
    slot: str,
    result: Mapping[str, Any],
    weekly_value: float,
    whole_bounds: Mapping[str, Any],
    rule: Mapping[str, Any],
) -> str:
    """Explain a Peer Gate rejection after the Tool's own rule passed."""
    signed_k = _as_float(result.get(f"{slot}_k_ori"))
    upward = signed_k > 0
    boundary_q = rule.get("peer_high" if upward else "peer_low")
    peer_percentiles = whole_bounds.get("percentiles", {})
    boundary_value = np.nan
    if boundary_q is not None and isinstance(peer_percentiles, Mapping):
        boundary_value = _as_float(peer_percentiles.get(str(int(boundary_q))))
    if pd.isna(boundary_value) and boundary_q is not None and int(boundary_q) in {5, 95}:
        boundary_value = _as_float(whole_bounds.get("P95" if upward else "P05"))
    if pd.isna(boundary_value):
        return "By Tool 大盤資料不足，無法通過大盤確認"
    direction = "高於" if upward else "低於"
    return (
        f"Weekly {percentile_label(rule.get('weekly'))}={_diagnostic_number(weekly_value)} "
        f"未{direction}大盤 {percentile_label(boundary_q)}={_diagnostic_number(boundary_value)}"
    )


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

    characteristic_name = str(chart_info.get("Characteristics", "Nominal") or "Nominal").title()
    active_profile = characteristic_profile(
        chart_info.get("by_tool_percentile_profile"), characteristic_name
    )
    k_threshold = float(chart_info.get("by_tool_median_shift_k_threshold", BY_TOOL_K_THRESHOLD))
    peer_bypass_k = float(
        chart_info.get("by_tool_peer_gate_bypass_k_threshold", WHOLE_BASELINE_K_MAX)
    )
    use_peer_gate = bool(chart_info.get("by_tool_use_peer_baseline", True))
    baseline_pct = {
        slot: percentile_at(calculation_baseline, active_profile[slot]["k_base"])
        for slot in PERCENTILES
    }
    weekly_pct = {
        slot: percentile_at(calculation_weekly, active_profile[slot]["weekly"])
        for slot in PERCENTILES
    }
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
        "Whole Baseline Count": int(whole_bounds.get("Count", 0)),
        "Whole Baseline P95": _as_float(whole_bounds.get("P95")),
        "Whole Baseline P05": _as_float(whole_bounds.get("P05")),
        "Exemption Reasons": "",
        "K > 5 Kept Rules": "",
        "Rule Formulas": {
            slot: by_tool_rule_formula(
                slot, active_profile[slot], characteristic_name, k_threshold, peer_bypass_k,
                use_peer_gate,
            )
            for slot in PERCENTILES
        },
        "K Alarm Threshold": k_threshold,
        "K Peer Bypass Threshold": peer_bypass_k,
        "Use Peer Baseline": use_peer_gate,
    }
    for percentile in PERCENTILES:
        row[f"Baseline {percentile}"] = baseline_pct[percentile]
        row[f"Weekly {percentile}"] = weekly_pct[percentile]
        row[f"{percentile} K"] = np.nan
        row[f"{percentile} Signed K"] = np.nan
        row[f"{percentile} HL"] = "NOT_ELIGIBLE" if not eligible else "NO_HIGHLIGHT"
        row[f"{percentile} No Alarm Reason"] = ""

    if not eligible:
        if weekly_count == 0:
            row["Ineligible Reason"] = "Weekly=0"
        else:
            row["Ineligible Reason"] = (
                f"Borrowing {rolling_required} point(s) leaves Baseline "
                f"{remaining_baseline}<{NEW_BASELINE_MIN}"
            )
        for percentile in PERCENTILES:
            row[f"{percentile} No Alarm Reason"] = f"Tool 不符合計算資格：{row['Ineligible Reason']}"
        return row

    rolled_result = call_production_kshift(
        calculation_baseline, calculation_weekly, chart_info, kshift_calculator
    )
    production_reasons = {
        slot: explain_tool_rule_no_alarm(
            slot, rolled_result, calculation_baseline, calculation_weekly,
            chart_info, active_profile[slot], k_threshold,
        )
        for slot in PERCENTILES
        if rolled_result.get(f"{slot}_shift") != "HIGHLIGHT"
    }

    # Weekly 3-4 was already calculable by the former production threshold.
    # Rolling is confirmatory and cannot promote an original NO_HIGHLIGHT.
    current_result = None
    if rolling_required and weekly_count >= 3:
        current_result = call_production_kshift(
            baseline_values, weekly_values, chart_info, kshift_calculator
        )

    final, exemptions, kept_over_5 = apply_whole_baseline_exemption(
        rolled_result, weekly_pct, whole_bounds, active_profile,
        k_threshold=k_threshold, peer_bypass_k=peer_bypass_k,
        use_peer_gate=use_peer_gate,
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
            row[f"{percentile} No Alarm Reason"] = ""
        elif rolled_result.get(shift_key) != "HIGHLIGHT":
            row[f"{percentile} No Alarm Reason"] = production_reasons.get(
                percentile, "未通過 Tool 自身規則"
            )
        elif pd.isna(_as_float(rolled_result.get(f"{percentile}_k"))) or _as_float(
            rolled_result.get(f"{percentile}_k")
        ) <= k_threshold:
            row[f"{percentile} No Alarm Reason"] = (
                f"K={_diagnostic_number(rolled_result.get(f'{percentile}_k'))} "
                f"未超過門檻 {_diagnostic_number(k_threshold)}"
            )
        elif current_result is not None and current_result.get(shift_key) != "HIGHLIGHT":
            row[f"{percentile} No Alarm Reason"] = "Rolling 後通過，但原始 Weekly 3~4 點未通過確認"
        elif use_peer_gate:
            row[f"{percentile} No Alarm Reason"] = explain_peer_no_alarm(
                percentile, rolled_result, weekly_pct[percentile],
                whole_bounds, active_profile[percentile],
            )
        else:
            row[f"{percentile} No Alarm Reason"] = "未通過最終判斷"
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

    baseline_groups = {name: group for name, group in baseline.groupby(matching_column)}
    weekly_groups = {name: group for name, group in weekly.groupby(matching_column)}
    baseline_values_by_tool = {name: _values(group) for name, group in baseline_groups.items()}
    active_profile = characteristic_profile(
        chart_info.get("by_tool_percentile_profile"), chart_info.get("Characteristics")
    )
    requested_peer_percentiles = {
        rule[key]
        for rule in active_profile.values()
        for key in ("peer_low", "peer_high")
        if rule.get(key) is not None
    }
    use_peer_gate = bool(chart_info.get("by_tool_use_peer_baseline", True))
    records = []
    for tool in sorted(set(baseline_groups) | set(weekly_groups)):
        base_values = baseline_values_by_tool.get(tool, np.array([], dtype=float))
        week_values = _values(weekly_groups.get(tool, pd.DataFrame(columns=["point_val"])))
        if use_peer_gate:
            peer_parts = [
                values for peer_tool, values in baseline_values_by_tool.items()
                if peer_tool != tool and len(values)
            ]
            peer_values = np.concatenate(peer_parts) if peer_parts else np.array([], dtype=float)
            peer_bounds = whole_baseline_bounds(
                pd.DataFrame({"point_val": peer_values}),
                min_points=NEW_BASELINE_MIN,
                requested_percentiles=requested_peer_percentiles,
            )
        else:
            peer_bounds = {"P95": np.nan, "P05": np.nan, "Count": 0, "percentiles": {}}
        records.append(
            evaluate_tool_new_rule(
                tool, base_values, week_values, peer_bounds, chart_info, kshift_calculator
            )
        )
    return pd.DataFrame(records)
