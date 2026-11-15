"""Generate broad current-rule By Tool fixtures under input/raw_charts."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
import shutil

import numpy as np
import pandas as pd


ROOT = Path(__file__).resolve().parent
INPUT = ROOT / "input"
RAW = INPUT / "raw_charts"
BOOK = INPUT / "All_Chart_Information.xlsx"
BACKUPS = INPUT / "backups"
GROUP = "BYTOOL_V2_EXTENDED"
TARGET = 100.0
DEFAULT_OFFSETS = [-2, -1, -1, 0, 0, 0, 1, 1, 2, 2]
CHARACTERISTICS = ("Nominal", "Bigger", "Smaller")


def clean_token(value):
    return str(value).replace("-", "m").replace("+", "p").replace(".", "p")


def draw(rng, center, sigma, n, distribution="normal"):
    if distribution == "normal":
        values = rng.normal(center, sigma, n)
    elif distribution == "right_skew":
        values = center + (rng.lognormal(0, 0.55, n) - np.exp(0.55 ** 2 / 2)) * sigma
    elif distribution == "left_skew":
        values = center - (rng.lognormal(0, 0.55, n) - np.exp(0.55 ** 2 / 2)) * sigma
    elif distribution == "heavy_tail":
        values = center + rng.standard_t(3, n) * sigma / np.sqrt(3)
    elif distribution == "bimodal":
        side = rng.choice([-1.0, 1.0], n)
        values = center + side * sigma + rng.normal(0, sigma * 0.35, n)
    elif distribution == "discrete":
        values = center + rng.choice([-2, -1, 0, 1, 2], n) * sigma
    elif distribution == "outlier":
        values = rng.normal(center, sigma, n)
        if n:
            values[-1] = center + 6 * sigma
    else:
        raise ValueError(distribution)
    return values


def spec(name, category, characteristic="Nominal", tool_count=10, baseline_n=30,
         weekly_n=5, target_tool=0, target_shift=0.0, shifts=None,
         offsets=None, baseline_sigmas=None, weekly_sigmas=None,
         distribution="normal", resolution=0.01, cl_low=94.0, cl_high=106.0,
         missing_weekly_tools=None, new_weekly_tool=False, blank_matching=False):
    return locals()


def build_specs():
    specs = []

    # Weekly sample sizes and the exact remaining-Baseline eligibility boundary.
    for characteristic in CHARACTERISTICS:
        for weekly_n in (1, 2, 3, 4, 5, 8, 10, 20):
            borrow = max(0, 5 - weekly_n)
            for remaining in (9, 10, 15, 30):
                baseline_n = remaining + borrow
                shift = -3 if characteristic == "Bigger" else 3
                specs.append(spec(
                    f"SAMPLE_{characteristic}_W{weekly_n:02d}_REM{remaining:02d}",
                    "sample_count", characteristic, baseline_n=baseline_n,
                    weekly_n=weekly_n, target_tool=(weekly_n + remaining) % 10,
                    target_shift=shift,
                ))

    # Fully cross Characteristic, Tool natural level, direction, and magnitude.
    representative_tools = {"m2": 0, "m1": 1, "z0": 3, "p1": 6, "p2": 8}
    for characteristic in CHARACTERISTICS:
        for offset_name, tool_index in representative_tools.items():
            for shift in (-6, -4, -2, -1, 0, 1, 2, 4, 6):
                weekly_n = (5, 10, 20)[(tool_index + int(abs(shift))) % 3]
                specs.append(spec(
                    f"CROSS_{characteristic}_{offset_name}_S{clean_token(shift)}_W{weekly_n}",
                    "characteristic_level_shift", characteristic,
                    weekly_n=weekly_n, target_tool=tool_index, target_shift=shift,
                ))

    # Fractional shifts around likely K and Whole-Baseline transition regions.
    for characteristic in CHARACTERISTICS:
        for index, shift in enumerate((-5.5, -5.0, -2.5, -2.0, -1.5, -0.5,
                                       0.5, 1.5, 2.0, 2.5, 5.0, 5.5)):
            specs.append(spec(
                f"FRACTION_{characteristic}_S{clean_token(shift)}",
                "fractional_shift", characteristic, weekly_n=10,
                target_tool=index % 10, target_shift=shift,
            ))

    # Multiple Tool movements in the same week.
    patterns = {
        "TWO_UP": {0: 3, 9: 3},
        "TWO_DOWN": {0: -3, 9: -3},
        "OPPOSITE": {0: 4, 9: -4},
        "SPLIT": {0: -3, 1: -2, 8: 2, 9: 3},
        "ALL_UP": {i: 2 for i in range(10)},
        "ALL_DOWN": {i: -2 for i in range(10)},
        "CENTER_ONLY": {3: 3, 4: 3, 5: 3},
        "EDGES_IN": {0: 2, 9: -2},
    }
    for characteristic in CHARACTERISTICS:
        for pattern_name, shifts in patterns.items():
            specs.append(spec(
                f"MULTI_{characteristic}_{pattern_name}", "multi_tool_shift",
                characteristic, weekly_n=10, shifts=shifts,
            ))

    # Different baseline/weekly variation and non-normal distributions.
    for characteristic in CHARACTERISTICS:
        for baseline_sigma, weekly_sigma in ((0.3, 1), (0.5, 2), (1, 0.25),
                                               (1, 2), (1, 3), (2, 1), (3, 3)):
            specs.append(spec(
                f"SIGMA_{characteristic}_B{clean_token(baseline_sigma)}_W{clean_token(weekly_sigma)}",
                "sigma_change", characteristic, weekly_n=10, target_tool=5,
                target_shift=(-3 if characteristic == "Bigger" else 3),
                baseline_sigmas=[baseline_sigma] * 10,
                weekly_sigmas=[weekly_sigma] * 10,
            ))
        for distribution in ("right_skew", "left_skew", "heavy_tail", "bimodal", "discrete", "outlier"):
            specs.append(spec(
                f"DIST_{characteristic}_{distribution}", "distribution",
                characteristic, weekly_n=10, target_tool=5,
                target_shift=(-3 if characteristic == "Bigger" else 3),
                distribution=distribution,
            ))

    # Resolution and control-limit sensitivity.
    for characteristic in CHARACTERISTICS:
        for resolution in (None, 0, 0.01, 0.1, 0.5, 1.0):
            specs.append(spec(
                f"RES_{characteristic}_{clean_token(resolution)}", "resolution",
                characteristic, weekly_n=10, target_tool=4, target_shift=1.5,
                resolution=resolution,
            ))
        for width in (3, 4, 6, 8, 12):
            specs.append(spec(
                f"CL_{characteristic}_PM{width}", "control_limit",
                characteristic, weekly_n=10, target_tool=4,
                target_shift=(-3 if characteristic == "Bigger" else 3),
                cl_low=TARGET - width, cl_high=TARGET + width,
            ))
        specs.append(spec(
            f"CL_{characteristic}_ASYM_UP", "control_limit", characteristic,
            weekly_n=10, target_tool=4, target_shift=3, cl_low=92, cl_high=104,
        ))
        specs.append(spec(
            f"CL_{characteristic}_ASYM_DOWN", "control_limit", characteristic,
            weekly_n=10, target_tool=4, target_shift=-3, cl_low=96, cl_high=108,
        ))

    # Tool population and Matching availability.
    for tool_count in (2, 5, 20, 30):
        offsets = np.linspace(-2, 2, tool_count).round(2).tolist()
        specs.append(spec(
            f"TOOLS_{tool_count:02d}", "tool_population", "Nominal",
            tool_count=tool_count, offsets=offsets, weekly_n=10,
            target_tool=tool_count - 1, target_shift=3,
        ))
    specs.extend([
        spec("MATCH_MISSING_WEEKLY", "matching_edge", missing_weekly_tools={8, 9}),
        spec("MATCH_NEW_WEEKLY_TOOL", "matching_edge", new_weekly_tool=True),
        spec("MATCH_BLANK", "matching_edge", blank_matching=True),
    ])
    return specs


def generate_chart(item, chart_index, rng):
    tool_count = item["tool_count"]
    offsets = item["offsets"] or (DEFAULT_OFFSETS if tool_count == 10 else np.linspace(-2, 2, tool_count).tolist())
    baseline_sigmas = item["baseline_sigmas"] or [1.0] * tool_count
    weekly_sigmas = item["weekly_sigmas"] or [1.0] * tool_count
    shifts = dict(item["shifts"] or {})
    shifts.setdefault(item["target_tool"], item["target_shift"])
    baseline_n = item["baseline_n"]
    weekly_n = item["weekly_n"]
    chart_name = f"EXT_{chart_index:03d}_{item['name']}"
    chart_id = f"BTV2E_{chart_index:03d}"
    rows = []
    baseline_dates = pd.date_range("2026-01-05 08:00", "2026-10-30 16:00", periods=baseline_n)
    # Keep every configured Weekly sample size inside one Monday-Sunday window.
    weekly_dates = pd.date_range(
        "2026-11-02 08:00", "2026-11-08 20:00", periods=weekly_n
    )
    for index in range(tool_count):
        tool = f"Tool{index + 1:02d}"
        matching = "" if item["blank_matching"] and index == tool_count - 1 else tool
        center = TARGET + float(offsets[index])
        base_values = draw(rng, center, float(baseline_sigmas[index]), baseline_n, item["distribution"])
        week_values = draw(
            rng, center + float(shifts.get(index, 0)), float(weekly_sigmas[index]),
            weekly_n, item["distribution"],
        )
        for number, (time, value) in enumerate(zip(baseline_dates, base_values), 1):
            rows.append({"GroupName": GROUP, "ChartName": chart_name,
                         "point_time": time.strftime("%Y/%m/%d %H:%M"),
                         "point_val": round(float(value), 2),
                         "Batch_ID": f"BSL_{tool}_{number:03d}", "Matching": matching})
        if index not in (item["missing_weekly_tools"] or set()):
            for number, (time, value) in enumerate(zip(weekly_dates, week_values), 1):
                rows.append({"GroupName": GROUP, "ChartName": chart_name,
                             "point_time": time.strftime("%Y/%m/%d %H:%M"),
                             "point_val": round(float(value), 2),
                             "Batch_ID": f"WK_{tool}_{number:03d}", "Matching": matching})
    if item["new_weekly_tool"]:
        for number, time in enumerate(weekly_dates, 1):
            rows.append({"GroupName": GROUP, "ChartName": chart_name,
                         "point_time": time.strftime("%Y/%m/%d %H:%M"),
                         "point_val": round(float(rng.normal(TARGET, 1)), 2),
                         "Batch_ID": f"WK_NewTool_{number:03d}", "Matching": "NewTool"})

    frame = pd.DataFrame(rows)
    path = RAW / f"{GROUP}_{chart_name}.csv"
    frame.to_csv(path, index=False, encoding="utf-8-sig")
    characteristic = item["characteristic"]
    metadata = {
        "GroupName": GROUP, "ChartName": chart_name, "ChartID": chart_id,
        "Material_no": f"MAT_{chart_id}", "Target": TARGET,
        "UCL": item["cl_high"], "LCL": item["cl_low"],
        "USL": 115.0, "LSL": 85.0, "Characteristics": characteristic,
        "DetectionLimit": np.nan,
        "ExpectedPattern": (
            f"category={item['category']}; weekly_n={weekly_n}; baseline_n={baseline_n}; "
            f"target_shift={item['target_shift']}; distribution={item['distribution']}"
        ),
        "SampleCount": len(frame), "Resolution": item["resolution"],
    }
    manifest = {
        **metadata, "ScenarioCategory": item["category"],
        "ConfiguredWeeklyN": weekly_n, "ConfiguredBaselineN": baseline_n,
        "ToolCount": tool_count, "TargetTool": item["target_tool"] + 1,
        "TargetShiftSigma": item["target_shift"], "Distribution": item["distribution"],
    }
    return metadata, manifest


def main():
    RAW.mkdir(parents=True, exist_ok=True)
    BACKUPS.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = BACKUPS / f"All_Chart_Information_before_extended_{timestamp}.xlsx"
    shutil.copy2(BOOK, backup)
    existing = pd.read_excel(BOOK, sheet_name="Chart")
    existing = existing[existing["GroupName"].astype(str) != GROUP].copy()

    specs = build_specs()
    seeds = iter(np.random.SeedSequence(20261116).spawn(len(specs)))
    metadata_rows, manifest_rows = [], []
    for index, item in enumerate(specs, 1):
        metadata, manifest = generate_chart(item, index, np.random.default_rng(next(seeds)))
        metadata_rows.append(metadata)
        manifest_rows.append(manifest)

    generated = pd.DataFrame(metadata_rows)
    pd.concat([existing, generated], ignore_index=True, sort=False).to_excel(
        BOOK, sheet_name="Chart", index=False
    )
    manifest_path = INPUT / "BYTOOL_V2_EXTENDED_MANIFEST.csv"
    pd.DataFrame(manifest_rows).to_csv(manifest_path, index=False, encoding="utf-8-sig")
    print(f"Generated extended charts: {len(generated)}")
    print(pd.Series([item['category'] for item in specs]).value_counts().sort_index().to_string())
    print(f"Metadata: {BOOK}")
    print(f"Manifest: {manifest_path}")
    print(f"Backup: {backup}")


if __name__ == "__main__":
    main()
