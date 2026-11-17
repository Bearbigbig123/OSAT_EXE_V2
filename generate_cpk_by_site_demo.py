"""Generate ten reproducible Cpk By Site demo charts covering about three months."""

from pathlib import Path

import numpy as np
import pandas as pd


OUTPUT_DIR = Path(__file__).resolve().parent / "test_data" / "cpk_by_site_demo_10"
RAW_DIR = OUTPUT_DIR / "raw_charts"
GROUP_NAME = "CPK_SITE_DEMO"
DATES = pd.date_range("2026-08-01 08:00", "2026-10-31 08:00", freq="D")


CASES = [
    {
        "chart_name": "01_Stable_4Sites",
        "chart_id": "CPKSITE001",
        "target": 100.0,
        "ucl": 103.0,
        "lcl": 97.0,
        "usl": 106.0,
        "lsl": 94.0,
        "characteristics": "Nominal",
        "site_count": 4,
        "sigma": 0.65,
        "pattern": "Stable process; four well-aligned sites",
    },
    {
        "chart_name": "02_OneSiteHighOffset",
        "chart_id": "CPKSITE002",
        "target": 100.0,
        "ucl": 103.0,
        "lcl": 97.0,
        "usl": 106.0,
        "lsl": 94.0,
        "characteristics": "Nominal",
        "site_count": 5,
        "sigma": 0.55,
        "site_offsets": [0.0, 0.1, -0.1, 0.15, 2.0],
        "pattern": "site5 has a persistent positive offset",
    },
    {
        "chart_name": "03_UpwardDrift_6Sites",
        "chart_id": "CPKSITE003",
        "target": 100.0,
        "ucl": 103.0,
        "lcl": 97.0,
        "usl": 106.0,
        "lsl": 94.0,
        "characteristics": "Nominal",
        "site_count": 6,
        "sigma": 0.55,
        "drift": 3.2,
        "pattern": "All six sites drift upward across the three-month period",
    },
    {
        "chart_name": "04_VariationIncrease",
        "chart_id": "CPKSITE004",
        "target": 100.0,
        "ucl": 103.5,
        "lcl": 96.5,
        "usl": 107.0,
        "lsl": 93.0,
        "characteristics": "Nominal",
        "site_count": 4,
        "sigma": 0.35,
        "end_sigma": 1.8,
        "pattern": "Within-site variation increases gradually over time",
    },
    {
        "chart_name": "05_IntermittentOutliers",
        "chart_id": "CPKSITE005",
        "target": 50.0,
        "ucl": 52.0,
        "lcl": 48.0,
        "usl": 54.0,
        "lsl": 46.0,
        "characteristics": "Nominal",
        "site_count": 8,
        "sigma": 0.45,
        "outliers": True,
        "pattern": "Mostly stable with intermittent site-level out-of-spec points",
    },
    {
        "chart_name": "06_SmallerNearUSL",
        "chart_id": "CPKSITE006",
        "target": 4.0,
        "ucl": 7.0,
        "lcl": 0.0,
        "usl": 8.0,
        "lsl": np.nan,
        "characteristics": "Smaller",
        "site_count": 3,
        "sigma": 0.55,
        "base_shift": 2.2,
        "pattern": "Smaller-is-better process operating closer to the USL",
    },
    {
        "chart_name": "07_BiggerOneSiteLow",
        "chart_id": "CPKSITE007",
        "target": 80.0,
        "ucl": 90.0,
        "lcl": 72.0,
        "usl": np.nan,
        "lsl": 70.0,
        "characteristics": "Bigger",
        "site_count": 5,
        "sigma": 1.0,
        "site_offsets": [0.0, 0.5, -0.3, 0.2, -6.5],
        "pattern": "Bigger-is-better process with site5 persistently low",
    },
    {
        "chart_name": "08_ManySites_12",
        "chart_id": "CPKSITE008",
        "target": 25.0,
        "ucl": 26.5,
        "lcl": 23.5,
        "usl": 28.0,
        "lsl": 22.0,
        "characteristics": "Nominal",
        "site_count": 12,
        "sigma": 0.35,
        "site_offset_span": 1.2,
        "pattern": "Twelve sites with small evenly distributed offsets",
    },
    {
        "chart_name": "09_MissingSiteReadings",
        "chart_id": "CPKSITE009",
        "target": 200.0,
        "ucl": 204.0,
        "lcl": 196.0,
        "usl": 208.0,
        "lsl": 192.0,
        "characteristics": "Nominal",
        "site_count": 6,
        "sigma": 0.9,
        "missing_rate": 0.18,
        "pattern": "Six sites with random missing readings; valid cells are pooled",
    },
    {
        "chart_name": "10_FinalMonthStepChange",
        "chart_id": "CPKSITE010",
        "target": 10.0,
        "ucl": 10.9,
        "lcl": 9.1,
        "usl": 11.8,
        "lsl": 8.2,
        "characteristics": "Nominal",
        "site_count": 4,
        "sigma": 0.22,
        "step_change": 0.95,
        "pattern": "All sites step upward at the start of the final month",
    },
]


def _calculate_expected_cpk(values, case):
    values = np.asarray(values, dtype=float)
    values = values[np.isfinite(values)]
    mean = values.mean()
    sigma = values.std(ddof=1)
    characteristic = case["characteristics"]
    if characteristic == "Nominal":
        cpk = min((case["usl"] - mean) / (3 * sigma), (mean - case["lsl"]) / (3 * sigma))
    elif characteristic == "Smaller":
        cpk = (case["usl"] - mean) / (3 * sigma)
    else:
        cpk = (mean - case["lsl"]) / (3 * sigma)
    return mean, sigma, cpk


def _build_case(case, case_index):
    rng = np.random.default_rng(20261118 + case_index)
    row_count = len(DATES)
    site_count = case["site_count"]
    progress = np.linspace(0.0, 1.0, row_count)

    if "site_offsets" in case:
        site_offsets = np.asarray(case["site_offsets"], dtype=float)
    else:
        span = case.get("site_offset_span", 0.35)
        site_offsets = np.linspace(-span / 2, span / 2, site_count)

    center = np.full(row_count, case["target"] + case.get("base_shift", 0.0), dtype=float)
    center += case.get("drift", 0.0) * progress
    if case.get("step_change"):
        center[DATES >= pd.Timestamp("2026-10-01 08:00")] += case["step_change"]

    start_sigma = case["sigma"]
    sigma_by_day = start_sigma + (case.get("end_sigma", start_sigma) - start_sigma) * progress
    site_values = np.empty((row_count, site_count), dtype=float)
    for site_index in range(site_count):
        site_values[:, site_index] = (
            center
            + site_offsets[site_index]
            + rng.normal(0.0, sigma_by_day, row_count)
        )

    if case.get("outliers"):
        outlier_cells = [(12, 1, 5.2), (29, 6, -4.8), (47, 3, 5.6), (71, 0, -5.1), (84, 5, 5.0)]
        for row_index, site_index, delta in outlier_cells:
            site_values[row_index, site_index] += delta

    if case.get("missing_rate"):
        missing_mask = rng.random(site_values.shape) < case["missing_rate"]
        # Keep at least one measurement per timestamp.
        missing_mask[:, 0] = False
        site_values[missing_mask] = np.nan

    site_columns = [f"site{index}" for index in range(1, site_count + 1)]
    raw_df = pd.DataFrame(site_values, columns=site_columns)
    raw_df.insert(0, "Matching", "SitePool")
    raw_df.insert(0, "Batch_ID", [f"{case['chart_id']}_{index:03d}" for index in range(1, row_count + 1)])
    raw_df.insert(0, "point_val", np.nanmean(site_values, axis=1))
    raw_df.insert(0, "point_time", DATES.strftime("%Y/%m/%d %H:%M"))
    raw_df.insert(0, "ChartName", case["chart_name"])
    raw_df.insert(0, "GroupName", GROUP_NAME)

    numeric_columns = ["point_val", *site_columns]
    raw_df[numeric_columns] = raw_df[numeric_columns].round(3)
    saved_site_values = raw_df[site_columns].to_numpy(dtype=float)
    mean, sigma, cpk = _calculate_expected_cpk(saved_site_values.ravel(), case)
    return raw_df, mean, sigma, cpk, int(np.isfinite(saved_site_values).sum())


def generate():
    RAW_DIR.mkdir(parents=True, exist_ok=True)
    chart_rows = []
    summary_rows = []

    for case_index, case in enumerate(CASES, start=1):
        raw_df, pooled_mean, pooled_sigma, expected_cpk, valid_site_values = _build_case(case, case_index)
        raw_path = RAW_DIR / f"{GROUP_NAME}_{case['chart_name']}.csv"
        raw_df.to_csv(raw_path, index=False, encoding="utf-8-sig")

        chart_rows.append({
            "GroupName": GROUP_NAME,
            "ChartName": case["chart_name"],
            "ChartID": case["chart_id"],
            "Material_no": f"MAT_{case['chart_id']}",
            "Target": case["target"],
            "UCL": case["ucl"],
            "LCL": case["lcl"],
            "USL": case["usl"],
            "LSL": case["lsl"],
            "Characteristics": case["characteristics"],
            "DetectionLimit": np.nan,
            "ExpectedPattern": case["pattern"],
            "SampleCount": valid_site_values,
            "Resolution": 0.001,
        })
        summary_rows.append({
            "ChartName": case["chart_name"],
            "DateStart": DATES.min(),
            "DateEnd": DATES.max(),
            "RawRows": len(DATES),
            "SiteCount": case["site_count"],
            "ValidSiteValues": valid_site_values,
            "ExpectedPooledMean": round(pooled_mean, 6),
            "ExpectedPooledSigma": round(pooled_sigma, 6),
            "ExpectedLongTermCpk": round(expected_cpk, 3),
            "ExpectedPattern": case["pattern"],
        })

    chart_df = pd.DataFrame(chart_rows)
    with pd.ExcelWriter(OUTPUT_DIR / "All_Chart_Information.xlsx", engine="openpyxl") as writer:
        chart_df.to_excel(writer, sheet_name="Chart", index=False)
    pd.DataFrame(summary_rows).to_csv(OUTPUT_DIR / "expected_summary.csv", index=False, encoding="utf-8-sig")

    readme = (
        "Cpk By Site demo dataset\n"
        "========================\n"
        "Charts: 10\n"
        "Date range: 2026-08-01 through 2026-10-31 (92 daily rows per chart)\n"
        "Input mode: choose 'By Site' on the Cpk page\n"
        "Site columns: site1...siteN (no underscore)\n"
        "point_val: mean of all available site readings at that timestamp\n"
        "Cpk By Site: all non-empty site cells in the selected period are pooled\n"
    )
    (OUTPUT_DIR / "README.txt").write_text(readme, encoding="utf-8")
    return chart_df, pd.DataFrame(summary_rows)


if __name__ == "__main__":
    charts, summary = generate()
    print(f"Generated {len(charts)} charts in {OUTPUT_DIR}")
    print(summary[["ChartName", "SiteCount", "ValidSiteValues", "ExpectedLongTermCpk"]].to_string(index=False))
