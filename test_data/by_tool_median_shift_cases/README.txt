OOB By-Tool Median Shift test package - many-tool, non-extreme version

Files:
- All_Chart_Information.xlsx: chart metadata
- raw_charts/*.csv: one raw data file per chart
- OOB_ByTool_MedianShift_TestWorkbook.xlsx: combined workbook with Chart and Time sheets

Expected quick checks:
- A_ByToolShift_ManyTools: many tools; Tool12-14 have a mild mean/median separation, no single huge outlier.
- B_BSLWideSigmaOOCDisplay: baseline/BSL can show red OOC points due to wider sigma, but weekly ooc_cnt should remain 0.
- C_RecordHighLow_SmallEdge: record high/low markers are only slightly beyond baseline extremes.
- D_ManyToolsInsufficientPoints: many tools but fewer than 3 weekly points/tool, so by-tool median shift should be NO_HIGHLIGHT.
- E_WeeklyOOC_ByToolWideSigma: shifted wider-sigma tools create weekly red OOC points and by-tool median shift.
