# -*- coding: utf-8 -*-
"""
多語言翻譯管理系統
支援中文繁體 (ZH_TW) 和英文 (EN)
"""

class TranslationManager:
    """全域翻譯管理器 - 單例模式"""
    _instance = None
    _current_lang = "ZH_TW"  # 預設繁體中文
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance
    
    def __init__(self):
        if not hasattr(self, '_initialized'):
            self._initialized = True
            self._observers = []  # 觀察者列表
    
    @property
    def current_lang(self):
        return self._current_lang
    
    @current_lang.setter
    def current_lang(self, lang):
        if self._current_lang != lang:
            self._current_lang = lang
            self._notify_observers()
    
    def register_observer(self, observer):
        """註冊觀察者（需要更新UI的 Widget）"""
        if observer not in self._observers:
            self._observers.append(observer)
    
    def unregister_observer(self, observer):
        """取消註冊觀察者"""
        if observer in self._observers:
            self._observers.remove(observer)
    
    def _notify_observers(self):
        """通知所有觀察者語言已改變"""
        for observer in self._observers:
            if hasattr(observer, 'refresh_ui_texts'):
                try:
                    observer.refresh_ui_texts()
                except Exception as e:
                    print(f"Error refreshing UI for {observer}: {e}")
    
    def get(self, key, default=""):
        """獲取翻譯文字"""
        translations = Translations.ZH_TW if self._current_lang == "ZH_TW" else Translations.EN
        return translations.get(key, default)
    
    def toggle_language(self):
        """切換語言"""
        self.current_lang = "EN" if self._current_lang == "ZH_TW" else "ZH_TW"
        return self.current_lang


class Translations:
    """翻譯字典"""
    
    EN = {
        # === Common ===
        "app_title": "SPC Chart Processor",
        "select": "Select",
        "cancel": "Cancel",
        "close": "Close",
        "save": "Save",
        "export": "Export",
        "start": "Start",
        "processing": "Processing...",
        "complete": "Complete",
        "error": "Error",
        "warning": "Warning",
        "ready": "Ready.",
        "success": "Success",
        "failed": "Failed",
        
        # === Main Menu ===
        "home": "Home",
        "split_data": "Split Data",
        "oob_spc_system": "OOB / SPC System",
        "cpk_calculator": "Cpk Calculator",
        "tool_matching": "Tool Matching",
        "cl_tighten": "CL Tighten Cal.",
        
        # === Language ===
        "lang_button": "🌐 中文",
        "current_language": "Current Language: English",
        
        # === Buttons ===
        "start_processing": "Start Processing",
        "export_results": "Export Results",
        "browse_files": "Browse Files...",
        "browse_folder": "Browse Folder...",
        "start_cl_calculation": "🚀 Start CL Calculation",
        "export_to_excel": "📁 Export to Excel",
        "clear_results": "🗑️ Clear Results",
        
        # === File Operations ===
        "select_excel_file": "Select Excel File",
        "select_output_folder": "Select Output Folder",
        "select_raw_data_folder": "Select Raw Data Folder",
        "excel_files": "Excel Files",
        "all_files": "All Files",
        
        # === Chart Processing ===
        "show_charts_gui": "Show Charts in GUI",
        "show_by_tool_charts": "Show By Tool Analysis Charts",
        "use_interactive_charts": "Use Interactive Charts",
        "use_batch_id_labels": "Use Batch ID as X-axis Labels",
        "custom_time_range": "Custom Analysis Time Range",
        "enable_custom_range": "Enable Custom Time Range",
        "start_time": "Start Time:",
        "end_time": "End Time:",
        "quick_select": "Quick Select:",
        "last_7_days": "Last 7 Days",
        "last_30_days": "Last 30 Days",
        "last_90_days": "Last 90 Days",
        "this_month": "This Month",
        "last_month": "Last Month",
        
        # === Summary Dashboard ===
        "summary_dashboard": "📊 Summary Dashboard",
        "total_charts": "Total Charts",
        "processed_successfully": "Processed Successfully",
        "no_data": "No Data",
        "charts_with_ooc": "Charts with OOC",
        "charts_with_we": "Charts with WE Rule",
        "charts_with_oob": "Charts with OOB",
        "charts_with_anomaly": "Charts with Anomaly",
        "normal_charts": "Normal Charts",
        
        # === CL Tighten ===
        "need_tighten": "Need Tighten",
        "no_tighten_needed": "No Tighten Needed",
        "chart_info_file": "Chart Information File:",
        "raw_data_folder": "Raw Data Folder:",
        "output_folder": "Output Folder:",
        "calculation_results": "Calculation Results",
        "no_results": "No calculation results yet",
        
        # === Split Data ===
        "input_excel_file": "Input Excel File:",
        "output_folder_label": "Output Folder:",
        "split_results": "Split Results",
        "total_groups": "Total Groups",
        "total_files": "Total Files Generated",
        "split_complete": "Split Complete!",
        
        # === Status Messages ===
        "loading": "Loading...",
        "calculating": "Calculating...",
        "exporting": "Exporting...",
        "exporting_charts": "Exporting chart data...",
        "export_progress": "Export Progress",
        "processing_chart": "Processing",
        "export_cancelled": "Cancelled",
        "export_cancelled_msg": "Export has been cancelled",
        "export_successful": "Export Successful",
        "export_successful_msg": "Excel exported to:",
        "export_failed": "Export Failed",
        "export_failed_msg": "Excel export failed:",
        "file_saved": "File saved successfully",
        "no_file_selected": "No file selected",
        "invalid_file": "Invalid file",
        "operation_cancelled": "Operation cancelled",
        "no_data": "No Data",
        "chart_error": "Chart Error",
        "chart_info_not_loaded": "Chart information not loaded. Please run analysis first.",
        "settings": "Settings",
        "calculation_mode_settings": "Calculation Mode Settings",
        "custom_calculation_mode": "Custom Calculation Interval Mode",
        "custom_mode_hint": "You can freely adjust the date range. The system will calculate Cpk based on the specified interval and automatically compare historical data of equal duration.",
        "auto_mode_hint": "The system will automatically detect the latest data timestamp and calculate Cpk for the most recent 3 months.",
        "start_date": "Start Date",
        "end_date": "End Date",
        
        # === Errors ===
        "error_loading_file": "Error loading file",
        "error_processing": "Error during processing",
        "error_saving": "Error saving file",
        "missing_columns": "Missing required columns",
        
        # === Table Headers ===
        "group_name": "Group Name",
        "chart_name": "Chart Name",
        "chart_id": "Chart ID",
        "material_no": "Material No.",
        "pattern": "Pattern",
        "suggest_ucl": "Suggest UCL",
        "suggest_lcl": "Suggest LCL",
        "static_ucl": "Static UCL",
        "static_lcl": "Static LCL",
        "tighten_needed": "Tighten Needed",
        "status": "Status",
        
        # === Tool Matching ===
        "tool_matching_title": "Tool Matching",
        "browse_files_with_icon": "📁 Browse Files...",
        "example_button": "💾 Example",
        "formula_explanation": "Formula Explanation",
        "mean_index_threshold": "Mean Index Threshold:",
        "sigma_index_threshold": "Sigma Index Threshold:",
        "fill_sample_size": "Fill Sample Size:",
        "data_filter_mode": "Data Filter Mode:",
        "all_data": "All Data",
        "specified_date": "Specified Date (1 month mean/6 months sigma)",
        "latest_entry": "Latest Entry (1 month mean/6 months sigma)",
        "specified_base_date": "Specified Base Date:",
        "run_analysis": "🚀 Run Analysis",
        "select_file_prompt": "Please select a file and click to start analysis.",
        "matching_group": "Matching Group",
        "mean_index": "Mean Index",
        "sigma_index": "Sigma Index",
        "k_value": "K",
        "mean": "Mean",
        "sigma": "Sigma",
        "mean_median": "Mean Median",
        "sigma_median": "Sigma Median",
        "sample_size": "Sample Size",
        "calculation_formula": "📘 Calculation Formula (Click to Expand)",
        "calculation_formula_hide": "📘 Calculation Formula (Click to Hide)",
        
        # === Split Data ===
        "split_data_title": "CSV File Splitting Tool",
        "split_data_description": "This tool can split CSV files with specific formats into multiple independent CSV files.",
        "split_data_type2_desc": "If the SPC Chart format is vertically arranged, please select **Type2** splitting method.",
        "split_data_type3_desc": "If horizontally arranged, please select **Type3** splitting method.",
        "select_input_files": "1. Select Input Files",
        "select_csv_files": "Please select one or more CSV files (separated by semicolon ';')...",
        "select_output_folder_title": "2. Select Output Folder",
        "select_processing_mode": "3. Select Processing Mode",
        "select_file_type": "Select File Type:",
        "type3_horizontal": "Type3_Horizontal (Horizontal Layout)",
        "type2_vertical": "Type2_Vertical (Vertical Layout)",
        "type3_example": "Type3 Example",
        "type2_example": "Type2 Example",
        "processing_progress": "Processing Progress: %p%",
        "browse": "Browse...",
        "start_processing": "Start Processing",
        "ready": "Ready.",
        
        # === SPC Cpk Dashboard ===
        "spc_cpk_dashboard": "SPC Cpk Dashboard",
        "run_analysis": "Run Analysis",
        "download_cpk_detail": "Download Cpk Detail",
        "chart": "Chart:",
        "start": "Start:",
        "end": "End:",
        "custom_time_mode": "Custom Time Mode",
        "cpk": "Cpk",
        "l1_cpk": "L1 Cpk",
        "l2_cpk": "L2 Cpk",
        "long_term_cpk": "Long-Term Cpk",
        "r1": "R1",
        "r2": "R2",
        "k": "K",
        "spc_chart": "SPC Chart",
        "prev": "◀ Prev",
        "next": "Next ▶",
        "no_data": "No Data",
        "chart_info_not_loaded": "Chart information not loaded yet!",
        
        # === Summary Dashboard ===
        "summary_dashboard": "Summary Dashboard",
        "total_charts": "Total Charts:",
        "processed_successfully": "Processed Successfully:",
        "no_data_charts": "No Data:",
        "charts_with_ooc": "Charts with OOC:",
        "charts_with_we_rule": "Charts with WE Rule:",
        "charts_with_oob": "Charts with OOB:",
        "charts_with_anomalies_details": "Charts with Anomalies Details",
        "group_name": "Group Name",
        "chart_name": "Chart Name",
        "ooc_count": "OOC Count",
        "we_rules": "WE Rules",
        "oob_rules": "OOB Rules",
        "processed": "Processed",
        
        # === Custom Time Range ===
        "custom_time_range": "Custom Time Range Analysis",
        "enable_custom_time_range": "Enable Custom Time Range",
        "start_time": "Start Time:",
        "end_time": "End Time:",
        "quick_select": "Quick Select:",
        "last_7_days": "Last 7 Days",
        "last_30_days": "Last 30 Days",
        "last_90_days": "Last 90 Days",
        "this_month": "This Month",
        "last_month": "Last Month",
        
        # === Tool Matching Notice ===
        "notice": "Notice:",
        "notice_abnormal_only": "The table below only shows abnormal items.",
        "mean_not_matched": "Mean Not Matched",
        "sigma_not_matched": "Sigma Not Matched",
        "insufficient_data": "Insufficient Data",
        "insufficient_data_desc": "Sample size < 5, no comparison performed",
        "click_formula_expand": "Click \"Calculation Formula\" below to expand/collapse detailed explanation.",
        
        # === OOB SPC System ===
        "start_process": "Start Process",
        "settings": "Settings",
        "threshold_settings": "Threshold Settings",
        "data_processing_settings": "Data Processing Settings",
        "chart_processing_settings": "Chart Processing Settings",
        "display_settings": "Display Settings",
        "overall_processing_status": "Overall Processing Status",
        "violation_rate": "Violation Rate (Processed Charts)",
        "charts_with_anomalies": "Charts with Anomalies",
        "violating": "Violating",
        "normal": "Normal",
        "all_normal": "All Normal",
        "ooc": "OOC",
        "we_rule": "WE_Rule",
        "oob": "OOB",
        "number_of_charts": "Number of Charts",
        "please_select_csv": "Please select a CSV file...",
        
        # === CL Tighten ===
        "calculation_range": "Calculation Range:",
        "chart_list": "Chart List",
        "search_placeholder": "Search charts...",
        "chart_details": "Chart Details",
        "chart_name_label": "Chart Name:",
        "group_name_label": "Group Name:",
        "current_ucl": "Current UCL:",
        "current_lcl": "Current LCL:",
        "suggested_ucl": "Suggested UCL:",
        "suggested_lcl": "Suggested LCL:",
        "tightening_factor": "Tightening Factor:",
        "data_points": "Data Points:",
        "mean_value": "Mean:",
        "sigma_value": "Sigma:",
        "no_chart_selected": "No chart selected",
        "select_chart_prompt": "Please select a chart from the list to view details",
        "no_data_loaded": "No data loaded",
        "need_tighten": "Need Tighten",
        "no_tighten_needed": "No Tighten Needed",
        "no_data_file": "No Data File",
        "calc_error": "Calc Error",
        "read_error": "Read Error",
        
        # === OOB System Tabs ===
        "chart_processing": "Chart Processing",
        "summary_dashboard_tab": "Summary Dashboard",
        
        # === Data Health Check ===
        "data_health_monitor": "Data Health Monitor",
        "start_check": "▶ Start Check",
        "allchartinfo_excel": "📂 AllChartInfo Excel",
        "export_report": "📁 Export Report",
        "checking": "Checking",
        "no_file_loaded": "No file loaded",
        "total_scanned": "Total Scanned",
        "passed": "Passed",
        "skipped": "Skipped",
        "critical_errors": "Critical Errors",
        "unable_to_execute": "Unable to Execute",
        "only_show_errors": "Only Show Errors",
        "check_details": "Check Details",
        "severity": "Status",
        "location": "Location",
        "issue_description": "Issue Description",
        "suggested_action": "Suggested Action",
        "open_csv": "Open File",
        "open": "📂 Open",
        "n_a": "N/A",
        "path_not_set": "Path not set properly.",
        "file_not_found": "File not found.",
        "csv_file_not_found": "CSV file not found",
        "export_log": "Export Log",
        "export_failed": "Export Failed",
        "export_success": "Report saved to",
        "permission_denied_export": "⚠️ Permission denied: Cannot write to file\n\nThe file might be opened in Excel or another program.\nPlease close the file and try again.",
        
        # === Health Check Messages ===
        "excel_file_not_found": "Excel file not found",
        "permission_denied": "Permission denied: File is locked or in use",
        "permission_denied_action": "⚠️ Please close the Excel file and try again. The file might be opened in Excel or another program.",
        "failed_to_open_excel": "Failed to open Excel",
        "missing_columns": "Missing columns",
        "add_missing_columns": "Add missing columns to Excel.",
        "groupname_chartname_empty": "GroupName or ChartName is empty",
        "fill_in_names": "Fill in the names.",
        "missing_target_ucl_lcl": "Missing Target/UCL/LCL",
        "fields_mandatory": "These fields are mandatory.",
        "lcl_greater_ucl": "Logic: LCL > UCL",
        "lcl_must_le_ucl": "LCL must be <= UCL.",
        "non_numeric_limits": "Non-numeric Control Limits",
        "ensure_limits_numbers": "Ensure limits are numbers.",
        "invalid_characteristic": "Invalid Characteristic",
        "use_nominal_smaller_bigger": "Use Nominal, Smaller, or Bigger.",
        "nominal_requires_usl_lsl": "Nominal requires USL and LSL",
        "fill_both_usl_lsl": "Fill both USL and LSL.",
        "logic_lsl_greater_usl": "Logic: LSL > USL",
        "lsl_must_le_usl": "LSL must be <= USL.",
        "smaller_requires_usl": "Smaller requires USL",
        "fill_usl": "Fill USL.",
        "bigger_requires_lsl": "Bigger requires LSL",
        "fill_lsl": "Fill LSL.",
        "file_not_found_msg": "File Not Found",
        "expected_csv": "Expected",
        "ensure_in_input": "Ensure it is in 'input/raw_charts'.",
        "empty_csv_file": "Empty CSV file",
        "no_data_rows": "CSV has no data rows.",
        "no_point_val_column": "No 'point_val' column",
        "check_csv_header": "Check CSV header.",
        "no_point_time_column": "No 'point_time' column",
        "time_format_error": "Time Format Error",
        "cannot_parse_datetime": "Cannot parse as datetime.",
        "partial_invalid_times": "Partial Invalid Times",
        "some_times_invalid": "Some time values cannot be parsed. Check for NaT/Empty/Invalid format.",
        "permission_denied_csv": "Permission denied: File is locked or in use",
        "close_csv_file": "⚠️ Please close this CSV file if opened in Excel or another program.",
        "read_error": "Read Error",
        "file_corrupted": "File might be corrupted or unreadable.",
        "all_checks_passed": "All checks passed",
        "csv_ready": "CSV file is ready for processing.",
        
        # Action messages (with row number placeholder)
        "check_excel_row_groupname_chartname": "Check Excel row {row}: GroupName and ChartName are mandatory.",
        "check_excel_row_target_ucl_lcl": "Check Excel row {row}: Target, UCL, LCL are mandatory.",
        "check_excel_row_lcl_ucl": "Check Excel row {row}: LCL must not exceed UCL.",
        "check_excel_row_lsl_usl": "Check Excel row {row}: LSL must not exceed USL.",
        "check_excel_row_numeric": "Check Excel row {row}: Control limits must be numeric.",
        "check_excel_row_characteristics": "Check Excel row {row}: Characteristics must be Nominal, Smaller, or Bigger.",
        "check_excel_row_nominal": "Check Excel row {row}: Nominal type requires both USL and LSL.",
        "check_excel_row_smaller": "Check Excel row {row}: Smaller type requires USL.",
        "check_excel_row_bigger": "Check Excel row {row}: Bigger type requires LSL.",
        "check_excel_row_logic_nominal": "Check Excel row {row}: Must satisfy USL >= UCL >= Target >= LCL >= LSL.",
        "check_excel_row_logic_smaller": "Check Excel row {row}: Smaller type must satisfy USL >= UCL >= Target >= LCL.",
        "check_excel_row_logic_bigger": "Check Excel row {row}: Bigger type must satisfy UCL >= Target >= LCL >= LSL.",
        "csv_empty_file": "CSV file is empty.",
        "csv_missing_point_val": "CSV file is missing 'point_val' column.",
        "csv_missing_point_time": "CSV file is missing 'point_time' column.",
        "csv_time_format_error": "Time format error. Correct format should be '%Y/%m/%d %H:%M'.",
        "csv_partial_invalid_times": "Some time values are invalid.",
        "csv_permission_denied": "⚠️ File is locked or in use. Please close this CSV file.",
        "csv_read_error": "File read error, may be corrupted.",
        "diagnosis": "🔍 Diagnosis",
        "csv_files_not_found": "CSV files not found",
        "likely_wrong_path": "⚠️ Likely cause: Wrong raw_data_dir path or incorrect file naming. Check 'input/raw_charts' folder location and ensure files follow 'GroupName_ChartName.csv' format.",
        "csv_found_but_errors": "All CSV files found but have errors",
        "csv_content_issue": "⚠️ Likely cause: CSV content issue. Ensure 'point_val' and 'point_time' columns exist and time format is '%Y/%m/%d %H:%M'.",
        "excel_config_errors": "rows have Excel config errors",
        "fix_excel_first": "⚠️ Likely cause: AllChartInfo Excel has missing/invalid values. Fix Excel configuration first before CSV checks can proceed.",
        "unexpected_crash": "Unexpected Crash",
        "contact_developer": "Contact Developer.",
        
        # === Preprocessing ===
        "preprocessing_chart_types": "Preprocessing chart types",
        "preprocessing_complete_starting_charts": "Data type preprocessing complete, starting chart processing...",
    }
    
    ZH_TW = {
        # === 通用 ===
        "app_title": "SPC 圖表處理系統",
        "select": "選擇",
        "cancel": "取消",
        "close": "關閉",
        "save": "儲存",
        "export": "匯出",
        "start": "開始",
        "processing": "處理中...",
        "complete": "完成",
        "error": "錯誤",
        "warning": "警告",
        "ready": "準備就緒。",
        "success": "成功",
        "failed": "失敗",
        
        # === 主選單 ===
        "home": "首頁",
        "split_data": "資料拆分",
        "oob_spc_system": "OOB / SPC 分析系統",
        "cpk_calculator": "Cpk 儀表板",
        "tool_matching": "機台一致性分析",
        "cl_tighten": "管制界線計算",
        
        # === 語言 ===
        "lang_button": "🌐 EN",
        "current_language": "目前語言：繁體中文",
        
        # === 按鈕 ===
        "start_processing": "開始執行",
        "export_results": "匯出結果",
        "browse_files": "瀏覽檔案...",
        "browse_folder": "瀏覽資料夾...",
        "start_cl_calculation": "🚀 開始 CL 計算",
        "export_to_excel": "📁 匯出至 Excel",
        "clear_results": "🗑️ 清除結果",
        
        # === 檔案操作 ===
        "select_excel_file": "選擇 Excel 檔案",
        "select_output_folder": "選擇輸出資料夾",
        "select_raw_data_folder": "選擇原始數據資料夾",
        "excel_files": "Excel 檔案",
        "all_files": "所有檔案",
        
        # === 圖表處理 ===
        "show_charts_gui": "在介面顯示圖表",
        "show_by_tool_charts": "顯示機台分析圖表 (By Tool)",
        "use_interactive_charts": "使用互動式圖表",
        "use_batch_id_labels": "使用 Batch ID 作為 X 軸標籤",
        "custom_time_range": "自訂分析時間範圍",
        "enable_custom_range": "啟用自訂時間範圍",
        "start_time": "開始時間：",
        "end_time": "結束時間：",
        "quick_select": "快速選擇：",
        "last_7_days": "最近 7 天",
        "last_30_days": "最近 30 天",
        "last_90_days": "最近 90 天",
        "this_month": "本月",
        "last_month": "上月",
        
        # === 統計儀表板 ===
        "summary_dashboard": "📊 統計儀表板",
        "total_charts": "總圖表數",
        "processed_successfully": "成功處理",
        "no_data": "無資料",
        "charts_with_ooc": "含 OOC 圖表",
        "charts_with_we": "含 WE 規則圖表",
        "charts_with_oob": "含 OOB 圖表",
        "charts_with_anomaly": "含異常圖表",
        "normal_charts": "正常圖表",
        
        # === CL 收緊 ===
        "need_tighten": "需要收緊",
        "no_tighten_needed": "無需收緊",
        "chart_info_file": "圖表資訊檔案：",
        "raw_data_folder": "原始數據資料夾：",
        "output_folder": "輸出資料夾：",
        "calculation_results": "計算結果",
        "no_results": "尚無計算結果",
        
        # === 資料拆分 ===
        "input_excel_file": "輸入 Excel 檔案：",
        "output_folder_label": "輸出資料夾：",
        "split_results": "拆分結果",
        "total_groups": "總群組數",
        "total_files": "總生成檔案數",
        "split_complete": "拆分完成！",
        
        # === 狀態訊息 ===
        "loading": "載入中...",
        "calculating": "計算中...",
        "exporting": "匯出中...",
        "exporting_charts": "正在匯出圖表資料...",
        "export_progress": "匯出進度",
        "processing_chart": "正在處理",
        "export_cancelled": "已取消",
        "export_cancelled_msg": "匯出已被取消",
        "export_successful": "匯出成功",
        "export_successful_msg": "Excel 已匯出至：",
        "export_failed": "匯出失敗",
        "export_failed_msg": "Excel 匯出失敗：",
        "file_saved": "檔案儲存成功",
        "no_file_selected": "未選擇檔案",
        "invalid_file": "無效檔案",
        "operation_cancelled": "操作已取消",        "no_data": "無資料",
        "chart_error": "圖表錯誤",
        "chart_info_not_loaded": "圖表資訊尚未載入，請先執行分析。",
        "settings": "設定",
        "calculation_mode_settings": "計算模式設定",
        "custom_calculation_mode": "自訂計算區間模式",
        "custom_mode_hint": "您可以自由調整日期範圍，系統將根據指定區間計算 Cpk，並自動對比等長度的歷史資料。",
        "auto_mode_hint": "系統將自動偵測最新資料時間點，計算最近 3 個月的 Cpk。",
        "start_date": "起始日期",
        "end_date": "結束日期",        
        # === 錯誤訊息 ===
        "error_loading_file": "載入檔案時發生錯誤",
        "error_processing": "處理過程中發生錯誤",
        "error_saving": "儲存檔案時發生錯誤",
        "missing_columns": "缺少必要欄位",
        
        # === 表格標題 ===
        "group_name": "群組名稱",
        "chart_name": "圖表名稱",
        "chart_id": "圖表 ID",
        "material_no": "料號",
        "pattern": "模式",
        "suggest_ucl": "建議 UCL",
        "suggest_lcl": "建議 LCL",
        "static_ucl": "靜態 UCL",
        "static_lcl": "靜態 LCL",
        "tighten_needed": "需要收緊",
        "status": "狀態",
        
        # === 機台配對 ===
        "tool_matching_title": "機台配對",
        "browse_files_with_icon": "📁 瀏覽檔案...",
        "example_button": "💾 範例",
        "formula_explanation": "公式說明",
        "mean_index_threshold": "均值指標門檻：",
        "sigma_index_threshold": "標準差指標門檻：",
        "fill_sample_size": "補滿樣本數：",
        "data_filter_mode": "資料篩選模式：",
        "all_data": "全部資料",
        "specified_date": "指定日期 (1個月均值/6個月標準差)",
        "latest_entry": "最新資料 (1個月均值/6個月標準差)",
        "specified_base_date": "指定基準日期：",
        "run_analysis": "🚀 執行分析",
        "select_file_prompt": "請選擇檔案並點擊開始分析。",
        "matching_group": "配對群組",
        "mean_index": "均值指標",
        "sigma_index": "標準差指標",
        "k_value": "K 值",
        "mean": "均值",
        "sigma": "標準差",
        "mean_median": "均值中位數",
        "sigma_median": "標準差中位數",
        "sample_size": "樣本數",
        "calculation_formula": "📘 計算公式 (點擊展開)",
        "calculation_formula_hide": "📘 計算公式 (點擊收合)",
        
        # === 資料拆分 ===
        "split_data_title": "CSV 檔案拆分工具",
        "split_data_description": "本工具可將特定格式的 CSV 檔案拆分為多個獨立的 CSV 檔案。",
        "split_data_type2_desc": "如果 SPC Chart 格式是垂直排列，請選擇 **Type2** 拆分方式。",
        "split_data_type3_desc": "如果水平排列，請選擇 **Type3** 拆分方式。",
        "select_input_files": "1. 選擇輸入檔案",
        "select_csv_files": "請選擇一個或多個 CSV 檔案 (多個檔案請以分號 ';' 分隔)...",
        "select_output_folder_title": "2. 選擇輸出資料夾",
        "select_processing_mode": "3. 選擇處理模式",
        "select_file_type": "選擇檔案類型：",
        "type3_horizontal": "Type3_橫向 (水平排列)",
        "type2_vertical": "Type2_縱向 (垂直排列)",
        "type3_example": "Type3 範例",
        "type2_example": "Type2 範例",
        "processing_progress": "處理進度: %p%",
        "browse": "瀏覽...",
        "start_processing": "開始處理",
        "ready": "準備就緒。",
        
        # === SPC Cpk 儀表板 ===
        "spc_cpk_dashboard": "SPC Cpk 儀表板",
        "run_analysis": "執行分析",
        "download_cpk_detail": "下載 Cpk 詳細資料",
        "chart": "圖表：",
        "start": "開始：",
        "end": "結束：",
        "custom_time_mode": "自訂時間模式",
        "cpk": "Cpk",
        "l1_cpk": "L1 Cpk",
        "l2_cpk": "L2 Cpk",
        "long_term_cpk": "長期 Cpk",
        "r1": "R1",
        "r2": "R2",
        "k": "K",
        "spc_chart": "SPC 圖表",
        "prev": "◀ 上一個",
        "next": "下一個 ▶",
        "no_data": "無資料",
        "chart_info_not_loaded": "圖表資訊尚未載入！",
        
        # === 摘要儀表板 ===
        "summary_dashboard": "摘要儀表板",
        "total_charts": "總圖表數：",
        "processed_successfully": "成功處理：",
        "no_data_charts": "無資料：",
        "charts_with_ooc": "含 OOC 圖表：",
        "charts_with_we_rule": "含 WE 規則圖表：",
        "charts_with_oob": "含 OOB 圖表：",
        "charts_with_anomalies_details": "異常圖表詳細資料",
        "group_name": "群組名稱",
        "chart_name": "圖表名稱",
        "ooc_count": "OOC 次數",
        "we_rules": "WE 規則",
        "oob_rules": "OOB 規則",
        "processed": "已處理",
        
        # === 自訂時間範圍 ===
        "custom_time_range": "自訂時間分析範圍",
        "enable_custom_time_range": "啟用自訂時間範圍",
        "start_time": "開始時間：",
        "end_time": "結束時間：",
        "quick_select": "快速選擇：",
        "last_7_days": "最近 7 天",
        "last_30_days": "最近 30 天",
        "last_90_days": "最近 90 天",
        "this_month": "本月",
        "last_month": "上個月",
        
        # === 機台配對注意事項 ===
        "notice": "注意：",
        "notice_abnormal_only": "下方表格僅顯示異常項目。",
        "mean_not_matched": "平均值不匹配",
        "sigma_not_matched": "變異數不匹配",
        "insufficient_data": "資料不足",
        "insufficient_data_desc": "樣本數 < 5，未執行比對",
        "click_formula_expand": "點擊下方「計算公式」可展開/收合詳細說明。",
        
        # === OOB SPC 系統 ===
        "start_process": "開始處理",
        "settings": "設定",
        "threshold_settings": "閾值設定",
        "data_processing_settings": "數據處理設定",
        "chart_processing_settings": "圖表處理設定",
        "display_settings": "顯示設定",
        "overall_processing_status": "整體處理狀態",
        "violation_rate": "違規率（已處理圖表）",
        "charts_with_anomalies": "異常圖表",
        "violating": "違規",
        "normal": "正常",
        "all_normal": "全部正常",
        "ooc": "OOC",
        "we_rule": "WE_Rule",
        "oob": "OOB",
        "number_of_charts": "圖表數量",
        "please_select_csv": "請選擇一個 CSV 檔案...",
        
        # === 管制線收緊 ===
        "calculation_range": "計算區間：",
        "chart_list": "圖表清單",
        "search_placeholder": "搜尋圖表...",
        "chart_details": "圖表詳細資訊",
        "chart_name_label": "圖表名稱：",
        "group_name_label": "群組名稱：",
        "current_ucl": "目前 UCL：",
        "current_lcl": "目前 LCL：",
        "suggested_ucl": "建議 UCL：",
        "suggested_lcl": "建議 LCL：",
        "tightening_factor": "收緊係數：",
        "data_points": "資料點數：",
        "mean_value": "平均值：",
        "sigma_value": "標準差：",
        "no_chart_selected": "未選擇圖表",
        "select_chart_prompt": "請從清單中選擇圖表以查看詳細資訊",
        "no_data_loaded": "未載入資料",
        "need_tighten": "需要收緊",
        "no_tighten_needed": "無需收緊",
        "no_data_file": "無資料檔案",
        "calc_error": "計算錯誤",
        "read_error": "讀取錯誤",
        
        # === OOB 系統標籤頁 ===
        "chart_processing": "圖表處理",
        "summary_dashboard_tab": "摘要儀表板",
        
        # === 資料健康檢查 ===
        "data_health_monitor": "資料健康監測",
        "start_check": "▶ 開始檢查",
        "allchartinfo_excel": "📂 AllChartInfo Excel",
        "export_report": "📁 匯出報告",
        "checking": "檢查中",
        "no_file_loaded": "未載入檔案",
        "total_scanned": "總掃描數",
        "passed": "通過",
        "skipped": "跳過",
        "critical_errors": "嚴重錯誤",
        "unable_to_execute": "無法執行",
        "only_show_errors": "只顯示錯誤項目",
        "check_details": "檢查詳情",
        "severity": "狀態",
        "location": "位置",
        "issue_description": "問題描述",
        "suggested_action": "建議措施",
        "open_csv": "開啟檔案",
        "open": "📂 開啟",
        "n_a": "無",
        "path_not_set": "路徑設定不正確。",
        "file_not_found": "找不到檔案。",
        "csv_file_not_found": "找不到 CSV 檔案",
        "export_log": "匯出日誌",
        "export_failed": "匯出失敗",
        "export_success": "報告已儲存至",
        "permission_denied_export": "⚠️ 權限被拒：無法寫入檔案\n\n該檔案可能已在 Excel 或其他程式中開啟。\n請關閉檔案後重試。",
        
        # === 健康檢查訊息 ===
        "excel_file_not_found": "找不到 Excel 檔案",
        "permission_denied": "權限被拒：檔案已鎖定或使用中",
        "permission_denied_action": "⚠️ 請關閉 Excel 檔案後重試。該檔案可能已在 Excel 或其他程式中開啟。",
        "failed_to_open_excel": "開啟 Excel 失敗",
        "missing_columns": "缺少欄位",
        "add_missing_columns": "請在 Excel 中新增缺少的欄位。",
        "groupname_chartname_empty": "GroupName 或 ChartName 為空",
        "fill_in_names": "請填入名稱。",
        "missing_target_ucl_lcl": "缺少 Target/UCL/LCL",
        "fields_mandatory": "這些欄位為必填。",
        "lcl_greater_ucl": "邏輯錯誤：LCL > UCL",
        "lcl_must_le_ucl": "LCL 必須 <= UCL。",
        "non_numeric_limits": "管制界限非數值",
        "ensure_limits_numbers": "確保界限為數字。",
        "invalid_characteristic": "無效的 Characteristic",
        "use_nominal_smaller_bigger": "請使用 Nominal、Smaller 或 Bigger。",
        "nominal_requires_usl_lsl": "Nominal 需要 USL 和 LSL",
        "fill_both_usl_lsl": "請填入 USL 和 LSL。",
        "logic_lsl_greater_usl": "邏輯錯誤：LSL > USL",
        "lsl_must_le_usl": "LSL 必須 <= USL。",
        "smaller_requires_usl": "Smaller 需要 USL",
        "fill_usl": "請填入 USL。",
        "bigger_requires_lsl": "Bigger 需要 LSL",
        "fill_lsl": "請填入 LSL。",
        "file_not_found_msg": "找不到檔案",
        "expected_csv": "預期檔案",
        "ensure_in_input": "確保檔案在 'input/raw_charts' 中。",
        "empty_csv_file": "CSV 檔案為空",
        "no_data_rows": "CSV 沒有資料列。",
        "no_point_val_column": "缺少 'point_val' 欄位",
        "check_csv_header": "檢查 CSV 標題。",
        "no_point_time_column": "缺少 'point_time' 欄位",
        "time_format_error": "時間格式錯誤",
        "cannot_parse_datetime": "無法解析為日期時間。",
        "partial_invalid_times": "部分時間無效",
        "some_times_invalid": "部分時間值無法解析。請檢查是否有 NaT/空值/無效格式。",
        "permission_denied_csv": "權限被拒：檔案已鎖定或使用中",
        "close_csv_file": "⚠️ 如果此 CSV 檔案在 Excel 或其他程式中開啟，請關閉它。",
        "read_error": "讀取錯誤",
        "file_corrupted": "檔案可能已損壞或無法讀取。",
        "all_checks_passed": "所有檢查通過",
        "csv_ready": "CSV 檔案可供處理。",
        
        # Action 訊息（帶行號占位符）
        "check_excel_row_groupname_chartname": "請檢查 Excel 第 {row} 行：GroupName 與 ChartName 為必填項。",
        "check_excel_row_target_ucl_lcl": "請檢查 Excel 第 {row} 行：Target、UCL、LCL 為必填項。",
        "check_excel_row_lcl_ucl": "請檢查 Excel 第 {row} 行：LCL 不得大於 UCL。",
        "check_excel_row_lsl_usl": "請檢查 Excel 第 {row} 行：LSL 不得大於 USL。",
        "check_excel_row_numeric": "請檢查 Excel 第 {row} 行：管制界限必須為數值。",
        "check_excel_row_characteristics": "請檢查 Excel 第 {row} 行：Characteristics 必須為 Nominal、Smaller 或 Bigger。",
        "check_excel_row_nominal": "請檢查 Excel 第 {row} 行：Nominal 類型需要同時填寫 USL 與 LSL。",
        "check_excel_row_smaller": "請檢查 Excel 第 {row} 行：Smaller 類型需要填寫 USL。",
        "check_excel_row_bigger": "請檢查 Excel 第 {row} 行：Bigger 類型需要填寫 LSL。",
        "check_excel_row_logic_nominal": "請檢查 Excel 第 {row} 行：必須滿足 USL >= UCL >= Target >= LCL >= LSL。",
        "check_excel_row_logic_smaller": "請檢查 Excel 第 {row} 行：Smaller 類型必須滿足 USL >= UCL >= Target >= LCL。",
        "check_excel_row_logic_bigger": "請檢查 Excel 第 {row} 行：Bigger 類型必須滿足 UCL >= Target >= LCL >= LSL。",
        "csv_empty_file": "CSV 檔案無資料。",
        "csv_missing_point_val": "CSV 檔案缺少 'point_val' 欄位。",
        "csv_missing_point_time": "CSV 檔案缺少 'point_time' 欄位。",
        "csv_time_format_error": "時間格式錯誤。正確格式應為 '%Y/%m/%d %H:%M'。",
        "csv_partial_invalid_times": "部分時間值無效。",
        "csv_permission_denied": "⚠️ 檔案被鎖定或正在使用中，請關閉此 CSV 檔案。",
        "csv_read_error": "檔案讀取錯誤，可能已損毀。",
        "diagnosis": "🔍 診斷",
        "csv_files_not_found": "找不到 CSV 檔案",
        "likely_wrong_path": "⚠️ 可能原因：raw_data_dir 路徑錯誤或檔名不正確。請檢查 'input/raw_charts' 資料夾位置，並確保檔案遵循 'GroupName_ChartName.csv' 格式。",
        "csv_found_but_errors": "所有 CSV 檔案已找到但有錯誤",
        "csv_content_issue": "⚠️ 可能原因：CSV 內容問題。請確保存在 'point_val' 和 'point_time' 欄位，且時間格式為 '%Y/%m/%d %H:%M'。",
        "excel_config_errors": "行有 Excel 配置錯誤",
        "fix_excel_first": "⚠️ 可能原因：AllChartInfo Excel 有缺失/無效值。請先修正 Excel 配置，再進行 CSV 檢查。",
        "unexpected_crash": "意外崩潰",
        "contact_developer": "請聯繫開發人員。",
        
        # === 預處理 ===
        "preprocessing_chart_types": "預處理圖表數據類型",
        "preprocessing_complete_starting_charts": "數據類型預處理完成，開始圖表處理...",
    }


# 創建全域翻譯管理器實例
_translator = TranslationManager()

def get_translator():
    """獲取全域翻譯管理器"""
    return _translator

def tr(key, default=""):
    """快速翻譯函數"""
    return _translator.get(key, default)
