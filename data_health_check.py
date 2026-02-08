import os
import sys
import pandas as pd
import numpy as np
from PyQt6 import QtWidgets, QtGui, QtCore
from PyQt6.QtCore import Qt
import pandas as pd
import os
import sys
from difflib import SequenceMatcher
from translations import tr, get_translator

# ============================================================================
# 1. Worker Thread (負責邏輯與資料檢查)
# ============================================================================
class DataValidatorWorker(QtCore.QThread):
    """
    背景工作執行緒：負責執行實際的資料檢查邏輯。
    包含 Excel 邏輯檢查與 CSV 格式/時間格式檢查。
    """
    progress_updated = QtCore.pyqtSignal(int, int)   # current, total
    log_added = QtCore.pyqtSignal(dict)              # Log entry
    stats_updated = QtCore.pyqtSignal(int, int, int, int) # unable_to_execute, warning, pass, skipped
    finished_check = QtCore.pyqtSignal(bool)         # is_success (no unable_to_execute errors)

    def __init__(self, excel_path, raw_data_dir):
        super().__init__()
        self.excel_path = excel_path
        self.raw_data_dir = raw_data_dir
        self._is_running = True

    def stop(self):
        self._is_running = False

    def find_csv_file_by_prefix(self, directory, prefix):
        """
        在指定目錄中尋找以指定前綴開頭的CSV檔案
        :param directory: 搜尋目錄
        :param prefix: 檔案前綴 (如 "GroupName_ChartName")
        :return: 找到的檔案完整路徑，如果沒找到返回 None
        """
        try:
            if not os.path.exists(directory):
                return None
            
            # 列出目錄中所有檔案
            all_files = os.listdir(directory)
            
            # 篩選出以指定前綴開頭且以.csv結尾的檔案
            matching_files = [
                f for f in all_files 
                if f.startswith(prefix) and f.lower().endswith('.csv')
            ]
            
            if matching_files:
                # 如果找到多個，選擇第一個（可以根據需要調整排序邏輯）
                selected_file = sorted(matching_files)[0]  # 按字母順序排序後選第一個
                return os.path.join(directory, selected_file)
            
            return None
            
        except Exception as e:
            print(f"Error searching for CSV files: {e}")
            return None

    def run(self):
        print("[DEBUG] DataValidatorWorker started")
        
        unable_count = 0
        warning_count = 0
        pass_count = 0
        skipped_count = 0
        
        # 診斷用計數器
        file_not_found_count = 0  # CSV 檔案不存在
        file_read_error_count = 0  # CSV 檔案存在但讀取失敗
        file_format_error_count = 0  # CSV 檔案欄位或格式錯誤
        excel_logic_error_count = 0  # Excel 邏輯錯誤（阻止 CSV 檢查）
        
        # 1. 檢查 Excel 檔案是否存在
        if not os.path.exists(self.excel_path):
            self.emit_log("Unable to Execute", "System", f"Excel file not found: {self.excel_path}", "Please check file path. Ensure it is in 'input/raw_charts'.", source="System")
            self.finished_check.emit(False)
            return

        try:
            # 2. 讀取 Excel
            try:
                df_info = pd.read_excel(self.excel_path, sheet_name='Chart', engine='openpyxl')
            except PermissionError:
                self.emit_log("Unable to Execute", "Excel Load", 
                              f"Permission denied: File is locked or in use", 
                              "⚠️ Please close the Excel file and try again. The file might be opened in Excel or another program.", source="Excel")
                self.finished_check.emit(False)
                return
            except Exception as e:
                self.emit_log("Unable to Execute", "Excel Load", f"Failed to open Excel: {e}", "Check if file is corrupted or path is correct.", source="Excel")
                self.finished_check.emit(False)
                return

            total_rows = len(df_info)
            
            # 檢查 Excel 必要欄位
            required_cols = ['GroupName', 'ChartName', 'UCL', 'LCL', 'Target', 'USL', 'LSL', 'Characteristics']
            missing_cols = [col for col in required_cols if col not in df_info.columns]
            
            if missing_cols:
                self.emit_log("Unable to Execute", "Excel Header", f"Missing columns: {missing_cols}", "Add missing columns to Excel.", source="Excel")
                self.finished_check.emit(False)
                return

            # 3. 逐行檢查
            for i, row in df_info.iterrows():
                if not self._is_running: 
                    print(f"[DEBUG] Worker stopped by user at row {i+1}/{total_rows}")
                    break
                
                row_num = i + 2
                has_error_in_row = False
                
                # 除錯資訊：顯示目前處理進度
                if i % 100 == 0 or i >= total_rows - 5:  # 每100行或最後5行顯示進度
                    print(f"[DEBUG] Processing row {i+1}/{total_rows} (Excel row {row_num})")
                
                # --- A. 讀取欄位資料 ---
                group = str(row.get('GroupName', '')).strip()
                chart = str(row.get('ChartName', '')).strip()
                chart_id = f"{group}_{chart}" if group and chart else f"Row {row_num}"

                target = row.get('Target')
                ucl = row.get('UCL')
                lcl = row.get('LCL')
                usl = row.get('USL')
                lsl = row.get('LSL')
                char_type = str(row.get('Characteristics', '')).strip()

                # --- B. Excel 邏輯檢查 ---

                # B1. 檢查名稱
                if not group or not chart or group.lower() == 'nan' or chart.lower() == 'nan':
                    self.emit_log("Unable to Execute", f"Row {row_num}", "GroupName or ChartName is empty", f"Check Excel row {row_num}: GroupName and ChartName are mandatory.", source="Excel", row_num=row_num)
                    unable_count += 1
                    has_error_in_row = True
                    excel_logic_error_count += 1

                # B2. 檢查管制界限 (Target, UCL, LCL 必填)
                if pd.isna(target) or pd.isna(ucl) or pd.isna(lcl):
                    self.emit_log("Unable to Execute", f"Row {row_num} ({chart_id})", "Missing Target/UCL/LCL", f"Check Excel row {row_num}: Target, UCL, LCL are mandatory.", source="Excel", row_num=row_num)
                    unable_count += 1
                    excel_logic_error_count += 1
                    has_error_in_row = True
                else:
                    try:
                        # 檢查是否為數值（不檢查大小關係，留給 B4 統一檢查）
                        float(target)
                        float(ucl)
                        float(lcl)
                    except:
                        self.emit_log("Unable to Execute", f"Row {row_num} ({chart_id})", "Non-numeric Control Limits", f"Check Excel row {row_num}: Control limits must be numeric.", source="Excel", row_num=row_num)
                        unable_count += 1
                        excel_logic_error_count += 1
                        has_error_in_row = True

                # B3. 檢查規格界限 (依據 Characteristics)
                valid_types = ['Nominal', 'Smaller', 'Bigger']
                match_type = next((t for t in valid_types if t.lower() == char_type.lower()), None)
                
                if not match_type:
                    excel_logic_error_count += 1
                    self.emit_log("Unable to Execute", f"Row {row_num} ({chart_id})", f"Invalid Characteristic: '{char_type}'", f"Check Excel row {row_num}: Characteristics must be Nominal, Smaller, or Bigger.", source="Excel", row_num=row_num)
                    unable_count += 1
                    has_error_in_row = True
                else:
                    if match_type == 'Nominal':
                        if pd.isna(usl) or pd.isna(lsl):
                            self.emit_log("Unable to Execute", f"Row {row_num} ({chart_id})", "Nominal requires USL and LSL", f"Check Excel row {row_num}: Nominal type requires both USL and LSL.", source="Excel", row_num=row_num)
                            unable_count += 1
                            has_error_in_row = True
                            excel_logic_error_count += 1
                        # 不在這裡檢查 LSL vs USL，留給 B4 統一檢查
                    
                    elif match_type == 'Smaller':
                        if pd.isna(usl):
                            self.emit_log("Unable to Execute", f"Row {row_num} ({chart_id})", "Smaller requires USL", f"Check Excel row {row_num}: Smaller type requires USL.", source="Excel", row_num=row_num)
                            unable_count += 1
                            has_error_in_row = True
                            excel_logic_error_count += 1
                    
                    elif match_type == 'Bigger':
                        if pd.isna(lsl):
                            self.emit_log("Unable to Execute", f"Row {row_num} ({chart_id})", "Bigger requires LSL", f"Check Excel row {row_num}: Bigger type requires LSL.", source="Excel", row_num=row_num)
                            unable_count += 1
                            has_error_in_row = True
                            excel_logic_error_count += 1
                
                # B4. 檢查完整邏輯關係：USL >= UCL >= Target >= LCL >= LSL
                # 移除 has_error_in_row 條件，確保所有邏輯關係都被檢查
                if match_type == 'Nominal' and not pd.isna(usl) and not pd.isna(lsl):
                    # 只有 Nominal 需要檢查完整的五個值的順序
                    try:
                        target_val = float(target)
                        ucl_val = float(ucl)
                        lcl_val = float(lcl)
                        usl_val = float(usl)
                        lsl_val = float(lsl)
                        
                        # 檢查完整的邏輯順序
                        violations = []
                        if usl_val < ucl_val:
                            violations.append(f"USL ({usl_val}) < UCL ({ucl_val})")
                        if ucl_val < target_val:
                            violations.append(f"UCL ({ucl_val}) < Target ({target_val})")
                        if target_val < lcl_val:
                            violations.append(f"Target ({target_val}) < LCL ({lcl_val})")
                        if lcl_val < lsl_val:
                            violations.append(f"LCL ({lcl_val}) < LSL ({lsl_val})")
                        
                        if violations:
                            violation_text = "; ".join(violations)
                            self.emit_log("Unable to Execute", f"Row {row_num} ({chart_id})", 
                                        f"Logic violation: {violation_text}", 
                                        f"Check Excel row {row_num}: Must satisfy USL >= UCL >= Target >= LCL >= LSL.", 
                                        source="Excel", row_num=row_num)
                            unable_count += 1
                            has_error_in_row = True
                            excel_logic_error_count += 1
                    except (ValueError, TypeError):
                        # 如果轉換失敗，前面的檢查已經捕捉到了，這裡不重複報錯
                        pass
                elif match_type == 'Smaller' and not pd.isna(usl):
                    # Smaller 只檢查 USL >= UCL >= Target >= LCL
                    try:
                        target_val = float(target)
                        ucl_val = float(ucl)
                        lcl_val = float(lcl)
                        usl_val = float(usl)
                        
                        violations = []
                        if usl_val < ucl_val:
                            violations.append(f"USL ({usl_val}) < UCL ({ucl_val})")
                        if ucl_val < target_val:
                            violations.append(f"UCL ({ucl_val}) < Target ({target_val})")
                        if target_val < lcl_val:
                            violations.append(f"Target ({target_val}) < LCL ({lcl_val})")
                        
                        if violations:
                            violation_text = "; ".join(violations)
                            self.emit_log("Unable to Execute", f"Row {row_num} ({chart_id})", 
                                        f"Logic violation: {violation_text}", 
                                        f"Check Excel row {row_num}: Smaller type must satisfy USL >= UCL >= Target >= LCL.", 
                                        source="Excel", row_num=row_num)
                            unable_count += 1
                            has_error_in_row = True
                            excel_logic_error_count += 1
                    except (ValueError, TypeError):
                        pass
                elif match_type == 'Bigger' and not pd.isna(lsl):
                    # Bigger 只檢查 UCL >= Target >= LCL >= LSL
                    try:
                        target_val = float(target)
                        ucl_val = float(ucl)
                        lcl_val = float(lcl)
                        lsl_val = float(lsl)
                        
                        violations = []
                        if ucl_val < target_val:
                            violations.append(f"UCL ({ucl_val}) < Target ({target_val})")
                        if target_val < lcl_val:
                            violations.append(f"Target ({target_val}) < LCL ({lcl_val})")
                        if lcl_val < lsl_val:
                            violations.append(f"LCL ({lcl_val}) < LSL ({lsl_val})")
                        
                        if violations:
                            violation_text = "; ".join(violations)
                            self.emit_log("Unable to Execute", f"Row {row_num} ({chart_id})", 
                                        f"Logic violation: {violation_text}", 
                                        f"Check Excel row {row_num}: Bigger type must satisfy UCL >= Target >= LCL >= LSL.", 
                                        source="Excel", row_num=row_num)
                            unable_count += 1
                            has_error_in_row = True
                            excel_logic_error_count += 1
                    except (ValueError, TypeError):
                        pass

                # --- C. CSV 檔案檢查 (模糊比對：以 groupname_chartname 開頭) ---
                csv_prefix = f"{group}_{chart}"
                found_file = self.find_csv_file_by_prefix(self.raw_data_dir, csv_prefix)
                csv_status = "Not Checked"
                
                # 如果 Excel 有錯誤，完全跳過 CSV 檢查（該行已計入 critical_count，不重複計數）
                if has_error_in_row:
                    csv_status = "Skipped (Excel Error)"
                elif not has_error_in_row:
                    if not found_file:
                        self.emit_log("Skipped", f"CSV ({chart_id})", "File Not Found", f"No file found with prefix: {csv_prefix}*. Ensure it is in 'input'.", csv_prefix)
                        skipped_count += 1
                        file_not_found_count += 1
                        csv_status = "File Not Found (Skipped)"
                    else:
                        try:
                            # 顯示實際找到的檔案名稱
                            actual_filename = os.path.basename(found_file)
                            print(f"  [DEBUG] 找到匹配檔案: {actual_filename} (前綴: {csv_prefix})")
                            
                            # 加入超時保護：只讀取前5行進行驗證
                            df_csv = pd.read_csv(found_file, nrows=5)
                            
                            if df_csv.empty:
                                self.emit_log("Unable to Execute", f"CSV ({chart_id})", "Empty CSV file", "CSV file is empty.", found_file, source="CSV")
                                unable_count += 1
                                has_error_in_row = True
                                file_format_error_count += 1
                                csv_status = "Empty File"
                            else:
                                # 檢查必要欄位
                                if 'point_val' not in df_csv.columns:
                                    self.emit_log("Unable to Execute", f"CSV ({chart_id})", "No 'point_val' column", "CSV file is missing 'point_val' column.", found_file, source="CSV")
                                    unable_count += 1
                                    has_error_in_row = True
                                    file_format_error_count += 1
                                    csv_status = "Missing point_val"
                                
                                if 'point_time' not in df_csv.columns:
                                    self.emit_log("Unable to Execute", f"CSV ({chart_id})", "No 'point_time' column", "CSV file is missing 'point_time' column.", found_file, source="CSV")
                                    unable_count += 1
                                    has_error_in_row = True
                                    file_format_error_count += 1
                                    csv_status = "Missing point_time"
                                elif 'point_val' in df_csv.columns:  # 只有當兩個欄位都存在時才檢查時間格式
                                    try:
                                        raw_times = df_csv['point_time']
                                        target_format = '%Y/%m/%d %H:%M'
                                        print(f"  [DEBUG] 檢查時間格式，共 {len(raw_times)} 筆時間資料")
                                        
                                        # 加入超時保護：限制時間轉換處理
                                        converted_times = pd.to_datetime(raw_times, format=target_format, errors='coerce')
                                        
                                        if converted_times.isna().all():
                                            print(f"  [DEBUG] 目標格式失敗，嘗試混合格式解析")
                                            converted_times = pd.to_datetime(raw_times, format='mixed', errors='coerce')
                                        
                                        # 檢查轉換結果
                                        if converted_times.isna().all():
                                            example = raw_times.iloc[0] if len(raw_times) > 0 else "N/A"
                                            print(f"  [DEBUG] 所有時間解析失敗，範例值: {example}")
                                            self.emit_log("Unable to Execute", f"CSV ({chart_id})", 
                                                          "Time Format Error", 
                                                          "Time format error. Correct format should be '%Y/%m/%d %H:%M'.", found_file, source="CSV")
                                            unable_count += 1
                                            has_error_in_row = True
                                            file_format_error_count += 1
                                            csv_status = "Time Format Error"
                                        elif converted_times.isna().any():
                                            # 改為 Unable to Execute：即使部分時間無效也視為嚴重問題
                                            invalid_count = converted_times.isna().sum()
                                            print(f"  [DEBUG] 部分時間解析失敗: {invalid_count}/{len(converted_times)}")
                                            self.emit_log("Unable to Execute", f"CSV ({chart_id})", 
                                                          f"Partial Invalid Times ({invalid_count}/{len(converted_times)})", 
                                                          "Some time values are invalid.", found_file, source="CSV")
                                            unable_count += 1
                                            has_error_in_row = True
                                            file_format_error_count += 1
                                            csv_status = "Partial Invalid Times"
                                        else:
                                            print(f"  [DEBUG] 時間格式檢查通過，所有 {len(converted_times)} 筆時間都有效")
                                            csv_status = "OK"
                                    
                                    except Exception as time_error:
                                        print(f"  [DEBUG] 時間格式檢查出錯: {time_error}")
                                        self.emit_log("Unable to Execute", f"CSV ({chart_id})", 
                                                      f"Time Conversion Error: {str(time_error)}", 
                                                      "Time format processing failed.", found_file, source="CSV")
                                        unable_count += 1
                                        has_error_in_row = True
                                        file_format_error_count += 1
                                        csv_status = "Time Conversion Error"

                        except PermissionError:
                            self.emit_log("Unable to Execute", f"CSV ({chart_id})", 
                                          "Permission denied: File is locked or in use", 
                                          "⚠️ File is locked or in use. Please close this CSV file.", found_file, source="CSV")
                            unable_count += 1
                            has_error_in_row = True
                            file_read_error_count += 1
                            csv_status = "Permission Error"
                        except Exception as e:
                            self.emit_log("Unable to Execute", f"CSV ({chart_id})", f"Read Error: {str(e)}", "File read error, may be corrupted.", found_file, source="CSV")
                            unable_count += 1
                            has_error_in_row = True
                            file_read_error_count += 1
                            csv_status = "Read Error"
                
                # 統計本行結果：只有完全無錯誤且 CSV 檢查通過時才標記 Pass
                if not has_error_in_row and csv_status == "OK":
                    pass_count += 1
                    actual_filename = os.path.basename(found_file) if found_file else "Unknown"
                    self.emit_log("Pass", chart_id, "All checks passed", f"CSV file '{actual_filename}' is ready for processing.", found_file, source="CSV")

                # 每行處理完畢後更新進度 - 優化：降低更新頻率
                update_interval = max(1, total_rows // 100)  # 每1%更新一次，至少每行
                if (i + 1) % update_interval == 0 or i >= total_rows - 5:  # 定期更新或最後幾行
                    try:
                        self.progress_updated.emit(i + 1, total_rows)
                        self.stats_updated.emit(unable_count, warning_count, pass_count, skipped_count)
                    except Exception as ui_error:
                        print(f"[DEBUG] UI更新錯誤 (row {i+1}): {ui_error}")
                
                # 每50行給UI一點時間處理事件，防止凍結
                if (i + 1) % 50 == 0:
                    self.msleep(1)  # 暫停1毫秒讓UI有時間響應
                
                # 最後幾行的詳細除錯資訊
                if i >= total_rows - 5:
                    print(f"  [DEBUG] Row {i+1} 完成: unable={unable_count}, pass={pass_count}, skip={skipped_count}")

            # 迴圈結束後的最終狀態確認
            print(f"[DEBUG] 迴圈完成: 處理了 {i+1}/{total_rows} 行")
            print(f"[DEBUG] 最終統計: unable={unable_count}, pass={pass_count}, skip={skipped_count}")
            
            # 確保最終進度顯示100%
            try:
                self.progress_updated.emit(total_rows, total_rows)
                self.stats_updated.emit(unable_count, warning_count, pass_count, skipped_count)
            except Exception as final_ui_error:
                print(f"[DEBUG] 最終UI更新錯誤: {final_ui_error}")

            # === 智能診斷：當 pass_count == 0 時，判斷根本原因 ===
            if pass_count == 0 and total_rows > 0:
                if file_not_found_count > 0:
                    # 情況 1: 有檔案找不到 → 路徑/檔名問題
                    self.emit_log("Unable to Execute", "🔍 Diagnosis", 
                                  f"{file_not_found_count}/{total_rows} CSV files not found", 
                                  "⚠️ Likely cause: Wrong raw_data_dir path or incorrect file naming. Check 'input' folder location and ensure files start with 'GroupName_ChartName' pattern.", source="System")
                    unable_count += 1
                
                else:
                    # 情況 2: 檔案都找到了，但有格式/讀取問題 → CSV 內容問題
                    if file_format_error_count > 0 or file_read_error_count > 0:
                        self.emit_log("Unable to Execute", "🔍 Diagnosis", 
                                      f"All CSV files found but have errors: {file_format_error_count} format errors, {file_read_error_count} read errors", 
                                      "⚠️ Likely cause: CSV content issue. Ensure 'point_val' and 'point_time' columns exist and time format is '%Y/%m/%d %H:%M'.", source="System")
                        unable_count += 1
                    # 情況 3: Excel 配置問題導致無法檢查 CSV
                    elif excel_logic_error_count > 0:
                        self.emit_log("Unable to Execute", "🔍 Diagnosis", 
                                      f"{excel_logic_error_count} rows have Excel config errors", 
                                      "⚠️ Likely cause: AllChartInfo Excel has missing/invalid values. Fix Excel configuration first before CSV checks can proceed.", source="Excel")
                        unable_count += 1
            
            # 確保最終狀態正確發送
            final_success = unable_count == 0
            print(f"[DEBUG] 準備發送 finished_check 信號: success={final_success}")
            
            try:
                self.finished_check.emit(final_success)
                print(f"[DEBUG] finished_check 信號已發送")
            except Exception as final_error:
                print(f"[DEBUG] finished_check 信號發送失敗: {final_error}")
                # 強制發送失敗狀態
                self.finished_check.emit(False)

        except Exception as e:
            print(f"[DEBUG] 系統發生未預期錯誤: {str(e)}")
            import traceback
            print(f"[DEBUG] 完整錯誤追蹤: {traceback.format_exc()}")
            
            self.emit_log("Unable to Execute", "System", f"Unexpected Crash: {str(e)}", "Contact Developer.", source="System")
            
            try:
                self.finished_check.emit(False)
                print(f"[DEBUG] 錯誤處理完成，finished_check 信號已發送")
            except Exception as final_error:
                print(f"[DEBUG] 最終信號發送失敗: {final_error}")

    def emit_log(self, severity, location, issue, action, csv_path=None, source="Unknown", row_num=None):
        # 批量收集logs，減少信號發送頻率
        log_entry = {
            "Severity": severity,
            "Location": location,
            "Issue": issue,
            "Action": action,
            "CSV_Path": csv_path or "",
            "Source": source,
            "RowNum": row_num
        }
        
        # 對於關鍵錯誤立即發送，對於一般狀態可以稍後批量發送
        if severity in ["Unable to Execute", "Warning"]:
            self.log_added.emit(log_entry)
        else:
            # Pass 和 Skipped 可以批量處理
            self.log_added.emit(log_entry)

# ============================================================================
# 2. UI Widget (嵌入式介面)
# ============================================================================
class DataHealthCheckWidget(QtWidgets.QWidget):
    """
    資料健檢主頁面 (嵌入式 Widget 版本)
    """
    def __init__(self, parent=None):
        super().__init__(parent)
        self.excel_path = ""
        self.raw_data_dir = ""
        self.all_logs = []
        self.pass_logs = []  # 暫存 Pass 日誌，最後批次顯示
        self.error_buffer = []  # 錯誤緩衝區，累積後批次顯示
        self._is_refreshing = False  # 防止重複渲染的旗標
        self._current_row_index = 0  # 追蹤當前表格行數
        self._batch_size = 25  # 每累積25筆錯誤就批次顯示一次
        
        self.init_ui()
        self.apply_styles()
        
        # 註冊為翻譯觀察者
        get_translator().register_observer(self)
    
    def msleep(self, msecs):
        """輔助方法：暫停指定毫秒數"""
        loop = QtCore.QEventLoop()
        QtCore.QTimer.singleShot(msecs, loop.quit)
        loop.exec()

    def update_paths(self, excel_path, raw_data_dir):
        """主程式呼叫此方法來更新路徑"""
        self.excel_path = excel_path
        self.raw_data_dir = raw_data_dir
        self.path_label.setText(f"{tr('checking', 'Checking')}: {os.path.basename(self.excel_path)}")
        self.btn_open_source.setEnabled(True)

    def init_ui(self):
        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(30, 30, 30, 30)
        layout.setSpacing(20)

        # 1. 標題與路徑顯示
        header_layout = QtWidgets.QHBoxLayout()
        title_icon = QtWidgets.QLabel("🩺")
        title_icon.setStyleSheet("font-size: 28px;")
        
        title_info_layout = QtWidgets.QVBoxLayout()
        self.title_label = QtWidgets.QLabel()
        self.title_label.setObjectName("titleLabel")
        self.title_label.setStyleSheet("font-size: 20px; font-weight: bold; color: #1E293B;")
        self.path_label = QtWidgets.QLabel()
        self.path_label.setStyleSheet("color: #64748B; font-size: 13px;")
        title_info_layout.addWidget(self.title_label)
        title_info_layout.addWidget(self.path_label)
        
        header_layout.addWidget(title_icon)
        header_layout.addLayout(title_info_layout)
        header_layout.addStretch()
        layout.addLayout(header_layout)

        # 2. 控制按鈕區 (新增 開啟檔案按鈕 + 篩選器)
        btn_layout = QtWidgets.QHBoxLayout()
        
        self.btn_start = QtWidgets.QPushButton("▶ Start Check")
        self.btn_start.setFixedSize(140, 40)
        self.btn_start.clicked.connect(self.start_check)
        self.btn_start.setCursor(QtCore.Qt.CursorShape.PointingHandCursor)
        
        # [新增] 開啟來源檔案按鈕
        self.btn_open_source = QtWidgets.QPushButton("📂 AllChartInfo Excel")
        self.btn_open_source.setObjectName("btnOpen")
        self.btn_open_source.setFixedSize(200, 40)
        self.btn_open_source.clicked.connect(self.open_source_file)
        self.btn_open_source.setEnabled(False) # 初始為禁用
        self.btn_open_source.setCursor(QtCore.Qt.CursorShape.PointingHandCursor)

        self.btn_export = QtWidgets.QPushButton("📁 Export Report")
        self.btn_export.setObjectName("btnExport")
        self.btn_export.setFixedSize(180, 40)
        self.btn_export.clicked.connect(self.export_report)
        self.btn_export.setEnabled(False)
        self.btn_export.setCursor(QtCore.Qt.CursorShape.PointingHandCursor)
        
        # [新增] 篩選器 Checkbox
        self.chk_filter_errors = QtWidgets.QCheckBox(tr("only_show_errors"))
        self.chk_filter_errors.setStyleSheet("font-size: 13px; font-weight: bold; color: #334155;")
        self.chk_filter_errors.stateChanged.connect(self.apply_filter)
        self.chk_filter_errors.setCursor(QtCore.Qt.CursorShape.PointingHandCursor)
        
        btn_layout.addWidget(self.btn_start)
        btn_layout.addWidget(self.btn_open_source) # 插入中間
        btn_layout.addWidget(self.btn_export)
        btn_layout.addWidget(self.chk_filter_errors)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        # 3. 進度條
        self.progress_bar = QtWidgets.QProgressBar()
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setFixedHeight(8)
        self.progress_bar.setStyleSheet("""
            QProgressBar { background-color: #E2E8F0; border-radius: 4px; border: none; }
            QProgressBar::chunk { background-color: #344CB7; border-radius: 4px; }
        """)
        layout.addWidget(self.progress_bar)

        # 4. 儀表板卡片
        stats_layout = QtWidgets.QHBoxLayout()
        stats_layout.setSpacing(20)
        
        self.card_total = self.create_stat_card("", "0", "#64748B", "total")
        self.card_pass = self.create_stat_card("", "0", "#10B981", "passed")
        self.card_skip = self.create_stat_card("", "0", "#94A3B8", "skipped")
        self.card_unable = self.create_stat_card("", "0", "#EF4444", "unable")
        
        stats_layout.addWidget(self.card_total)
        stats_layout.addWidget(self.card_pass)
        stats_layout.addWidget(self.card_skip)
        stats_layout.addWidget(self.card_unable)
        layout.addLayout(stats_layout)

        # 5. 詳細表格
        self.table = QtWidgets.QTableWidget()
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(["", "", "", "", ""])
        
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QtWidgets.QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(1, QtWidgets.QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(2, QtWidgets.QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(3, QtWidgets.QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(4, QtWidgets.QHeaderView.ResizeMode.Interactive)
        
        # 設定固定寬度給特定欄位
        self.table.setColumnWidth(0, 120)  # Severity
        self.table.setColumnWidth(4, 100)  # Open CSV
        
        self.table.verticalHeader().setVisible(False)
        self.table.setEditTriggers(QtWidgets.QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setAlternatingRowColors(True)
        self.table.setSortingEnabled(True)
        self.table.setWordWrap(True)  # 啟用自動換行
        self.table.verticalHeader().setSectionResizeMode(QtWidgets.QHeaderView.ResizeMode.ResizeToContents)  # 自動調整行高
        
        layout.addWidget(self.table)
        
        # 初始化 UI 文字
        self.refresh_ui_texts()

    def create_stat_card(self, title, value, color_code, key):
            card = QtWidgets.QWidget()
            card.setObjectName("StatCard")
            shadow = QtWidgets.QGraphicsDropShadowEffect()
            shadow.setBlurRadius(15)
            shadow.setColor(QtGui.QColor(0, 0, 0, 20))
            shadow.setOffset(0, 4)
            card.setGraphicsEffect(shadow)
            
            # [修改] 在樣式表中加入 QLabel { background-color: transparent; }
            # 這會強制卡片內的所有文字標籤背景全透明，消除色塊
            card.setStyleSheet(f"""
                #StatCard {{
                    background-color: white;
                    border: 1px solid #E2E8F0;
                    border-radius: 12px;
                }}
                QLabel {{
                    background-color: transparent;
                    border: none;
                }}
            """)
            
            vbox = QtWidgets.QVBoxLayout(card)
            vbox.setContentsMargins(20, 20, 20, 20)
            
            lbl_title = QtWidgets.QLabel(title)
            lbl_title.setObjectName("card_title")
            lbl_title.setStyleSheet("color: #64748B; font-size: 13px; font-weight: 600; text-transform: uppercase;")
            
            lbl_value = QtWidgets.QLabel(value)
            # 注意：這裡原本的樣式只設定了 color，現在會自動繼承上面的 transparent
            lbl_value.setStyleSheet(f"color: {color_code}; font-size: 32px; font-weight: 800;")
            lbl_value.setAlignment(QtCore.Qt.AlignmentFlag.AlignRight)
            
            # 使用 key 參數而不是從 title 解析
            setattr(self, f"lbl_val_{key}", lbl_value)
            setattr(card, 'title_label', lbl_title)  # 儲存標題 label 引用
            setattr(card, 'value_label', lbl_value)  # 儲存數值 label 引用
            
            vbox.addWidget(lbl_title)
            vbox.addWidget(lbl_value)
            return card

    def refresh_ui_texts(self):
        """刷新 UI 中所有可翻譯的文字"""
        # 標題與路徑
        self.title_label.setText(tr("data_health_monitor", "Data Health Monitor"))
        if self.excel_path:
            self.path_label.setText(os.path.basename(self.excel_path))
        else:
            self.path_label.setText(tr("no_file_loaded", "No file loaded"))
        
        # 按鈕
        self.btn_start.setText(tr("start_check", "▶ Start Check"))
        self.btn_open_source.setText(tr("allchartinfo_excel", "📂 AllChartInfo Excel"))
        self.btn_export.setText(tr("export_report", "📁 Export Report"))
        
        # 統計卡片標題
        if hasattr(self.card_total, 'title_label'):
            self.card_total.title_label.setText(tr("total_scanned", "Total Scanned"))
        if hasattr(self.card_pass, 'title_label'):
            self.card_pass.title_label.setText(tr("passed", "Passed"))
        if hasattr(self.card_skip, 'title_label'):
            self.card_skip.title_label.setText(tr("skipped", "Skipped"))
        if hasattr(self.card_unable, 'title_label'):
            self.card_unable.title_label.setText(tr("unable_to_execute"))
        
        # 篩選 checkbox
        if hasattr(self, 'chk_filter_errors'):
            self.chk_filter_errors.setText(tr("only_show_errors"))
        
        # 表格欄位標題
        self.table.setHorizontalHeaderLabels([
            tr("severity", "Severity"),
            tr("location", "Location"),
            tr("issue_description", "Issue Description"),
            tr("suggested_action", "Suggested Action"),
            tr("open_csv", "Open CSV")
        ])
        
        # 重新顯示日誌以更新內容 - 只在語言切換時更新，避免頁面切換時重新渲染
        if self.all_logs and not self._is_refreshing:
            self._is_refreshing = True  # 設定旗標，防止重複觸發
            self.display_sorted_logs()
            self._is_refreshing = False
    
    def translate_log_message(self, text):
        """翻譯日誌訊息 - 嘗試匹配已知的錯誤訊息並翻譯"""
        # 訊息映射表：英文 -> 翻譯鍵
        message_map = {
            # Excel errors
            "GroupName or ChartName is empty": "groupname_chartname_empty",
            "Fill in the names.": "fill_in_names",
            "Missing Target/UCL/LCL": "missing_target_ucl_lcl",
            "These fields are mandatory.": "fields_mandatory",
            "Non-numeric Control Limits": "non_numeric_limits",
            "Ensure limits are numbers.": "ensure_limits_numbers",
            "Invalid Characteristic": "invalid_characteristic",
            "Use Nominal, Smaller, or Bigger.": "use_nominal_smaller_bigger",
            "Nominal requires USL and LSL": "nominal_requires_usl_lsl",
            "Fill both USL and LSL.": "fill_both_usl_lsl",
            "Smaller requires USL": "smaller_requires_usl",
            "Fill USL.": "fill_usl",
            "Bigger requires LSL": "bigger_requires_lsl",
            "Fill LSL.": "fill_lsl",
            "Add missing columns to Excel.": "add_missing_columns",
            "Missing columns": "missing_columns",
            
            # CSV errors
            "File Not Found": "file_not_found_msg",
            "Empty CSV file": "empty_csv_file",
            "CSV has no data rows.": "no_data_rows",
            "No 'point_val' column": "no_point_val_column",
            "Check CSV header.": "check_csv_header",
            "No 'point_time' column": "no_point_time_column",
            "Time Format Error": "time_format_error",
            "Cannot parse as datetime.": "cannot_parse_datetime",
            "Partial Invalid Times": "partial_invalid_times",
            "Read Error": "read_error",
            "File might be corrupted or unreadable.": "file_corrupted",
            "All checks passed": "all_checks_passed",
            "CSV file is ready for processing.": "csv_ready",
            
            # System errors
            "Permission denied: File is locked or in use": "permission_denied",
            "Failed to open Excel": "failed_to_open_excel",
            "Excel file not found": "excel_file_not_found",
        }
        
        # 先檢查完全匹配
        if text in message_map:
            return tr(message_map[text], text)
        
        # 處理包含變數的訊息（如 "Logic: LCL > UCL"）
        if "Logic: LCL" in text and "UCL" in text:
            # 提取數值部分
            import re
            match = re.search(r'Logic: LCL \(([^)]+)\) > UCL \(([^)]+)\)', text)
            if match:
                lcl_val, ucl_val = match.groups()
                base_msg = tr("lcl_greater_ucl", "Logic: LCL > UCL")
                return f"{base_msg} ({lcl_val} > {ucl_val})"
            return tr("lcl_greater_ucl", "Logic: LCL > UCL")
        
        if "LCL must be <= UCL" in text:
            return tr("lcl_must_le_ucl", text)
        
        if "Logic: LSL" in text and "USL" in text:
            return tr("logic_lsl_greater_usl", "Logic: LSL > USL")
        
        if "LSL must be <= USL" in text:
            return tr("lsl_must_le_usl", text)
        
        if "Expected" in text and ".csv" in text:
            # "Expected: GroupName_ChartName.csv"
            parts = text.split("Expected")
            if len(parts) > 1:
                filename = parts[1].strip(": ")
                return f"{tr('expected_csv', 'Expected')}: {filename}"
        
        if "Ensure it is in 'input/raw_charts'" in text:
            return tr("ensure_in_input", text)
        
        if "Some time values cannot be parsed" in text:
            return tr("some_times_invalid", text)
        
        if "Please close the Excel file" in text or "⚠️ Please close" in text:
            return tr("permission_denied_action", text)
        
        if "Please close this CSV file" in text:
            return tr("close_csv_file", text)
        
        # === 處理 Excel 行號相關訊息（支援中英文）===
        import re
        
        # 處理英文訊息：Check Excel row X: ...
        if "Check Excel row" in text and ":" in text:
            # 提取行號
            match = re.search(r'Check Excel row (\d+):', text)
            if match:
                row_num = match.group(1)
                # 判斷具體錯誤類型
                if "GroupName and ChartName are mandatory" in text:
                    return tr("check_excel_row_groupname_chartname", f"Check Excel row {row_num}: GroupName and ChartName are mandatory.").replace("{row}", row_num)
                elif "Target, UCL, LCL are mandatory" in text:
                    return tr("check_excel_row_target_ucl_lcl", f"Check Excel row {row_num}: Target, UCL, LCL are mandatory.").replace("{row}", row_num)
                elif "Control limits must be numeric" in text:
                    return tr("check_excel_row_numeric", f"Check Excel row {row_num}: Control limits must be numeric.").replace("{row}", row_num)
                elif "Characteristics must be Nominal, Smaller, or Bigger" in text:
                    return tr("check_excel_row_characteristics", f"Check Excel row {row_num}: Characteristics must be Nominal, Smaller, or Bigger.").replace("{row}", row_num)
                elif "Nominal type requires both USL and LSL" in text:
                    return tr("check_excel_row_nominal", f"Check Excel row {row_num}: Nominal type requires both USL and LSL.").replace("{row}", row_num)
                elif "Smaller type requires USL" in text and "must satisfy" not in text:
                    return tr("check_excel_row_smaller", f"Check Excel row {row_num}: Smaller type requires USL.").replace("{row}", row_num)
                elif "Bigger type requires LSL" in text and "must satisfy" not in text:
                    return tr("check_excel_row_bigger", f"Check Excel row {row_num}: Bigger type requires LSL.").replace("{row}", row_num)
                elif "Must satisfy USL >= UCL >= Target >= LCL >= LSL" in text:
                    return tr("check_excel_row_logic_nominal", f"Check Excel row {row_num}: Must satisfy USL >= UCL >= Target >= LCL >= LSL.").replace("{row}", row_num)
                elif "Smaller type must satisfy USL >= UCL >= Target >= LCL" in text:
                    return tr("check_excel_row_logic_smaller", f"Check Excel row {row_num}: Smaller type must satisfy USL >= UCL >= Target >= LCL.").replace("{row}", row_num)
                elif "Bigger type must satisfy UCL >= Target >= LCL >= LSL" in text:
                    return tr("check_excel_row_logic_bigger", f"Check Excel row {row_num}: Bigger type must satisfy UCL >= Target >= LCL >= LSL.").replace("{row}", row_num)
        
        # 處理中文訊息（向後兼容）
        if "請檢查 Excel" in text and "行：" in text:
            # 提取行號
            match = re.search(r'第 (\d+) 行', text)
            if match:
                row_num = match.group(1)
                # 判斷具體錯誤類型
                if "GroupName 與 ChartName 為必填項" in text:
                    return tr("check_excel_row_groupname_chartname", f"Check Excel row {row_num}: GroupName and ChartName are mandatory.").replace("{row}", row_num)
                elif "Target、UCL、LCL 為必填項" in text:
                    return tr("check_excel_row_target_ucl_lcl", f"Check Excel row {row_num}: Target, UCL, LCL are mandatory.").replace("{row}", row_num)
                elif "LCL 不得大於 UCL" in text:
                    return tr("check_excel_row_lcl_ucl", f"Check Excel row {row_num}: LCL must not exceed UCL.").replace("{row}", row_num)
                elif "LSL 不得大於 USL" in text:
                    return tr("check_excel_row_lsl_usl", f"Check Excel row {row_num}: LSL must not exceed USL.").replace("{row}", row_num)
                elif "管制界限必須為數值" in text:
                    return tr("check_excel_row_numeric", f"Check Excel row {row_num}: Control limits must be numeric.").replace("{row}", row_num)
                elif "Characteristics 必須為 Nominal、Smaller 或 Bigger" in text:
                    return tr("check_excel_row_characteristics", f"Check Excel row {row_num}: Characteristics must be Nominal, Smaller, or Bigger.").replace("{row}", row_num)
                elif "Nominal 類型需要同時填寫 USL 與 LSL" in text:
                    return tr("check_excel_row_nominal", f"Check Excel row {row_num}: Nominal type requires both USL and LSL.").replace("{row}", row_num)
                elif "Smaller 類型需要填寫 USL" in text:
                    return tr("check_excel_row_smaller", f"Check Excel row {row_num}: Smaller type requires USL.").replace("{row}", row_num)
                elif "Bigger 類型需要填寫 LSL" in text:
                    return tr("check_excel_row_bigger", f"Check Excel row {row_num}: Bigger type requires LSL.").replace("{row}", row_num)
        
        # CSV 相關錯誤
        if "CSV 檔案無資料：" in text:
            return tr("csv_empty_file", "CSV file is empty.")
        if "CSV 檔案缺少 'point_val' 欄位：" in text:
            return tr("csv_missing_point_val", "CSV file is missing 'point_val' column.")
        if "CSV 檔案缺少 'point_time' 欄位：" in text:
            return tr("csv_missing_point_time", "CSV file is missing 'point_time' column.")
        if "時間格式錯誤" in text and "正確格式應為" in text:
            return tr("csv_time_format_error", "Time format error. Correct format should be '%Y/%m/%d %H:%M'.")
        if "部分時間值無效" in text:
            return tr("csv_partial_invalid_times", "Some time values are invalid.")
        if "檔案被鎖定或正在使用中，請關閉此 CSV 檔案：" in text:
            return tr("csv_permission_denied", "⚠️ File is locked or in use. Please close this CSV file.")
        if "檔案讀取錯誤，可能已損毀：" in text:
            return tr("csv_read_error", "File read error, may be corrupted.")
        
        # 如果沒有匹配，返回原文
        return text

    def apply_styles(self):
        self.setStyleSheet("""
            QWidget { font-family: 'Microsoft JhengHei', sans-serif; }
            #titleLabel { font-size: 26px; font-weight: 800; color: #1E293B; }
            
            QPushButton {
                background-color: #344CB7;
                color: white;
                border-radius: 8px;
                font-weight: bold;
                font-size: 14px;
                border: none;
            }
            QPushButton:hover { background-color: #4D64D6; }
            QPushButton:disabled { background-color: #CBD5E1; color: #94A3B8; }
            
            /* 特殊按鈕樣式 */
            #btnOpen {
                background-color: #10B981; /* 綠色 */
            }
            #btnOpen:hover { background-color: #059669; }
            #btnOpen:disabled { background-color: #D1FAE5; color: #A7F3D0; }

            #btnExport {
                background-color: white;
                color: #344CB7;
                border: 1.5px solid #344CB7;
            }
            #btnExport:hover { background-color: #F0F4FF; }

            QGroupBox#detailsGroup {
                background-color: white;
                border: 1px solid #E2E8F0;
                border-radius: 12px;
                margin-top: 20px;
                padding-top: 25px;
                font-weight: 700;
                font-size: 14px;
                color: #1E293B;
            }
            QGroupBox#detailsGroup::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 8px 16px;
                background-color: #F8FAFC;
                border: 1px solid #E2E8F0;
                border-radius: 6px;
                color: #475569;
                font-weight: 600;
                font-size: 13px;
                left: 20px;
                top: 10px;
            }
            QTableWidget { border: none; background-color: white; gridline-color: #F1F5F9; }
            QHeaderView::section {
                background-color: #F1F5F9;
                color: #64748B;
                border: none;
                border-bottom: 2px solid #E2E8F0;
                padding: 12px;
                font-weight: bold;
                text-align: left;
            }
        """)

    def apply_filter(self):
        """根據 checkbox 狀態篩選表格顯示"""
        show_errors_only = self.chk_filter_errors.isChecked()
        
        for row in range(self.table.rowCount()):
            severity_item = self.table.item(row, 0)
            if severity_item:
                severity_text = severity_item.text()
                # 檢查是否包含 "Pass" 或 "✅"
                is_pass = "Pass" in severity_text or "✅" in severity_text
                # 如果勾選「只顯示錯誤」，就隱藏 Pass 行
                self.table.setRowHidden(row, show_errors_only and is_pass)
    
    def start_check(self):
        if not self.excel_path or not os.path.exists(self.excel_path):
            QtWidgets.QMessageBox.warning(self, tr("error", "Error"), tr("path_not_set", "Path not set properly."))
            return

        self.btn_start.setEnabled(False)
        self.btn_export.setEnabled(False)
        self.btn_open_source.setEnabled(False) # 開始時禁用
        self.table.setRowCount(0)
        self.all_logs = []
        self.pass_logs = []  # 清空暫存的 Pass 日誌
        self.error_buffer = []  # 清空錯誤緩衝區
        self._current_row_index = 0  # 重置行索引
        
        self.lbl_val_total.setText("0")
        self.lbl_val_unable.setText("0")
        self.lbl_val_skipped.setText("0")
        self.lbl_val_passed.setText("0")
        
        # 暫停排序功能，檢查期間禁用
        self.table.setSortingEnabled(False)

        self.worker = DataValidatorWorker(self.excel_path, self.raw_data_dir)
        self.worker.progress_updated.connect(self.update_progress)
        self.worker.stats_updated.connect(self.update_stats)
        self.worker.log_added.connect(self.add_log_entry)
        self.worker.finished_check.connect(self.on_check_finished)
        self.worker.start()

    def update_progress(self, current, total):
        self.progress_bar.setMaximum(total)
        self.progress_bar.setValue(current)
        self.lbl_val_total.setText(f"{current} / {total}")

    def update_stats(self, unable, warning, passed, skipped):
        self.lbl_val_unable.setText(str(unable))
        self.lbl_val_skipped.setText(str(skipped))
        self.lbl_val_passed.setText(str(passed))

    def add_log_entry(self, log):
        """收集 log 記錄，錯誤使用小批次緩衝顯示，Pass 最後批次顯示"""
        self.all_logs.append(log)
        
        severity = log.get('Severity', '')
        
        # 錯誤和警告加入緩衝區
        if severity in ['Unable to Execute', 'Warning', 'Skipped']:
            self.error_buffer.append(log)
            
            # 累積到批次大小時，批次顯示
            if len(self.error_buffer) >= self._batch_size:
                self._flush_error_buffer()
        elif severity == 'Pass':
            # Pass 日誌暫存，最後批次顯示
            self.pass_logs.append(log)
    
    def _flush_error_buffer(self):
        """批次顯示緩衝區中的錯誤日誌"""
        if not self.error_buffer:
            return
        
        # 暫停UI更新
        self.table.setUpdatesEnabled(False)
        
        # 一次性分配行數
        start_row = self._current_row_index
        self.table.setRowCount(start_row + len(self.error_buffer))
        
        # 批次渲染
        for i, log in enumerate(self.error_buffer):
            self._add_log_to_table_optimized(log, start_row + i)
        
        # 更新行索引
        self._current_row_index += len(self.error_buffer)
        
        # 恢復UI更新（觸發一次重繪）
        self.table.setUpdatesEnabled(True)
        
        # 清空緩衝區
        self.error_buffer.clear()
        
        # 讓UI有時間響應
        QtWidgets.QApplication.processEvents()
    
    def display_sorted_logs(self, force_refresh=False):
        """按照優先級排序並顯示所有 logs - 針對大量資料優化
        
        Args:
            force_refresh: 強制刷新表格（用於檢查完成或語言切換），否則只在需要時更新
        """
        # 如果正在刷新且非強制，跳過
        if self._is_refreshing and not force_refresh:
            return
            
        # 定義排序優先級
        severity_order = {
            'Unable to Execute': 1,
            'Warning': 2,
            'Skipped': 3,
            'Pass': 4,
            'Info': 5
        }
        
        # 過濾掉 INFO 級別，只保留其他級別
        filtered_logs = [log for log in self.all_logs if log['Severity'] != 'Info']
        
        # 排序 logs
        sorted_logs = sorted(filtered_logs, key=lambda x: severity_order.get(x['Severity'], 999))
        
        # [優化] 大量資料時顯示處理進度
        total_logs = len(sorted_logs)
        if total_logs > 100:
            print(f"[DEBUG] 正在顯示 {total_logs} 筆日誌資料...")
        
        # [關鍵優化] 暫停所有UI更新，提高性能
        self.table.setUpdatesEnabled(False)
        self.table.setSortingEnabled(False)
        
        # 清空表格並一次性設定行數
        self.table.setRowCount(0)
        self.table.setRowCount(total_logs)  # 一次性分配所有行
        
        # [優化] 更小的批次處理，每25筆為一批
        batch_size = 25
        for batch_start in range(0, total_logs, batch_size):
            batch_end = min(batch_start + batch_size, total_logs)
            
            # 顯示進度
            if total_logs > 100:
                progress = int((batch_start / total_logs) * 100)
                print(f"[DEBUG] 日誌顯示進度: {progress}% ({batch_start}/{total_logs})")
            
            # 處理這一批logs
            for i in range(batch_start, batch_end):
                log = sorted_logs[i]
                self._add_log_to_table_optimized(log, i)
            
            # 每處理一小批就讓UI有機會更新
            if total_logs > 50:
                QtWidgets.QApplication.processEvents()
                
                # 每200筆強制短暫暫停，防止UI凍結
                if (batch_start + batch_size) % 200 == 0:
                    self.msleep(5)  # 短暫暫停5毫秒
        
        # 重新啟用UI更新和排序
        self.table.setUpdatesEnabled(True)
        self.table.setSortingEnabled(True)
        
        # 應用篩選器
        self.apply_filter()
        
        if total_logs > 100:
            print(f"[DEBUG] 日誌顯示完成，共 {total_logs} 筆")
    
    def _add_log_to_table_optimized(self, log, row_index):
        """優化版本：直接設定表格項目，不插入行"""
    def _add_log_to_table(self, log):
        """將單一log加入表格的輔助方法"""
    def _add_log_to_table_optimized(self, log, row_index):
        """優化版本：直接設定表格項目，不插入行"""
        sev_text = log['Severity']
        # 定義不同等級的圖示與顏色
        if sev_text == "Unable to Execute":
            icon, color_code = "🔴", "#EF4444"
        elif sev_text == "Warning":
            icon, color_code = "⚠️", "#F59E0B"
        elif sev_text == "Pass":
            icon, color_code = "✅", "#10B981"
        elif sev_text == "Skipped":
            icon, color_code = "⏭️", "#94A3B8"
        elif sev_text == "Info":
            icon, color_code = "ℹ️", "#3B82F6"
        else:
            icon, color_code = "ℹ️", "#3B82F6"
        
        # [優化] 創建所有項目但延遲設定格式
        item_sev = QtWidgets.QTableWidgetItem(f"{icon} {sev_text}")
        item_sev.setForeground(QtGui.QColor(color_code))
        item_sev.setTextAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        
        item_location = QtWidgets.QTableWidgetItem(str(log['Location']))
        item_location.setTextAlignment(QtCore.Qt.AlignmentFlag.AlignCenter | QtCore.Qt.AlignmentFlag.AlignVCenter)
        
        # 翻譯 Issue 和 Action 訊息（只針對錯誤訊息翻譯）
        if sev_text in ["Unable to Execute", "Warning"]:
            translated_issue = self.translate_log_message(str(log['Issue']))
            translated_action = self.translate_log_message(str(log['Action']))
        else:
            # Pass 和 Skipped 訊息不翻譯，提高性能
            translated_issue = str(log['Issue'])
            translated_action = str(log['Action'])
        
        item_issue = QtWidgets.QTableWidgetItem(translated_issue)
        item_issue.setTextAlignment(QtCore.Qt.AlignmentFlag.AlignCenter | QtCore.Qt.AlignmentFlag.AlignVCenter)
        
        item_action = QtWidgets.QTableWidgetItem(translated_action)
        item_action.setTextAlignment(QtCore.Qt.AlignmentFlag.AlignCenter | QtCore.Qt.AlignmentFlag.AlignVCenter)
        
        # 直接設定到指定行，不插入
        self.table.setItem(row_index, 0, item_sev)
        self.table.setItem(row_index, 1, item_location)
        self.table.setItem(row_index, 2, item_issue)
        self.table.setItem(row_index, 3, item_action)
        
        # [優化] 簡化按鈕創建邏輯
        source = log.get('Source', 'Unknown')
        csv_path = log.get('CSV_Path', '')
        
        if source == "Excel":
            # Excel 錯誤：開啟 AllChartInfo Excel
            btn_open = QtWidgets.QPushButton("📂 Excel")  # 恢復文字
            btn_open.setToolTip("Open AllChartInfo Excel file")  # 詳細說明
            btn_open.setFixedSize(80, 32)  # 增加高度從25到32
            btn_open.setStyleSheet("""
                QPushButton { 
                    background-color: #10B981; 
                    color: white; 
                    border: none; 
                    border-radius: 6px;
                    font-size: 11px;
                    font-weight: bold;
                    padding: 3px 6px;
                }
                QPushButton:hover {
                    background-color: #059669;
                }
            """)
            btn_open.setCursor(QtCore.Qt.CursorShape.PointingHandCursor)
            btn_open.clicked.connect(lambda: self.open_source_file())
            self.table.setCellWidget(row_index, 4, btn_open)
        elif source == "CSV" and csv_path:
            # CSV 錯誤：開啟對應 CSV 檔案 - 修復路徑問題
            if os.path.exists(csv_path):
                btn_open = QtWidgets.QPushButton("📁 CSV")  # 恢復文字
                btn_open.setToolTip(f"Open: {os.path.basename(csv_path)}")  # 顯示檔案名
                btn_open.setFixedSize(80, 32)  # 增加高度從25到32
                btn_open.setStyleSheet("""
                    QPushButton { 
                        background-color: #3B82F6; 
                        color: white; 
                        border: none; 
                        border-radius: 6px;
                        font-size: 11px;
                        font-weight: bold;
                        padding: 3px 6px;
                    }
                    QPushButton:hover {
                        background-color: #2563EB;
                    }
                """)
                btn_open.setCursor(QtCore.Qt.CursorShape.PointingHandCursor)
                # 修復lambda閉包問題 - 創建本地變數
                file_path = str(csv_path)  # 確保是字符串
                btn_open.clicked.connect(lambda checked, path=file_path: self.open_csv_file(path))
                self.table.setCellWidget(row_index, 4, btn_open)
            else:
                # CSV路徑無效時顯示錯誤
                btn_error = QtWidgets.QPushButton("❌ Missing")
                btn_error.setToolTip(f"File not found: {csv_path}")
                btn_error.setFixedSize(80, 32)  # 保持一致的高度
                btn_error.setStyleSheet("QPushButton { background-color: #EF4444; color: white; border: none; border-radius: 6px; font-size: 11px; padding: 3px 6px; }")
                btn_error.setEnabled(False)
                self.table.setCellWidget(row_index, 4, btn_error)
        else:
            item_na = QtWidgets.QTableWidgetItem(tr("n_a", "N/A"))
            item_na.setTextAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
            self.table.setItem(row_index, 4, item_na)
        
        # 設定行高以容納按鈕
        self.table.setRowHeight(row_index, 36)  # 稍微比按鈕高一點

    def on_check_finished(self, passed):
        self.progress_bar.setValue(self.progress_bar.maximum())
        
        # [優化] 先清空剩餘的錯誤緩衝區
        if self.error_buffer:
            print(f"[DEBUG] 清空剩餘 {len(self.error_buffer)} 筆錯誤緩衝")
            self._flush_error_buffer()
        
        # [優化] 批次顯示所有 Pass 日誌 - 使用高效能方法
        if self.pass_logs:
            total_pass = len(self.pass_logs)
            print(f"[DEBUG] 開始批次顯示 {total_pass} 筆 Pass 日誌")
            
            # [關鍵優化] 暫停UI更新，大幅提升性能
            self.table.setUpdatesEnabled(False)
            
            # 一次性分配所有行（不使用insertRow，避免重複觸發佈局）
            pass_start_row = self._current_row_index
            self.table.setRowCount(pass_start_row + total_pass)
            
            # 批次渲染，每100筆更新進度
            for i, log in enumerate(self.pass_logs):
                self._add_log_to_table_optimized(log, pass_start_row + i)
                
                # 每100筆顯示進度
                if (i + 1) % 100 == 0:
                    progress = int((i + 1) / total_pass * 100)
                    print(f"[DEBUG] Pass 日誌顯示進度: {progress}% ({i + 1}/{total_pass})")
            
            # 恢復UI更新
            self.table.setUpdatesEnabled(True)
            print(f"[DEBUG] Pass 日誌顯示完成，共 {total_pass} 筆")
        
        # 重新啟用排序功能
        self.table.setSortingEnabled(True)
        
        # 應用篩選器
        self.apply_filter()
        
        # [修改] 不論 passed 為 True/False，都啟用這些按鈕
        self.btn_start.setEnabled(True)
        self.btn_export.setEnabled(True)
        self.btn_open_source.setEnabled(True)

    def open_source_file(self):
        """開啟目前正在檢查的 Excel 檔案"""
        if self.excel_path and os.path.exists(self.excel_path):
            try:
                os.startfile(self.excel_path)
            except AttributeError:
                # 兼容非 Windows 系統
                import subprocess
                if sys.platform == 'darwin':
                    subprocess.call(('open', self.excel_path))
                else:
                    subprocess.call(('xdg-open', self.excel_path))
        else:
            QtWidgets.QMessageBox.warning(self, tr("error", "Error"), tr("file_not_found", "File not found."))
    
    def open_csv_file(self, csv_path):
        """開啟指定的 CSV 檔案 - 加強錯誤處理"""
        print(f"[DEBUG] 嘗試開啟CSV檔案: {csv_path}")
        
        if not csv_path:
            QtWidgets.QMessageBox.warning(self, tr("error", "Error"), tr("csv_file_not_found", "CSV file path is empty"))
            return
            
        # 轉換為絕對路徑
        if not os.path.isabs(csv_path):
            csv_path = os.path.abspath(csv_path)
            print(f"[DEBUG] 轉換為絕對路徑: {csv_path}")
        
        if csv_path and os.path.exists(csv_path):
            try:
                print(f"[DEBUG] 檔案存在，正在開啟: {csv_path}")
                os.startfile(csv_path)
            except AttributeError:
                # 兼容非 Windows 系統
                import subprocess
                if sys.platform == 'darwin':
                    subprocess.call(('open', csv_path))
                else:
                    subprocess.call(('xdg-open', csv_path))
            except Exception as e:
                print(f"[DEBUG] 開啟檔案時發生錯誤: {e}")
                QtWidgets.QMessageBox.critical(self, tr("error", "Error"), f"無法開啟檔案: {str(e)}")
        else:
            error_msg = tr("csv_file_not_found", "CSV file not found") + f":\n{csv_path}"
            print(f"[DEBUG] 檔案不存在: {csv_path}")
            QtWidgets.QMessageBox.warning(self, tr("error", "Error"), error_msg)

    def export_report(self):
        path, _ = QtWidgets.QFileDialog.getSaveFileName(self, tr("export_log", "Export Log"), "Data_Health_Log.xlsx", "Excel Files (*.xlsx)")
        if path:
            try:
                df = pd.DataFrame(self.all_logs)
                df.to_excel(path, index=False)
                QtWidgets.QMessageBox.information(self, tr("export_log", "Export Log"), tr("export_success", "Report saved to") + f":\n{path}")
            except PermissionError:
                QtWidgets.QMessageBox.critical(self, tr("export_failed", "Export Failed"), 
                    tr("permission_denied_export", "⚠️ Permission denied: Cannot write to file\n\nThe file might be opened in Excel or another program.\nPlease close the file and try again.") + f"\n\nFile: {path}")
            except Exception as e:
                QtWidgets.QMessageBox.critical(self, tr("export_failed", "Export Failed"), f"Failed to export: {e}")

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    win = DataHealthCheckWidget()
    
    # 測試路徑設定
    test_excel_path = "C:/Users/hsa00/Desktop/OOB/OOB_NGK/input/TestData.xlsx"
    test_raw_data_dir = "C:/Users/hsa00/Desktop/OOB/OOB_NGK/input"  # 改為 input 而非 input/raw_charts
    
    if not os.path.exists(test_excel_path):
        test_excel_path, _ = QtWidgets.QFileDialog.getOpenFileName(None, "Select Excel", "", "Excel (*.xlsx)")
        if test_excel_path:
            test_raw_data_dir = os.path.dirname(test_excel_path)
        else:
            sys.exit(0)
    
    win.update_paths(test_excel_path, test_raw_data_dir)
    win.show()
    sys.exit(app.exec())