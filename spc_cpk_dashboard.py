import os
import sys
import pandas as pd
import numpy as np
from PyQt6 import QtWidgets, QtCore, QtGui
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import oob_module_NGK_nostatic as oob_module

class SlidingToggleSwitch(QtWidgets.QAbstractButton):
    """iOS 風格滑動開關"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setCheckable(True)
        self.setMinimumWidth(66)
        self.setMinimumHeight(32)
        self.setCursor(QtCore.Qt.CursorShape.PointingHandCursor)
        
    def paintEvent(self, event):
        painter = QtGui.QPainter(self)
        painter.setRenderHint(QtGui.QPainter.RenderHint.Antialiasing)
        
        # 背景軌道
        track_color = QtGui.QColor("#2563eb") if self.isChecked() else QtGui.QColor("#d1d5db")
        painter.setBrush(track_color)
        painter.setPen(QtCore.Qt.PenStyle.NoPen)
        painter.drawRoundedRect(0, 0, self.width(), self.height(), self.height() / 2, self.height() / 2)
        
        # 滑塊
        thumb_radius = self.height() - 6
        thumb_x = self.width() - thumb_radius - 3 if self.isChecked() else 3
        painter.setBrush(QtGui.QColor("#ffffff"))
        painter.drawEllipse(thumb_x, 3, thumb_radius, thumb_radius)
        
    def sizeHint(self):
        return QtCore.QSize(66, 32)

class DateSettingsDialog(QtWidgets.QDialog):
    """日期設定對話框"""
    def __init__(self, parent=None, current_mode=False, start_date=None, end_date=None):
        super().__init__(parent)
        from translations import tr
        self.setWindowTitle(tr('settings'))
        self.setModal(True)
        self.resize(450, 250)
        
        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(18)
        
        # 標題
        title = QtWidgets.QLabel(tr('calculation_mode_settings'))
        title.setStyleSheet("font-size: 16px; font-weight: bold; color: #1f2937;")
        layout.addWidget(title)
        
        # 第一行：滑動開關 + 標籤
        toggle_row = QtWidgets.QHBoxLayout()
        self.toggle_switch = SlidingToggleSwitch()
        self.toggle_switch.setChecked(current_mode)
        toggle_label = QtWidgets.QLabel(tr('custom_calculation_mode'))
        toggle_label.setStyleSheet("font-size: 14px; color: #374151;")
        toggle_row.addWidget(toggle_label)
        toggle_row.addWidget(self.toggle_switch)
        toggle_row.addStretch()
        layout.addLayout(toggle_row)
        
        # 說明文字
        self.mode_hint = QtWidgets.QLabel()
        self.mode_hint.setWordWrap(True)
        self.mode_hint.setStyleSheet("font-size: 12px; color: #6b7280; padding: 8px; background: #f3f4f6; border-radius: 6px;")
        layout.addWidget(self.mode_hint)
        
        # 第二、三行：日期選擇器
        date_layout = QtWidgets.QFormLayout()
        date_layout.setSpacing(12)
        
        self.start_date = QtWidgets.QDateEdit()
        self.start_date.setCalendarPopup(True)
        self.start_date.setDate(start_date if start_date else QtCore.QDate.currentDate().addMonths(-3))
        self.start_date.setFixedHeight(36)
        
        self.end_date = QtWidgets.QDateEdit()
        self.end_date.setCalendarPopup(True)
        self.end_date.setDate(end_date if end_date else QtCore.QDate.currentDate())
        self.end_date.setFixedHeight(36)
        
        date_layout.addRow(tr('start_date') + ":", self.start_date)
        date_layout.addRow(tr('end_date') + ":", self.end_date)
        layout.addLayout(date_layout)
        
        layout.addStretch()
        
        # 按鈕
        button_box = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.StandardButton.Ok | 
            QtWidgets.QDialogButtonBox.StandardButton.Cancel
        )
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
        
        # 連接開關事件
        self.toggle_switch.toggled.connect(self._on_mode_changed)
        self._on_mode_changed(current_mode)
        
    def _on_mode_changed(self, checked):
        from translations import tr
        if checked:
            # 自訂模式
            self.start_date.setEnabled(True)
            self.end_date.setEnabled(True)
            self.mode_hint.setText(tr('custom_mode_hint'))
        else:
            # 自動偵測模式
            self.start_date.setEnabled(False)
            self.end_date.setEnabled(False)
            self.mode_hint.setText(tr('auto_mode_hint'))
    
    def get_settings(self):
        """返回 (is_custom_mode, start_date, end_date)"""
        return (
            self.toggle_switch.isChecked(),
            self.start_date.date(),
            self.end_date.date()
        )

def calculate_cpk(raw_df, chart_info):
    print(raw_df)
    print(f"[DEBUG] chart_info type: {type(chart_info)}")
    print(f"[DEBUG] chart_info keys: {list(chart_info.keys()) if hasattr(chart_info, 'keys') else chart_info}")
    print(f"[DEBUG] chart_info: {chart_info}")
    
    # 檢查數據是否足夠
    if raw_df is None or raw_df.empty or len(raw_df) < 2:
        print(f"[DEBUG] 數據不足，無法計算 Cpk")
        return {'Cpk': None}
    
    mean = raw_df['point_val'].mean()
    print(f"[DEBUG] mean: {mean}")
    std = raw_df['point_val'].std()
    print(f"[DEBUG] std: {std}")
    
    # 引入極小閾值，避免除以接近0的數
    # 使用較大的閾值以避免計算出不合理的超大 Cpk 值
    epsilon = 1e-6
    if std < epsilon:
        print(f"[DEBUG] 標準差過小 (std={std} < {epsilon})，數據為定值或變異極小，無法計算 Cpk")
        return {'Cpk': None}
    
    characteristic = chart_info['Characteristics']
    usl = chart_info.get('USL', None)
    print(f"[DEBUG] usl: {usl}")
    lsl = chart_info.get('LSL', None)
    print(f"[DEBUG] lsl: {lsl}")
    print(f"[DEBUG] usl: {usl}, lsl: {lsl}, characteristic: {characteristic}")
    cpk = None
    if std > epsilon:  # 再次確認
        characteristic_lower = str(characteristic).lower()
        if characteristic_lower == 'nominal':
            if usl is not None and lsl is not None:
                cpu = (usl - mean) / (3 * std)
                cpl = (mean - lsl) / (3 * std)
                cpk = min(cpu, cpl)
        elif characteristic_lower == 'smaller':
            if usl is not None:
                cpk = (usl - mean) / (3 * std)
        elif characteristic_lower == 'bigger':
            if lsl is not None:
                cpk = (mean - lsl) / (3 * std)
    if cpk is not None:
        cpk = round(cpk, 3)
    return {'Cpk': cpk}

def get_app_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(__file__)

class SPCCpkDashboard(QtWidgets.QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        from translations import tr, TranslationManager
        self.setWindowTitle(tr('spc_cpk_dashboard'))
        self.resize(1200, 800)
        # 資料結構初始化
        self.all_charts_info = None
        self.raw_charts_dict = {}
        self.cpk_results = {}  # {(group_name, chart_name): {'Cpk': value}}
        self.chart_date_states = {}  # 每張圖的日期狀態：{'custom': bool, 'start': date, 'end': date}
        self.axis_mode = 'index'  # 'index' (等距) 或 'time'
        
        # 計算模式設定
        self.is_custom_mode = False  # False: 自動偵測, True: 自訂計算區間
        self.custom_start_date = QtCore.QDate.currentDate().addMonths(-3)
        self.custom_end_date = QtCore.QDate.currentDate()
        
        # 新增子圖控制物件
        self.fig_sub1 = Figure(figsize=(4, 4))
        self.canvas_sub1 = FigureCanvas(self.fig_sub1)
        self.fig_sub2 = Figure(figsize=(4, 4))
        self.canvas_sub2 = FigureCanvas(self.fig_sub2)
        self.fig_sub3 = Figure(figsize=(4, 4))
        self.canvas_sub3 = FigureCanvas(self.fig_sub3)
        
        self.init_ui()
        TranslationManager().register_observer(self)
    
    def refresh_ui_texts(self):
        """刷新UI文字（當語言切換時）"""
        from translations import tr
        
        self.setWindowTitle(tr('spc_cpk_dashboard'))
        self.recalc_btn.setText("▶ " + tr('run_analysis'))
        self.export_excel_btn.setText("⬇ " + tr('download_cpk_detail'))
        self.lbl_chart.setText(tr('chart'))
        self.settings_btn.setText("⚙ " + tr('settings'))
        
        # 更新 metric cards 標題
        self.metric_cards['cpk']['title_label'].setText(tr('cpk'))
        self.metric_cards['l1']['title_label'].setText(tr('l1_cpk'))
        self.metric_cards['l2']['title_label'].setText(tr('l2_cpk'))
        self.metric_cards['custom']['title_label'].setText(tr('long_term_cpk'))
        self.metric_cards['r1']['title_label'].setText(tr('r1'))
        self.metric_cards['r2']['title_label'].setText(tr('r2'))
        self.metric_cards['kval']['title_label'].setText(tr('k'))
        
        # 更新圖表標題和按鈕
        self.title_lbl.setText(tr('spc_chart'))
        self.prev_chart_btn.setText(tr('prev'))
        self.next_chart_btn.setText(tr('next'))
    
    def open_settings_dialog(self):
        """打開設定對話框"""
        dialog = DateSettingsDialog(
            self, 
            self.is_custom_mode, 
            self.custom_start_date, 
            self.custom_end_date
        )
        if dialog.exec() == QtWidgets.QDialog.DialogCode.Accepted:
            is_custom, start_date, end_date = dialog.get_settings()
            self.is_custom_mode = is_custom
            self.custom_start_date = start_date
            self.custom_end_date = end_date
            # 重新計算當前圖表
            idx = self.chart_combo.currentIndex()
            if idx >= 0 and self.all_charts_info is not None and idx < len(self.all_charts_info):
                chart_info = self.all_charts_info.iloc[idx]
                self._update_current_chart_dynamic(chart_info)
    def load_all_chart_data(self):
        # 已廢棄，邏輯移到 recalculate
        pass

    def init_ui(self):
        # 重新打造為 Dashboard 版型
        from translations import tr
        
        root = QtWidgets.QVBoxLayout(self)
        root.setContentsMargins(20, 16, 20, 16)
        root.setSpacing(14)
        # ===== Top Filter / Action Bar =====
        top_bar = QtWidgets.QHBoxLayout()
        top_bar.setSpacing(12)
        self.chart_combo = QtWidgets.QComboBox()
        self.chart_combo.setMinimumWidth(280)
        self.chart_combo.setFixedHeight(34)
        # 只保留執行分析按鈕，並重新設計
        self.recalc_btn = QtWidgets.QPushButton("▶ " + tr('run_analysis'))
        self.recalc_btn.setFixedHeight(40)
        self.recalc_btn.setFixedWidth(140)
        self.recalc_btn.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #2563eb, stop:1 #1d4fd8);
                color: #fff;
                border: none;
                border-radius: 8px;
                font-size: 14px;
                font-weight: bold;
                padding: 0px;
                min-width: 140px;
                max-width: 140px;
                min-height: 40px;
                max-height: 40px;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #1d4fd8, stop:1 #2563eb);
            }
            QPushButton:pressed {
                background: #163fae;
            }
        """)
        # 新增下載 Excel 按鈕
        self.export_excel_btn = QtWidgets.QPushButton("⬇ " + tr('download_cpk_detail'))
        self.export_excel_btn.setFixedHeight(40)
        self.export_excel_btn.setFixedWidth(180)
        self.export_excel_btn.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #2563eb, stop:1 #1d4fd8);
                color: #fff;
                border: none;
                border-radius: 8px;
                font-size: 14px;
                font-weight: bold;
                padding: 0px;
                min-width: 180px;
                max-width: 180px;
                min-height: 40px;
                max-height: 40px;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #1d4fd8, stop:1 #2563eb);
            }
            QPushButton:pressed {
                background: #163fae;
            }
        """)
        self.lbl_chart = QtWidgets.QLabel(tr('chart'))
        self.lbl_chart.setObjectName("plainLabel")
        
        # 新增設定按鈕
        self.settings_btn = QtWidgets.QPushButton("⚙ " + tr('settings'))
        self.settings_btn.setFixedHeight(40)
        self.settings_btn.setFixedWidth(120)
        self.settings_btn.setStyleSheet("""
            QPushButton {
                background: #6b7280;
                color: #fff;
                border: none;
                border-radius: 8px;
                font-size: 13px;
                font-weight: 600;
                padding: 0px;
            }
            QPushButton:hover {
                background: #4b5563;
            }
            QPushButton:pressed {
                background: #374151;
            }
        """)
        
        top_bar.addWidget(self.lbl_chart)
        top_bar.addWidget(self.chart_combo)
        top_bar.addWidget(self.settings_btn)
        top_bar.addWidget(self.recalc_btn)
        top_bar.addWidget(self.export_excel_btn)
        top_bar.addStretch(1)
        root.addLayout(top_bar)
        # ===== Metric Cards Row =====
        self.metric_cards = {}
        cards_layout = QtWidgets.QGridLayout()
        cards_layout.setHorizontalSpacing(12)  # 減少水平間距
        cards_layout.setVerticalSpacing(10)  # 減少垂直間距
        def create_metric_card(key, title, col, row=0):
            frame = QtWidgets.QFrame()
            frame.setObjectName("metricCard")
            frame.setProperty("status", "neutral")
            pal = frame.palette()
            pal.setColor(QtGui.QPalette.ColorRole.Window, QtGui.QColor("#ffffff"))
            frame.setAutoFillBackground(True)
            frame.setPalette(pal)
            layout = QtWidgets.QVBoxLayout(frame)
            layout.setContentsMargins(12, 8, 12, 8)  # 減少內邊距
            layout.setSpacing(2)  # 減少間距
            title_label = QtWidgets.QLabel(title)
            title_label.setObjectName("metricTitle")
            title_label.setAutoFillBackground(True)
            tpal = title_label.palette()
            tpal.setColor(QtGui.QPalette.ColorRole.Window, QtGui.QColor("#ffffff"))
            title_label.setPalette(tpal)
            value_label = QtWidgets.QLabel("-")
            value_label.setObjectName("metricValue")
            value_label.setAlignment(QtCore.Qt.AlignmentFlag.AlignLeft | QtCore.Qt.AlignmentFlag.AlignVCenter)
            value_label.setAutoFillBackground(True)
            vpal = value_label.palette()
            vpal.setColor(QtGui.QPalette.ColorRole.Window, QtGui.QColor("#ffffff"))
            value_label.setPalette(vpal)
            layout.addWidget(title_label)
            layout.addWidget(value_label)
            layout.addStretch(1)
            cards_layout.addWidget(frame, row, col)
            self.metric_cards[key] = {"frame": frame, "value_label": value_label, "title_label": title_label}
        create_metric_card("cpk", tr('cpk'), 0)
        create_metric_card("l1", tr('l1_cpk'), 1)
        create_metric_card("l2", tr('l2_cpk'), 2)
        create_metric_card("custom", tr('long_term_cpk'), 3)
        create_metric_card("r1", tr('r1'), 4)
        create_metric_card("r2", tr('r2'), 5)
        create_metric_card("kval", tr('k'), 6)
        root.addLayout(cards_layout)
        # ===== Chart Area =====
        self.chart_frame = QtWidgets.QFrame()
        self.chart_frame.setObjectName("chartFrame")
        main_vbox = QtWidgets.QVBoxLayout(self.chart_frame)
        # 壓縮內部邊距以節省垂直空間（左, 上, 右, 下）
        main_vbox.setContentsMargins(10, 5, 10, 10)
        main_vbox.setSpacing(5)  # 縮小元素間間距
        
        # 1. 頂部標題與按鈕
        header = QtWidgets.QHBoxLayout()
        self.title_lbl = QtWidgets.QLabel(tr('spc_chart'))
        self.title_lbl.setObjectName("sectionTitle")
        header.addWidget(self.title_lbl)
        header.addStretch(1)
        
        # 加上切換按鈕
        self.prev_chart_btn = QtWidgets.QPushButton(tr('prev'))
        self.prev_chart_btn.setFixedHeight(34)
        self.prev_chart_btn.setFixedWidth(80)
        self.prev_chart_btn.setStyleSheet("""
            QPushButton {
                background: #6b7280;
                color: #fff;
                border: none;
                border-radius: 6px;
                font-size: 12px;
                font-weight: 500;
                padding: 0px;
                min-width: 80px;
                max-width: 80px;
                min-height: 34px;
                max-height: 34px;
            }
            QPushButton:hover {
                background: #4b5563;
            }
            QPushButton:pressed {
                background: #374151;
            }
        """)
        
        self.next_chart_btn = QtWidgets.QPushButton(tr('next'))
        self.next_chart_btn.setFixedHeight(34)
        self.next_chart_btn.setFixedWidth(80)
        self.next_chart_btn.setStyleSheet("""
            QPushButton {
                background: #6b7280;
                color: #fff;
                border: none;
                border-radius: 6px;
                font-size: 12px;
                font-weight: 500;
                padding: 0px;
                min-width: 80px;
                max-width: 80px;
                min-height: 34px;
                max-height: 34px;
            }
            QPushButton:hover {
                background: #4b5563;
            }
            QPushButton:pressed {
                background: #374151;
            }
        """)
        
        header.addWidget(self.prev_chart_btn)
        header.addWidget(self.next_chart_btn)
        main_vbox.addLayout(header)

        # 2. 上方：SPC 大圖
        self.figure = Figure(figsize=(8, 5.5))
        self.canvas = FigureCanvas(self.figure)
        main_vbox.addWidget(self.canvas, 6) # 權重調高

        # 3. 下方：兩張分析小圖 (水平排列)
        analysis_layout = QtWidgets.QHBoxLayout()
        analysis_layout.setSpacing(10) # 縮小間距
        analysis_layout.setContentsMargins(0, 0, 0, 0) # 移除邊距
        
        # 包裝小圖的函數（移除標題以節省空間）
        def wrap_sub_chart(canvas):
            container = QtWidgets.QVBoxLayout()
            container.setContentsMargins(0, 0, 0, 0)
            container.addWidget(canvas)
            return container

        # 只保留 Grouped Tool 和 QQ Plot
        analysis_layout.addLayout(wrap_sub_chart(self.canvas_sub2), 1)
        analysis_layout.addLayout(wrap_sub_chart(self.canvas_sub3), 1)
        
        main_vbox.addLayout(analysis_layout, 4) # 權重調低
        
        root.addWidget(self.chart_frame, 1)
        # 事件連接
        self.recalc_btn.clicked.connect(self.recalculate)
        self.chart_combo.currentIndexChanged.connect(self.update_cpk_labels)
        self.settings_btn.clicked.connect(self.open_settings_dialog)
        self.export_excel_btn.clicked.connect(self.export_chart_info_excel)
        self.prev_chart_btn.clicked.connect(self.prev_chart)
        self.next_chart_btn.clicked.connect(self.next_chart)
        self.apply_theme()
    def export_chart_info_excel(self):
        # 匯出所有 chart 的 group_name@chart_name@characteristics 及 Cpk 指標到 Excel，並加上 debug log
        from translations import tr
        if self.all_charts_info is None:
            QtWidgets.QMessageBox.warning(self, tr('warning'), tr('chart_info_not_loaded'))
            return
        
        # 創建進度對話框
        total_charts = len(self.all_charts_info)
        progress = QtWidgets.QProgressDialog(tr('exporting_charts'), tr('cancel'), 0, total_charts, self)
        progress.setWindowTitle(tr('export_progress'))
        progress.setWindowModality(QtCore.Qt.WindowModality.WindowModal)
        progress.setMinimumDuration(0)
        progress.setValue(0)
        
        rows = []
        chart_images = []
        for idx, (_, chart_info) in enumerate(self.all_charts_info.iterrows()):
            # 更新進度
            progress.setValue(idx)
            progress.setLabelText(f"{tr('processing_chart')} {idx+1}/{total_charts}: {chart_info.get('GroupName', '')}@{chart_info.get('ChartName', '')}")
            QtWidgets.QApplication.processEvents()
            
            # 檢查使用者是否取消
            if progress.wasCanceled():
                QtWidgets.QMessageBox.information(self, tr('export_cancelled'), tr('export_cancelled_msg'))
                # 清理已產生的臨時圖片
                for img_path in chart_images:
                    try:
                        import os
                        os.unlink(img_path)
                    except:
                        pass
                return
            group_name = str(chart_info.get('GroupName', ''))
            chart_name = str(chart_info.get('ChartName', ''))
            characteristics = str(chart_info.get('Characteristics', ''))
            usl = chart_info.get('USL', None)
            lsl = chart_info.get('LSL', None)
            target = None
            for key_t in ['Target', 'TARGET', 'TargetValue', '中心線', 'Center']:
                if key_t in chart_info and pd.notna(chart_info[key_t]):
                    target = chart_info[key_t]
                    break
            key = (group_name, chart_name)
            cpk = None
            cpk_last_month = None
            cpk_last2_month = None
            custom_cpk = None
            r1 = None
            r2 = None
            mean_month = sigma_month = mean_last_month = sigma_last_month = mean_last2_month = sigma_last2_month = mean_all = sigma_all = None
            
            # 使用當前模式計算 Cpk（與 UI 一致）
            if key in self.raw_charts_dict:
                raw_df = self.raw_charts_dict[key]
                
                # 使用統一的計算方法
                cpk_res = self._recompute_cpk_for_chart(chart_info)
                print(f"[DEBUG] cpk_res: {cpk_res}")
                cpk = cpk_res.get('Cpk')
                cpk_last_month = cpk_res.get('Cpk_last_month')
                cpk_last2_month = cpk_res.get('Cpk_last2_month')
                if raw_df is not None:
                    custom_cpk = calculate_cpk(raw_df, chart_info)['Cpk']
                # 只有當所有 Cpk 值都有效時才計算衰退率
                if cpk is not None and cpk_last_month is not None and cpk_last_month != 0 and cpk <= cpk_last_month:
                    r1 = (1 - (cpk / cpk_last_month)) * 100
                if cpk is not None and cpk_last_month is not None and cpk_last2_month is not None and cpk_last2_month != 0 and cpk <= cpk_last_month <= cpk_last2_month:
                    r2 = (1 - (cpk / cpk_last2_month)) * 100
                
                # 若任何 Cpk 為 None，對應的 R 也應為 None
                if cpk is None or cpk_last_month is None:
                    r1 = None
                if cpk is None or cpk_last_month is None or cpk_last2_month is None:
                    r2 = None
                # 計算四個區間的 mean, sigma 並印出
                def print_mean_sigma(df, label, group_name, chart_name):
                    if df is not None and not df.empty:
                        mean = df['point_val'].mean()
                        sigma = df['point_val'].std()
                        print(f"[MEAN_SIGMA][{group_name}@{chart_name}][{label}] mean: {mean:.4f}, sigma: {sigma:.4f}")
                    else:
                        print(f"[MEAN_SIGMA][{group_name}@{chart_name}][{label}] 無資料")
                # 取得各區間資料（使用等長區間）
                if raw_df is not None and not raw_df.empty and 'point_time' in raw_df.columns:
                    raw_df_local = raw_df.copy()
                    raw_df_local['point_time'] = pd.to_datetime(raw_df_local['point_time'])
                    
                    # 根據模式決定區間
                    if self.is_custom_mode:
                        end_time = pd.to_datetime(self.custom_end_date.toPyDate()) + pd.Timedelta(days=1) - pd.Timedelta(milliseconds=1)
                        start_time = pd.to_datetime(self.custom_start_date.toPyDate())
                    else:
                        end_time = raw_df_local['point_time'].max()
                        start_time = end_time - pd.DateOffset(months=3)
                    
                    # --- 關鍵修正：平分總時長 ---
                    total_range = end_time - start_time
                    duration_segment = total_range / 3  # 每一小格佔 1/3 時長
                    
                    # 重新定義三個等長區間
                    # [Start] --(L2)-- [Start+1/3] --(L1)-- [Start+2/3] --(Current)-- [End]
                    curr_start, curr_end = end_time - duration_segment, end_time
                    l1_start, l1_end     = curr_start - duration_segment, curr_start
                    l2_start, l2_end     = l1_start - duration_segment, l1_start
                    
                    df_all = raw_df_local  # 全部資料
                    df_month = raw_df_local[(raw_df_local['point_time'] >= curr_start) & (raw_df_local['point_time'] <= curr_end)]
                    df_last_month = raw_df_local[(raw_df_local['point_time'] >= l1_start) & (raw_df_local['point_time'] < l1_end)]
                    df_last2_month = raw_df_local[(raw_df_local['point_time'] >= l2_start) & (raw_df_local['point_time'] < l2_end)]
                    
                    mean_month = df_month['point_val'].mean() if not df_month.empty else None
                    sigma_month = df_month['point_val'].std() if not df_month.empty else None
                    mean_last_month = df_last_month['point_val'].mean() if not df_last_month.empty else None
                    sigma_last_month = df_last_month['point_val'].std() if not df_last_month.empty else None
                    mean_last2_month = df_last2_month['point_val'].mean() if not df_last2_month.empty else None
                    sigma_last2_month = df_last2_month['point_val'].std() if not df_last2_month.empty else None
                    mean_all = df_all['point_val'].mean() if not df_all.empty else None
                    sigma_all = df_all['point_val'].std() if not df_all.empty else None
                    print_mean_sigma(df_month, 'Current', group_name, chart_name)
                    print_mean_sigma(df_last_month, 'L1', group_name, chart_name)
                    print_mean_sigma(df_last2_month, 'L2', group_name, chart_name)
                    print_mean_sigma(df_all, 'All', group_name, chart_name)
            # --- 新增：用與 UI 完全一致的方式繪製圖表並存成圖片 ---
            # 計算 K 參數
            kval = None
            try:
                epsilon = 1e-9
                if target is not None and usl is not None and lsl is not None:
                    mean_val = None
                    plot_df = self.raw_charts_dict.get(key)
                    if plot_df is not None and not plot_df.empty:
                        mean_val = plot_df['point_val'].mean()
                    # 檢查 USL-LSL 是否過小（接近0）
                    if abs(usl - lsl) < epsilon:
                        kval = None  # 規格範圍過小，無法計算 K
                    else:
                        rng = (usl - lsl) / 2
                        if mean_val is not None and target is not None:
                            kval = abs(mean_val - target) / rng
            except Exception:
                kval = None
            import tempfile
            tmp_img = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
            tmp_img.close()
            # 用 draw_spc_chart 的繪圖邏輯（x軸刻度與UI一致，點與點等距）
            from matplotlib.figure import Figure
            fig = Figure(figsize=(8, 5))
            ax = fig.add_subplot(111)
            characteristics = chart_info.get('Characteristics', '')
            ax.set_title(f"{group_name}@{chart_name}@{characteristics}", pad=18)
            ax.set_xlabel("")
            ax.set_ylabel("")
            plot_df = self.raw_charts_dict.get(key)
            if plot_df is None or plot_df.empty:
                ax.text(0.5, 0.5, tr('no_data'), ha='center', va='center', transform=ax.transAxes)
            else:
                plot_df2 = plot_df.copy()
                # 日期過濾 (若有 point_time 欄位)
                if 'point_time' in plot_df2.columns:
                    try:
                        plot_df2['point_time'] = pd.to_datetime(plot_df2['point_time'])
                        plot_df2 = plot_df2.sort_values('point_time').reset_index(drop=True)
                    except Exception:
                        pass
                y = plot_df2['point_val'].values
                # x軸等距模式（與UI一致）
                x = range(1, len(y) + 1)
                # === 在圖上標示「當月/上月/上上月」區間 ===
                # --- 徹底修正：使用索引百分比來分配底色 ---
                if 'point_time' in plot_df2.columns and not plot_df2.empty:
                    try:
                        import matplotlib.transforms as mtransforms
                        n = len(plot_df2)
                        x_min, x_max = 0.5, n + 0.5
                        
                        # 取得數據的時間區間
                        times = pd.to_datetime(plot_df2['point_time']).to_numpy()
                        t_start, t_end = times.min(), times.max()
                        
                        # 根據模式定義時間界線
                        if self.is_custom_mode:
                            end_dt = pd.to_datetime(self.custom_end_date.toPyDate()) + pd.Timedelta(days=1) - pd.Timedelta(milliseconds=1)
                            start_dt = pd.to_datetime(self.custom_start_date.toPyDate())
                        else:
                            end_dt = pd.Timestamp(t_end)
                            start_dt = end_dt - pd.DateOffset(months=3)

                        # 1. 算出各時間界線對應的「浮點數索引」位置
                        def get_x_pos(target_t):
                            target_t = np.datetime64(target_t)
                            if target_t <= t_start: return x_min
                            if target_t >= t_end: return x_max
                            # 找到最接近的兩個點，進行線性插值，確保底色平滑銜接
                            idx = np.searchsorted(times, target_t)
                            return (idx) + 0.5

                        # 2. 定義三個區間的物理邊界點 (由右往左推)
                        total_delta = end_dt - start_dt
                        seg = total_delta / 3
                        
                        p3 = get_x_pos(end_dt)            # Current 結束 (最右)
                        p2 = get_x_pos(end_dt - seg)      # Current 開始 / L1 結束
                        p1 = get_x_pos(end_dt - 2 * seg)  # L1 開始 / L2 結束
                        p0 = get_x_pos(start_dt)          # L2 開始

                        # 強制讓 Current 延伸到最右邊界，解決懸空問題
                        p3 = x_max 

                        draw_windows = [
                            (p2, p3, 'L0', '#fee2e2'),
                            (p1, p2, 'L1', '#fef9c3'),
                            (p0, p1, 'L2', '#ede9fe'),
                        ]

                        text_trans = mtransforms.blended_transform_factory(ax.transData, ax.transAxes)
                        for xl, xr, lab, col in draw_windows:
                            if xr <= xl: continue
                            # 加入 linewidth=0 消除重疊處的深色條紋
                            ax.axvspan(xl, xr, color=col, alpha=0.25, zorder=0, linewidth=0)
                            ax.text((xl + xr) / 2, 1.04, lab, transform=text_trans, ha='center', va='top', 
                                    fontsize=9, color='#374151', alpha=0.9)
                            
                    except Exception as e:
                        print(f"[DEBUG] 底色繪製失敗: {e}")
                # 主數據線
                ax.plot(x, y, linestyle='-', marker='o', color='#2563eb', markersize=5, linewidth=1.2, label='_nolegend_')
                usl = chart_info.get('USL', None)
                lsl = chart_info.get('LSL', None)
                target = None
                for key_t in ['Target', 'TARGET', 'TargetValue', '中心線', 'Center']:
                    if key_t in chart_info and pd.notna(chart_info[key_t]):
                        target = chart_info[key_t]
                        break
                import numpy as np
                # mean_val = float(np.mean(y)) if len(y) else None
                # 超規點
                if usl is not None:
                    ax.scatter([xi for xi, yi in zip(x, y) if yi > usl], [yi for yi in y if yi > usl], color='#dc2626', s=36, zorder=5, label='_nolegend_')
                if lsl is not None:
                    ax.scatter([xi for xi, yi in zip(x, y) if yi < lsl], [yi for yi in y if yi < lsl], color='#dc2626', marker='s', s=36, zorder=5, label='_nolegend_')
                # Y軸範圍（納入 USL/LSL/Target/Mean）
                extra_vals = [v for v in [usl, lsl, target, mean_val]
                              if v is not None and not (isinstance(v, float) and np.isnan(v))]
                if len(y) > 0:
                    ymin_sel = float(np.min(y))
                    ymax_sel = float(np.max(y))
                else:
                    ymin_sel, ymax_sel = (0.0, 1.0)
                if extra_vals:
                    ymin_sel = min(ymin_sel, min(extra_vals))
                    ymax_sel = max(ymax_sel, max(extra_vals))
                rng = ymax_sel - ymin_sel
                margin = 0.05 * rng if rng > 0 else 1.0
                ax.set_ylim(ymin_sel - margin, ymax_sel + margin)
                # 畫短水平線，文字貼右
                from matplotlib import transforms as mtransforms
                trans = mtransforms.blended_transform_factory(ax.transAxes, ax.transData)
                def segment_with_label(val, name, color, va='center'):
                    if val is None or (isinstance(val, float) and np.isnan(val)):
                        return
                    x0, x1 = 0.0, 0.965
                    ax.plot([x0, x1], [val, val], transform=trans, color=color, linestyle='--', linewidth=1.1)
                    ax.text(x1, val, name, transform=trans, color=color, va=va, ha='left', fontsize=9)
                segment_with_label(usl, 'USL', '#ef4444', va='center')
                segment_with_label(lsl, 'LSL', '#ef4444', va='center')
                segment_with_label(target, 'Target', '#f59e0b', va='center')
                # segment_with_label(mean_val, 'Mean', '#16a34a', va='center')
                # x軸刻度（等距模式顯示日期）
                if 'point_time' in plot_df2.columns and not plot_df2.empty:
                    times = plot_df2['point_time'].tolist()
                    total = len(times)
                    if total <= 12:
                        tick_idx = list(range(1, total + 1))
                    else:
                        step = max(1, total // 8)
                        tick_idx = list(range(1, total + 1, step))
                        if tick_idx[-1] != total:
                            tick_idx.append(total)
                    labels = [times[i-1].strftime('%Y-%m-%d') for i in tick_idx]
                    ax.set_xticks(tick_idx)
                    ax.set_xticklabels(labels, rotation=90, ha='center', fontsize=8)
                ax.grid(True, linestyle=':', linewidth=0.6, alpha=0.5)
            fig.tight_layout()
            fig.savefig(tmp_img.name)
            chart_images.append(tmp_img.name)
            
            # 生成子圖 1: Tool Variation Comparison
            tmp_img_sub2 = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
            tmp_img_sub2.close()
            self._generate_tool_comparison_chart(plot_df, tmp_img_sub2.name)
            chart_images.append(tmp_img_sub2.name)
            
            # 生成子圖 2: Normal Probability Plot  
            tmp_img_sub3 = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
            tmp_img_sub3.close()
            self._generate_qq_plot_chart(plot_df, tmp_img_sub3.name)
            chart_images.append(tmp_img_sub3.name)
            
            # 釋放記憶體：清除並關閉圖形對象
            fig.clear()
            import matplotlib.pyplot as plt
            plt.close(fig)
            del fig, ax
            rows.append({
                'ChartImage': '',  # 主圖佔位
                'SubChart1': '',  # 子圖1佔位
                'SubChart2': '',  # 子圖2佔位
                'ChartKey': f"{group_name}@{chart_name}@{characteristics}",
                'GroupName': group_name,
                'ChartName': chart_name,
                'Characteristics': characteristics,
                'USL': usl,
                'LSL': lsl,
                'Target': target,
                'K': kval,
                'Cpk_Current': cpk,
                'Cpk_L1': cpk_last_month,
                'Cpk_L2': cpk_last2_month,
                'Custom_Cpk': custom_cpk,
                'R1(%)': r1,
                'R2(%)': r2,
                'Mean_Current': mean_month,
                'Sigma_CurrentMonth': sigma_month,
                'Mean_LastMonth': mean_last_month,
                'Sigma_LastMonth': sigma_last_month,
                'Mean_Last2Month': mean_last2_month,
                'Sigma_Last2Month': sigma_last2_month,
                'Mean_All': mean_all,
                'Sigma_All': sigma_all
            })
        
        # 完成進度
        progress.setValue(total_charts)
        progress.close()
        
        df = pd.DataFrame(rows)
        path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Download Chart Information Excel", "cpk_analysis_report.xlsx", "Excel Files (*.xlsx)")
        if path:
            try:
                import xlsxwriter
                # 欄位順序：三張圖片 + 其他欄位
                columns = ['ChartImage', 'SubChart1', 'SubChart2'] + [c for c in df.columns if c not in ['ChartImage', 'SubChart1', 'SubChart2']]
                workbook = xlsxwriter.Workbook(path)
                worksheet = workbook.add_worksheet()
                # 設定欄寬（前三欄放圖片，每欄設寬 40，ChartKey 設寬 40，其他 18）
                worksheet.set_column(0, 2, 40)  # 三張圖片的欄位
                for i in range(3, len(columns)):
                    if columns[i] == 'ChartKey':
                        worksheet.set_column(i, i, 40)
                    else:
                        worksheet.set_column(i, i, 18)
                # 標題粗體
                bold = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter'})
                cell_format = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
                for col_idx, col_name in enumerate(columns):
                    worksheet.write(0, col_idx, col_name, bold)
                
                # 寫入資料 & 插入圖片
                # 主圖尺寸 (15x2 inch, 100 dpi) -> 1500x200 px
                # 子圖尺寸 (4x4 inch, 100 dpi) -> 400x400 px
                main_img_width_px = 1500
                main_img_height_px = 200
                sub_img_width_px = 400
                sub_img_height_px = 400
                
                # Excel 單位：欄寬用字元寬度（96 DPI 下，1 字元 ≈ 7 px），列高用點 (1 點 = 1.333 px)
                # 目標顯示尺寸（像素）- 稍微加大以改善視覺效果
                target_main_width = 500   # 主圖目標顯示寬度（像素）
                target_main_height = 50  # 主圖目標顯示高度（像素）
                target_sub_width = 300    # 子圖目標顯示寬度（像素）
                target_sub_height = 150   # 子圖目標顯示高度（像素）
                
                # 計算縮放比例
                main_x_scale = target_main_width / main_img_width_px
                main_y_scale = target_main_height / main_img_height_px
                sub_x_scale = target_sub_width / sub_img_width_px
                sub_y_scale = target_sub_height / sub_img_height_px
                
                # 設定欄寬（字元單位）- 根據目標像素寬度調整
                # 主圖 300px ≈ 43 字元，加大到 45 避免邊緣被切
                # 子圖 150px ≈ 21 字元，設為 22
                main_col_width = 45  # 主圖欄位寬度
                sub_col_width = 45  # 子圖欄位寬度
                
                worksheet.set_column(0, 0, main_col_width)  # 主圖欄
                worksheet.set_column(1, 1, sub_col_width)   # 子圖1欄
                worksheet.set_column(2, 2, sub_col_width)   # 子圖2欄
                
                for row_idx, row in enumerate(df.to_dict('records')):
                    # 插入三張圖片：主圖、子圖1、子圖2 
                    img_idx = row_idx * 3  # 每個 row 對應 3 張圖
                    main_img = chart_images[img_idx]
                    sub1_img = chart_images[img_idx + 1]
                    sub2_img = chart_images[img_idx + 2]
                    
                    # 設定列高（點數）
                    # 1 點 = 1.333 像素，所以像素高度 * 0.75 = 點數
                    # 加 15 點作為上下留白，避免圖片頂到格線
                    max_img_height = max(target_main_height, target_sub_height)
                    row_height_pts = (max_img_height * 0.75) + 15
                    worksheet.set_row(row_idx+1, row_height_pts)
                    
                    # 圖片偏移量（像素）
                    # x_offset: 向右偏移，避免緊貼左邊界
                    # y_offset: 向下偏移，避免頂到上方格線
                    main_x_offset = 10
                    main_y_offset = 8
                    sub_x_offset = 10
                    sub_y_offset = 8
                    
                    # 主圖 (object_position=2: 移動並隨儲存格調整大小)
                    worksheet.insert_image(row_idx+1, 0, main_img, {
                        'x_scale': main_x_scale,
                        'y_scale': main_y_scale,
                        'object_position': 2,
                        'x_offset': main_x_offset,
                        'y_offset': main_y_offset
                    })
                    
                    # 子圖1
                    worksheet.insert_image(row_idx+1, 1, sub1_img, {
                        'x_scale': sub_x_scale,
                        'y_scale': sub_y_scale,
                        'object_position': 2,
                        'x_offset': sub_x_offset,
                        'y_offset': sub_y_offset
                    })
                    
                    # 子圖2
                    worksheet.insert_image(row_idx+1, 2, sub2_img, {
                        'x_scale': sub_x_scale,
                        'y_scale': sub_y_scale,
                        'object_position': 2,
                        'x_offset': sub_x_offset,
                        'y_offset': sub_y_offset
                    })
                    
                    # 其他欄位（置中）
                    for col_idx, col_name in enumerate(columns[3:], 3):  # 從第4欄開始（前3欄是圖片）
                        val = row.get(col_name, '')
                        # 修正 NaN/Inf/None 問題
                        import math
                        if val is None:
                            val = 'N/A'
                        elif isinstance(val, float):
                            if math.isnan(val) or math.isinf(val):
                                val = 'N/A'
                        worksheet.write(row_idx+1, col_idx, val, cell_format)
                workbook.close()
                QtWidgets.QMessageBox.information(self, tr('export_successful'), f"{tr('export_successful_msg')} {path}")
            except Exception as e:
                QtWidgets.QMessageBox.critical(self, tr('export_failed'), f"{tr('export_failed_msg')} {e}")
            finally:
                # 清理臨時圖片檔案
                import os
                for img_path in chart_images:
                    try:
                        if os.path.exists(img_path):
                            os.unlink(img_path)
                    except Exception as cleanup_err:
                        print(f"[WARN] 清理臨時圖片失敗: {img_path}, {cleanup_err}")

    def _generate_tool_comparison_chart(self, df, output_path):
        """生成 Tool Variation Comparison 子圖並儲存"""
        import scipy.stats as stats
        from matplotlib.figure import Figure
        import matplotlib.pyplot as plt
        
        fig = Figure(figsize=(4, 4))
        ax = fig.add_subplot(111)
        
        if df is None or df.empty:
            ax.text(0.5, 0.5, "No Data", ha='center', va='center')
            fig.savefig(output_path)
            plt.close(fig)
            return
        
        # 自動偵測 Tool 欄位
        tool_col = next((c for c in ['ByTool', 'tool_id', 'Tool', '機台'] if c in df.columns), None)
        
        if tool_col:
            df_plot = df.copy()
            df_plot['point_val'] = pd.to_numeric(df_plot['point_val'], errors='coerce')
            df_plot = df_plot.dropna(subset=['point_val'])
            df_plot = df_plot[pd.notna(df_plot[tool_col])].copy()
            
            if not df_plot.empty:
                tools = sorted(df_plot[tool_col].unique())
                color_cycle = ['#2563eb', '#dc2626', '#16a34a', '#f59e0b', '#7c3aed', '#0891b2']
                tool_colors = {t: color_cycle[i % len(color_cycle)] for i, t in enumerate(tools)}
                
                ax.set_title("Tool Variation Comparison", fontsize=10, fontweight='bold', pad=10)
                ax.set_ylabel("Measured Value", fontsize=8)
                
                group_width = 10 
                gap = 5
                for i, t in enumerate(tools):
                    subset = df_plot[df_plot[tool_col] == t]
                    if 'point_time' in subset.columns:
                        subset = subset.sort_values('point_time').reset_index(drop=True)
                    else:
                        subset = subset.reset_index(drop=True)
                        
                    start_x = i * (group_width + gap)
                    x_internal = start_x + subset.index
                    color = tool_colors[t]
                    
                    ax.plot(x_internal, subset['point_val'], color=color, alpha=0.3, linewidth=1, zorder=2, label='_nolegend_')
                    ax.scatter(x_internal, subset['point_val'], color=color, s=20, alpha=0.7, zorder=3, edgecolors='white', linewidth=0.5, label=str(t))
                
                ax.set_xticks([])
                ax.grid(True, axis='y', linestyle=':', alpha=0.4)
                ax.tick_params(labelsize=8)
                ax.legend(fontsize=7, loc='upper left', frameon=True, ncol=3, framealpha=0.9, edgecolor='#d1d5db')
            else:
                ax.text(0.5, 0.5, "No Tool Info", ha='center', va='center')
        else:
            ax.text(0.5, 0.5, "No Tool Column", ha='center', va='center')
        
        fig.subplots_adjust(left=0.18, bottom=0.22, right=0.95, top=0.85)
        fig.savefig(output_path)
        plt.close(fig)

    def _generate_qq_plot_chart(self, df, output_path):
        """生成 Normal Probability Plot 子圖並儲存"""
        import scipy.stats as stats
        from matplotlib.figure import Figure
        import matplotlib.pyplot as plt
        
        fig = Figure(figsize=(4, 4))
        ax = fig.add_subplot(111)
        
        if df is None or df.empty:
            ax.text(0.5, 0.5, "No Data", ha='center', va='center')
            fig.savefig(output_path)
            plt.close(fig)
            return
        
        # 自動偵測 Tool 欄位
        tool_col = next((c for c in ['ByTool', 'tool_id', 'Tool', '機台'] if c in df.columns), None)
        
        if tool_col:
            df_plot = df.copy()
            df_plot['point_val'] = pd.to_numeric(df_plot['point_val'], errors='coerce')
            df_plot = df_plot.dropna(subset=['point_val'])
            df_plot = df_plot[pd.notna(df_plot[tool_col])].copy()
            
            if not df_plot.empty:
                tools = sorted(df_plot[tool_col].unique())
                color_cycle = ['#2563eb', '#dc2626', '#16a34a', '#f59e0b', '#7c3aed', '#0891b2']
                tool_colors = {t: color_cycle[i % len(color_cycle)] for i, t in enumerate(tools)}
                
                ax.set_title("Normal Probability Plot", fontsize=10, fontweight='bold', pad=10)
                ax.set_xlabel("Theoretical Quantiles (σ)", fontsize=8)
                ax.set_ylabel("Sample Quantiles (Value)", fontsize=8)
                
                for t in tools:
                    tool_data = df_plot[df_plot[tool_col] == t]['point_val']
                    if len(tool_data) > 3:
                        (osm, osr), (slope, intercept, r) = stats.probplot(tool_data, dist="norm")
                        ax.scatter(osm, osr, color=tool_colors[t], s=15, alpha=0.5, label=t, edgecolors='none')
                        ax.plot(osm, slope*osm + intercept, color=tool_colors[t], alpha=0.2, linewidth=1)
                
                ax.set_xticks([-3, -2, -1, 0, 1, 2, 3])
                ax.grid(True, linestyle=':', alpha=0.4)
                ax.tick_params(labelsize=8)
                ax.legend(fontsize=7, loc='upper left', frameon=True, ncol=3, framealpha=0.9, edgecolor='#d1d5db')
            else:
                ax.text(0.5, 0.5, "No Tool Info", ha='center', va='center')
        else:
            ax.text(0.5, 0.5, "No Tool Column", ha='center', va='center')
        
        fig.subplots_adjust(left=0.18, bottom=0.22, right=0.95, top=0.85)
        fig.savefig(output_path)
        plt.close(fig)

    # === File Loading ===
    def load_csv(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Select CSV File", "", "CSV Files (*.csv)")
        if path:
            try:
                self.data = pd.read_csv(path)
                self.file_label.setText(f"Loaded: {os.path.basename(path)}")
            except Exception as e:
                QtWidgets.QMessageBox.critical(self, "Error", f"Loading failed: {e}")
                self.file_label.setText("Loading failed")
            self.recalculate()

    # === 重新計算 ===
    def recalculate(self):
        print("[DEBUG] recalculate called")
        
        # 1. 設定路徑
        chart_excel_path = os.path.join(get_app_dir(), 'input', 'All_Chart_Information.xlsx')
        
        # 2. 載入 Excel (直接使用上面 import 的 oob_module)
        try:
            # 這裡直接用 oob_module，不用再寫一堆 try...except import
            self.all_charts_info = oob_module.load_chart_information(chart_excel_path)
            if self.all_charts_info is None:
                raise ValueError("Excel 返回內容為空")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"無法讀取 Chart 設定檔: {e}")
            return

        # 3. 更新下拉選單 (直接清除 Select Chart，選中第一筆)
        self.chart_combo.blockSignals(True)
        self.chart_combo.clear()
        for _, info in self.all_charts_info.iterrows():
            self.chart_combo.addItem(f"{info['GroupName']} - {info['ChartName']}")
        self.chart_combo.blockSignals(False)

        # 4. 重置內部資料狀態
        self.raw_charts_dict = {}
        self.cpk_results = {}
        self.chart_date_states = {}

        # 5. 載入所有圖表的 raw data 並計算初始 Cpk
        raw_data_dir = os.path.join(get_app_dir(), 'input', 'raw_charts')
        for _, chart_info in self.all_charts_info.iterrows():
            if not isinstance(chart_info, pd.Series):
                continue
            group_name = str(chart_info['GroupName'])
            chart_name = str(chart_info['ChartName'])
            raw_path = oob_module.find_matching_file(raw_data_dir, group_name, chart_name)
            if raw_path and os.path.exists(raw_path):
                try:
                    raw_df = pd.read_csv(raw_path)
                    usl = chart_info.get('USL', None)
                    lsl = chart_info.get('LSL', None)
                    # 過濾超規點
                    if usl is not None and lsl is not None:
                        raw_df = raw_df[(raw_df['point_val'] <= usl) & (raw_df['point_val'] >= lsl)]
                    elif usl is not None:
                        raw_df = raw_df[raw_df['point_val'] <= usl]
                    elif lsl is not None:
                        raw_df = raw_df[raw_df['point_val'] >= lsl]
                    self.raw_charts_dict[(group_name, chart_name)] = raw_df
                    quick_cpk = calculate_cpk(raw_df, chart_info)['Cpk']
                    self.cpk_results[(group_name, chart_name)] = {'Cpk': quick_cpk}
                    self.chart_date_states[(group_name, chart_name)] = {'custom': False, 'start': None, 'end': None}
                except Exception as e:
                    self.raw_charts_dict[(group_name, chart_name)] = None
                    self.cpk_results[(group_name, chart_name)] = {'Cpk': None}
                    print(f"[ERROR] raw chart 載入失敗 {group_name}/{chart_name}: {e}")
            else:
                self.raw_charts_dict[(group_name, chart_name)] = None
                self.cpk_results[(group_name, chart_name)] = {'Cpk': None}

        # 6. 自動觸發第一張圖表的分析
        if self.chart_combo.count() > 0:
            self.chart_combo.setCurrentIndex(0)
            self.update_cpk_labels()

    def apply_theme(self, mode: str = "light"):
        if mode == "light":
            self.setStyleSheet("""
            QWidget { background:#eef1f5; color:#222; font-family:'Microsoft YaHei'; font-size:13px; }
            QComboBox, QDateEdit { background:#ffffff; border:1px solid #c5ccd4; padding:4px 8px; border-radius:7px; }
            QComboBox:hover, QDateEdit:hover { border:1px solid #98a3af; }
            QPushButton { background:#2563eb; color:#fff; border:none; padding:7px 18px; border-radius:8px; font-weight:600; }
            QPushButton:hover { background:#1d4fd8; }
            QPushButton:pressed { background:#163fae; }
            QFrame#metricCard, QFrame#metricCard * { background:#ffffff !important; }
            QFrame#metricCard { border:1px solid #d8dde2; border-radius:16px; }
            QLabel#metricTitle { font-size:11px; font-weight:600; color:#6c7681; letter-spacing:1px; }
            QLabel#metricValue { font-size:30px; font-weight:700; color:#111827; }
            QFrame#metricCard:hover { border:1px solid #aeb5bb; }
            QFrame#chartFrame { background:#ffffff; border:1px solid #d2d7dc; border-radius:22px; }
            QLabel#sectionTitle { font-size:15px; font-weight:600; color:#1f2937; background:transparent; }
            QLabel#plainLabel { font-size:13px; font-weight:600; color:#1f2937; background:transparent; }
            """)
        for meta in self.metric_cards.values():
            if meta["frame"].graphicsEffect() is None:
                eff = QtWidgets.QGraphicsDropShadowEffect(self)
                eff.setBlurRadius(18)
                eff.setOffset(0, 4)
                eff.setColor(QtGui.QColor(0, 0, 0, 26))
                meta["frame"].setGraphicsEffect(eff)
        if self.chart_frame.graphicsEffect() is None:
            eff2 = QtWidgets.QGraphicsDropShadowEffect(self)
            eff2.setBlurRadius(28)
            eff2.setOffset(0, 5)
            eff2.setColor(QtGui.QColor(0, 0, 0, 30))
            self.chart_frame.setGraphicsEffect(eff2)

    # ==== 重複定義刪除 (上方已有 recalculate) ====

    # (duplicate apply_theme & recalculate removed)

    def _apply_card_status(self, key: str, status: str):
        # 不再改變邊框顏色，保持固定樣式
        return

    def update_cpk_labels(self):
        """選擇 chart 時：若該 chart 尚未自訂日期 -> 自動用最新往回三個月"""
        idx = self.chart_combo.currentIndex()
        for key, comp in self.metric_cards.items():
            comp["value_label"].setText("-")
        
        # idx=-1 或無效索引時不處理
        if idx < 0 or self.all_charts_info is None or idx >= len(self.all_charts_info):
            self.figure.clear()
            ax = self.figure.add_subplot(111)
            ax.set_title("SPC Control Chart (Not Selected)")
            self.canvas.draw()
            return
            
        # chart_combo 的索引直接對應 DataFrame 索引
        chart_info = self.all_charts_info.iloc[idx]
        group_name = str(chart_info['GroupName'])
        chart_name = str(chart_info['ChartName'])
        key = (group_name, chart_name)
        raw_df = self.raw_charts_dict.get(key)
        
        state = self.chart_date_states.get(key)
        if state is None:
            state = {'custom': False, 'start': None, 'end': None}
            self.chart_date_states[key] = state

        # --- 修正點：移除對 self.start_date / self.end_date 的呼叫 ---
        if (not state['custom']) and raw_df is not None and not raw_df.empty:
            try:
                tmp = raw_df.copy()
                tmp['point_time'] = pd.to_datetime(tmp['point_time'], errors='coerce')
                tmp = tmp.dropna(subset=['point_time'])
                if not tmp.empty:
                    latest = tmp['point_time'].max()
                    start_candidate = latest - pd.DateOffset(months=3)
                    
                    # 直接更新狀態變數，主視窗不需要去 setDate 給不存在的 widget
                    state['start'] = start_candidate.date()
                    state['end'] = latest.date()
                    
                    # 同步更新主視窗保存的預設日期
                    self.custom_start_date = QtCore.QDate(start_candidate.year, start_candidate.month, start_candidate.day)
                    self.custom_end_date = QtCore.QDate(latest.year, latest.month, latest.day)
            except Exception as e:
                print(f"[WARN] 自動日期設定失敗: {e}")

        # 執行繪圖與計算
        self._update_current_chart_dynamic(chart_info)

    # === Cpk 動態計算 - 重構支持兩種模式 ===
    def _compute_cpk_equal_duration_windows(self, raw_df: pd.DataFrame, chart_info: pd.Series, start_time: pd.Timestamp, end_time: pd.Timestamp):
        """根據指定區間，將其平分為三段：L2, L1, Current"""
        result = {'Cpk': None, 'Cpk_last_month': None, 'Cpk_last2_month': None}
        
        if raw_df is None or raw_df.empty:
            return result
        
        df = raw_df.copy()
        df['point_time'] = pd.to_datetime(df['point_time'])
        
        # --- 關鍵修正：平分總時長 ---
        total_range = end_time - start_time
        duration_segment = total_range / 3  # 每一小格佔 1/3 時長
        
        # 重新定義三個等長區間
        # [Start] --(L2)-- [Start+1/3] --(L1)-- [Start+2/3] --(Current)-- [End]
        curr_start, curr_end = end_time - duration_segment, end_time
        l1_start, l1_end     = curr_start - duration_segment, curr_start
        l2_start, l2_end     = l1_start - duration_segment, l1_start
        
        # 篩選數據
        mask_curr = (df['point_time'] >= curr_start) & (df['point_time'] <= curr_end)
        mask_l1   = (df['point_time'] >= l1_start)   & (df['point_time'] < l1_end)
        mask_l2   = (df['point_time'] >= l2_start)   & (df['point_time'] < l2_end)
        
        # 分段計算 Cpk
        if mask_curr.any():
            result['Cpk'] = calculate_cpk(df[mask_curr], chart_info)['Cpk']
        if mask_l1.any():
            result['Cpk_last_month'] = calculate_cpk(df[mask_l1], chart_info)['Cpk']
        if mask_l2.any():
            result['Cpk_last2_month'] = calculate_cpk(df[mask_l2], chart_info)['Cpk']
            
        return result

    def _recompute_cpk_for_chart(self, chart_info: pd.Series):
        """根據當前模式重新計算 Cpk"""
        group_name = str(chart_info['GroupName'])
        chart_name = str(chart_info['ChartName'])
        raw_df = self.raw_charts_dict.get((group_name, chart_name))
        
        if raw_df is None or raw_df.empty:
            return {'Cpk': None, 'Cpk_last_month': None, 'Cpk_last2_month': None}
        
        if 'point_time' not in raw_df.columns:
            return {'Cpk': calculate_cpk(raw_df, chart_info)['Cpk'], 'Cpk_last_month': None, 'Cpk_last2_month': None}
        
        raw_df_local = raw_df.copy()
        raw_df_local['point_time'] = pd.to_datetime(raw_df_local['point_time'])
        
        if self.is_custom_mode:
            # 自訂模式：使用指定的起始和結束時間
            start_time = pd.to_datetime(self.custom_start_date.toPyDate())
            end_time = pd.to_datetime(self.custom_end_date.toPyDate()) + pd.Timedelta(days=1) - pd.Timedelta(milliseconds=1)
        else:
            # 自動偵測模式：找最新時間點，往回推 3 個月
            latest = raw_df_local['point_time'].max()
            end_time = latest
            start_time = end_time - pd.DateOffset(months=3)
        
        return self._compute_cpk_equal_duration_windows(raw_df_local, chart_info, start_time, end_time)

    def _update_current_chart_dynamic(self, chart_info: pd.Series):
        group_name = str(chart_info['GroupName'])
        chart_name = str(chart_info['ChartName'])
        
        # 定義卡片更新函數（放在開頭以避免未定義錯誤）
        def set_card(key, value, is_percent=False):
            comp = self.metric_cards[key]
            if value is None:
                comp['value_label'].setText('-')
            else:
                comp['value_label'].setText(f"{value:.1f}%" if is_percent else f"{value:.3f}")
        
        raw_df = self.raw_charts_dict.get((group_name, chart_name))
        
        # 檢查數據有效性
        if raw_df is None or raw_df.empty:
            # 無數據，所有指標設為空
            for key in ['kval', 'cpk', 'l1', 'l2', 'custom', 'r1', 'r2']:
                set_card(key, None)
            self.draw_spc_chart(group_name, chart_name, chart_info)
            return
        
        # 根據當前模式計算 Cpk
        cpk_res = self._recompute_cpk_for_chart(chart_info)
        
        # 計算全部資料 Cpk (Long-term Cpk)
        all_data_cpk = calculate_cpk(raw_df, chart_info)['Cpk']
        
        # 計算 K 參數
        kval = None
        try:
            epsilon = 1e-9
            usl = chart_info.get('USL', None)
            lsl = chart_info.get('LSL', None)
            target = None
            for key_t in ['Target', 'TARGET', 'TargetValue', '中心線', 'Center']:
                if key_t in chart_info and pd.notna(chart_info[key_t]):
                    target = chart_info[key_t]
                    break
            
            mean_val = None
            if raw_df is not None and not raw_df.empty:
                # 根據模式決定使用哪段數據的 mean
                if self.is_custom_mode:
                    # 自訂模式：使用指定區間的數據
                    start_time = pd.to_datetime(self.custom_start_date.toPyDate())
                    end_time = pd.to_datetime(self.custom_end_date.toPyDate()) + pd.Timedelta(days=1) - pd.Timedelta(milliseconds=1)
                    if 'point_time' in raw_df.columns:
                        filtered_df = raw_df[(pd.to_datetime(raw_df['point_time']) >= start_time) & 
                                            (pd.to_datetime(raw_df['point_time']) <= end_time)]
                        if not filtered_df.empty:
                            mean_val = filtered_df['point_val'].mean()
                    else:
                        mean_val = raw_df['point_val'].mean()
                else:
                    # 自動模式：使用最近 3 個月的數據
                    if 'point_time' in raw_df.columns:
                        df_local = raw_df.copy()
                        df_local['point_time'] = pd.to_datetime(df_local['point_time'])
                        latest = df_local['point_time'].max()
                        start_time = latest - pd.DateOffset(months=3)
                        filtered_df = df_local[(df_local['point_time'] >= start_time) & (df_local['point_time'] <= latest)]
                        if not filtered_df.empty:
                            mean_val = filtered_df['point_val'].mean()
                    else:
                        mean_val = raw_df['point_val'].mean()
            
            # 檢查 USL-LSL 是否過小
            if usl is not None and lsl is not None:
                if abs(usl - lsl) < epsilon:
                    kval = None
                else:
                    rng = (usl - lsl) / 2
                    if mean_val is not None and target is not None:
                        kval = abs(mean_val - target) / rng
        except Exception:
            kval = None
        
        # 更新卡片
        set_card('kval', kval)
        set_card('cpk', cpk_res.get('Cpk'))
        set_card('l1', cpk_res.get('Cpk_last_month'))
        set_card('l2', cpk_res.get('Cpk_last2_month'))
        set_card('custom', all_data_cpk)
        
        # 計算衰退率
        cpk = cpk_res.get('Cpk')
        l1 = cpk_res.get('Cpk_last_month')
        l2 = cpk_res.get('Cpk_last2_month')
        
        r1 = r2 = None
        if cpk is not None and l1 is not None and l1 != 0 and cpk <= l1:
            r1 = (1 - (cpk / l1)) * 100
        if cpk is not None and l1 is not None and l2 is not None and l2 != 0 and cpk <= l1 <= l2:
            r2 = (1 - (cpk / l2)) * 100
        
        # 若任何 Cpk 為 None，對應的 R 也應為 None
        if cpk is None or l1 is None:
            r1 = None
        if cpk is None or l1 is None or l2 is None:
            r2 = None
        
        set_card('r1', r1, is_percent=True)
        set_card('r2', r2, is_percent=True)
        
        # 依目前設定重畫圖
        self.draw_spc_chart(group_name, chart_name, chart_info)

    # === X 軸模式切換 ===
    def toggle_axis_mode(self):
        self.axis_mode = 'time' if self.axis_mode == 'index' else 'index'
        # 更新按鈕文字
        self.axis_mode_btn.setText('Equal Spacing' if self.axis_mode == 'time' else 'Time Axis')
        # 重新繪圖（若已選 chart）
        idx = self.chart_combo.currentIndex()
        if idx >= 0 and self.all_charts_info is not None:
            chart_info = self.all_charts_info.iloc[idx]
            group_name = str(chart_info['GroupName'])
            chart_name = str(chart_info['ChartName'])
            self.draw_spc_chart(group_name, chart_name, chart_info)

    def draw_spc_chart(self, group_name: str, chart_name: str, chart_info):
        try:
            raw_df = self.raw_charts_dict.get((group_name, chart_name))
            self.figure.clear()
            ax = self.figure.add_subplot(111)
            # 標題格式: [GroupName@ChartName@Characteristics]
            characteristics = chart_info.get('Characteristics', '')
            ax.set_title(f"{group_name}@{chart_name}@{characteristics}", pad=10)  # 減少 pad 以節省空間
            ax.set_xlabel("" if self.axis_mode == 'index' else "")
            # ax.set_ylabel("值")
            if raw_df is None or raw_df.empty:
                from translations import tr
                ax.text(0.5, 0.5, tr('no_data'), ha='center', va='center', transform=ax.transAxes)
                self.canvas.draw()
                return
            
            plot_df = raw_df.copy()
            
            # 數據清理：先過濾掉非數字的 point_val
            plot_df['point_val'] = pd.to_numeric(plot_df['point_val'], errors='coerce')
            plot_df = plot_df.dropna(subset=['point_val'])
            
            # X 軸處理：確保時間排序
            if 'point_time' in plot_df.columns:
                try:
                    plot_df['point_time'] = pd.to_datetime(plot_df['point_time'], errors='coerce')
                    plot_df = plot_df.dropna(subset=['point_time'])
                    plot_df = plot_df.sort_values('point_time').reset_index(drop=True)
                    
                    # 同步指標計算區間：若為自動模式，只顯示最近 3 個月的數據
                    if not self.is_custom_mode and not plot_df.empty:
                        latest = plot_df['point_time'].max()
                        three_months_ago = latest - pd.DateOffset(months=3)
                        plot_df = plot_df[plot_df['point_time'] >= three_months_ago].reset_index(drop=True)
                except Exception as e:
                    print(f"[WARN] 時間轉換失敗: {e}")
                    pass
            
            if plot_df.empty:
                from translations import tr
                ax.text(0.5, 0.5, tr('no_data'), ha='center', va='center', transform=ax.transAxes)
                self.canvas.draw()
                return
            
            # X 軸處理：
            y = plot_df['point_val'].values
            use_time_axis = False
            if self.axis_mode == 'time' and 'point_time' in plot_df.columns:
                try:
                    plot_df['point_time'] = pd.to_datetime(plot_df['point_time'])
                    plot_df = plot_df.sort_values('point_time')
                    x = plot_df['point_time'].values
                    use_time_axis = True
                except Exception:
                    x = range(1, len(y) + 1)
            else:
                # 等距模式：保持所有點等距，避免同時間戳疊在一起
                if 'point_time' in plot_df.columns:
                    try:
                        plot_df['point_time'] = pd.to_datetime(plot_df['point_time'])
                        plot_df = plot_df.sort_values('point_time').reset_index(drop=True)
                    except Exception:
                        pass
                x = range(1, len(y) + 1)

            # === 在圖上標示三個等長區間 (Curr, L1, L2) ===
            # --- 徹底修正：使用索引百分比來分配底色 ---
            if 'point_time' in plot_df.columns and not plot_df.empty:
                try:
                    # 確保所有 point_time 都是有效的 datetime
                    plot_df_times = plot_df[pd.notna(plot_df['point_time'])].copy()
                    if plot_df_times.empty:
                        raise ValueError("無有效時間數據")
                    import matplotlib.transforms as mtransforms
                    n = len(plot_df)
                    x_min, x_max = 0.5, n + 0.2
                    
                    # 取得數據的時間區間
                    times = pd.to_datetime(plot_df['point_time']).to_numpy()
                    t_start, t_end = times.min(), times.max()
                    
                    # 根據模式定義時間界線
                    if self.is_custom_mode:
                        end_dt = pd.to_datetime(self.custom_end_date.toPyDate()) + pd.Timedelta(days=1) - pd.Timedelta(milliseconds=1)
                        start_dt = pd.to_datetime(self.custom_start_date.toPyDate())
                    else:
                        end_dt = pd.Timestamp(t_end)
                        start_dt = end_dt - pd.DateOffset(months=3)

                    # 1. 算出各時間界線對應的「浮點數索引」位置
                    def get_x_pos(target_t):
                        try:
                            # 確保 target_t 是 pandas Timestamp
                            if not isinstance(target_t, pd.Timestamp):
                                target_t = pd.Timestamp(target_t)
                            
                            # 轉為 numpy datetime64 進行比較
                            target_np = np.datetime64(target_t)
                            if target_np <= t_start: 
                                return x_min
                            if target_np >= t_end: 
                                return x_max
                            
                            # 找到最接近的索引位置
                            time_diff = np.abs((times - target_np).astype('timedelta64[s]').astype(float))
                            closest_idx = np.argmin(time_diff)
                            return closest_idx + 1  # 轉換為 1-based 索引
                        except Exception as e:
                            print(f"[WARN] get_x_pos 失敗 for {target_t}: {e}")
                            return x_min

                    # 2. 定義三個區間的物理邊界點 (由右往左推)
                    total_delta = end_dt - start_dt
                    seg = total_delta / 3
                    
                    # 確保時間計算正確
                    if total_delta.total_seconds() <= 0:
                        raise ValueError("時間區間無效")
                    
                    p3 = get_x_pos(end_dt)            # Current 結束 (最右)
                    p2 = get_x_pos(end_dt - seg)      # Current 開始 / L1 結束
                    p1 = get_x_pos(end_dt - 2 * seg)  # L1 開始 / L2 結束
                    p0 = get_x_pos(start_dt)          # L2 開始

                    # 強制讓 Current 延伸到最右邊界，解決懸空問題
                    p3 = x_max 

                    draw_windows = [
                        (p2, p3, 'Current', '#fee2e2'),
                        (p1, p2, 'L1', '#fef9c3'),
                        (p0, p1, 'L2', '#ede9fe'),
                    ]

                    text_trans = mtransforms.blended_transform_factory(ax.transData, ax.transAxes)
                    for xl, xr, lab, col in draw_windows:
                        # 確保區間有效且不是 NaN
                        if xl < xr and not np.isnan(xl) and not np.isnan(xr):
                            # 加入 linewidth=0 消除重疊處的深色條紋
                            ax.axvspan(xl, xr, color=col, alpha=0.2, zorder=0, linewidth=0)
                            mid_x = (xl + xr) / 2
                            if not np.isnan(mid_x):
                                ax.text(mid_x, 1.05, lab, transform=text_trans, ha='center', va='top', 
                                        fontsize=8, color='#374151', alpha=0.9)
                        
                except Exception as e:
                    print(f"[DEBUG] 底色繪製失敗: {e}")
            
            # 計算統計線
            usl = chart_info.get('USL', None)
            lsl = chart_info.get('LSL', None)
            target = None
            for key in ['Target', 'TARGET', 'TargetValue', '中心線', 'Center']:
                if key in chart_info and pd.notna(chart_info[key]):
                    target = chart_info[key]
                    break
            mean_val = float(np.mean(y)) if len(y) else None
            # 繪製點與線 (主數據線與超規點不加入 legend)
            # 找到原本 ax.plot 畫趨勢線的地方，將它設為底層 zorder
            ax.plot(x, y, linestyle='-', color='#d1d5db', linewidth=1, zorder=1) 

            # --- 新增：按 Tool 分色畫點 ---
            if 'ByTool' in plot_df.columns:
                try:
                    # 過濾掉 NaN 和無效的 Tool 值
                    plot_df_with_tool = plot_df.copy()
                    plot_df_with_tool['ByTool'] = plot_df_with_tool['ByTool'].astype(str)
                    plot_df_tools = plot_df_with_tool[plot_df_with_tool['ByTool'] != 'nan'].copy()
                    
                    if not plot_df_tools.empty:
                        tools = sorted(plot_df_tools['ByTool'].unique())
                        colors = ['#2563eb', '#dc2626', '#16a34a', '#f59e0b', '#7c3aed']
                        for i, tool in enumerate(tools):
                            mask = plot_df_with_tool['ByTool'] == tool
                            if mask.any():
                                x_masked = [x[j] for j in range(len(x)) if mask.iloc[j]]
                                y_masked = [y[j] for j in range(len(y)) if mask.iloc[j]]
                                if x_masked and y_masked:
                                    ax.scatter(x_masked, y_masked, 
                                               color=colors[i % len(colors)], label=str(tool), 
                                               s=35, zorder=3, edgecolors='white', linewidth=0.5)
                        if tools:  # 只有在有有效 Tool 時才顯示圖例
                            ax.legend(loc='upper left', fontsize=7, frameon=True, ncol=3, framealpha=0.9, edgecolor='#d1d5db')
                except Exception as e:
                    print(f"[WARN] 繪製 By Tool 點失敗: {e}")
            if usl is not None:
                ax.scatter([xi for xi, yi in zip(x, y) if yi > usl], [yi for yi in y if yi > usl], color='#dc2626', s=36, zorder=5, label='_nolegend_')
            if lsl is not None:
                ax.scatter([xi for xi, yi in zip(x, y) if yi < lsl], [yi for yi in y if yi < lsl], color='#dc2626', marker='s', s=36, zorder=5, label='_nolegend_')
            # 計算 y 範圍（納入 USL/LSL/Target/Mean）避免被裁切
            extra_vals = [v for v in [usl, lsl, target, mean_val]
                          if v is not None and not (isinstance(v, float) and np.isnan(v))]
            if len(y) > 0:
                ymin_sel = float(np.min(y))
                ymax_sel = float(np.max(y))
            else:
                ymin_sel, ymax_sel = (0.0, 1.0)
            if extra_vals:
                ymin_sel = min(ymin_sel, min(extra_vals))
                ymax_sel = max(ymax_sel, max(extra_vals))
            rng = ymax_sel - ymin_sel
            margin = 0.05 * rng if rng > 0 else 1.0
            ax.set_ylim(ymin_sel - margin, ymax_sel + margin)

            # 畫短水平線，並讓文字直接接在線的末端
            from matplotlib import transforms as mtransforms
            trans = mtransforms.blended_transform_factory(ax.transAxes, ax.transData)
            def segment_with_label(val, name, color, va='center'):
                if val is None or (isinstance(val, float) and np.isnan(val)):
                    return
                x0, x1 = 0.0, 0.965  # 線更長，文字更貼近右邊界
                ax.plot([x0, x1], [val, val], transform=trans, color=color, linestyle='--', linewidth=1.1)
                ax.text(x1, val, name, transform=trans, color=color, va=va, ha='left', fontsize=9)

            segment_with_label(usl, 'USL', '#ef4444', va='center')
            segment_with_label(lsl, 'LSL', '#ef4444', va='center')
            segment_with_label(target, 'Target', '#f59e0b', va='center')
            # segment_with_label(mean_val, 'Mean', '#16a34a', va='center')
            # 時間軸格式化
            if use_time_axis:
                try:
                    import matplotlib.dates as mdates
                    locator = mdates.AutoDateLocator(minticks=3, maxticks=8)
                    formatter = mdates.ConciseDateFormatter(locator)
                    ax.xaxis.set_major_locator(locator)
                    ax.xaxis.set_major_formatter(formatter)
                    for label in ax.get_xticklabels():
                        label.set_rotation(90)
                        label.set_ha('center')
                except Exception:
                    pass
            else:
                # 等距模式若有時間欄位，挑選部分刻度顯示對應日期字串
                if 'point_time' in plot_df.columns and not plot_df.empty:
                    times = plot_df['point_time'].tolist()
                    total = len(times)
                    # 減少標籤數量以避免水平顯示時重疊
                    if total <= 8:
                        tick_idx = list(range(1, total + 1))
                    else:
                        step = max(1, total // 6)  # 從 8 改為 6，顯示更少標籤
                        tick_idx = list(range(1, total + 1, step))
                        if tick_idx[-1] != total:
                            tick_idx.append(total)
                    labels = [times[i-1].strftime('%Y-%m-%d') for i in tick_idx]
                    ax.set_xticks(tick_idx)
                    # 改為水平顯示，左對齊以避免重疊
                    ax.set_xticklabels(labels, rotation=0, ha='center', fontsize=8)
            ax.grid(True, linestyle=':', linewidth=0.6, alpha=0.5)
            # 使用 subplots_adjust 手動控制邊距，讓圖表更往上頂
            self.figure.subplots_adjust(top=0.90, bottom=0.12, left=0.08, right=0.95)
            self.canvas.draw()
            
            # 繪製完成後呼叫子圖繪製
            self._draw_analysis_plots(raw_df, chart_info)
        except Exception as e:
            # 發生任何錯誤時，顯示錯誤訊息而非閃退
            print(f"[ERROR] 繪製圖表時發生錯誤: {e}")
            from translations import tr
            self.figure.clear()
            ax = self.figure.add_subplot(111)
            ax.text(0.5, 0.5, f"{tr('chart_error')}\n{str(e)}", ha='center', va='center', transform=ax.transAxes, color='red')
            self.canvas.draw()

    def _draw_analysis_plots(self, df, chart_info):
        """繪製下方兩張工具分析圖"""
        import scipy.stats as stats
        import pandas as pd
        import numpy as np
        
        self.fig_sub2.clear()
        self.fig_sub3.clear()

        # 自動偵測 Tool 欄位名稱
        tool_col = next((c for c in ['ByTool', 'tool_id', 'TOOL_ID', 'Tool', 'Machine', 'EQP_ID', '機台'] if c in df.columns), None)

        if df is not None and not df.empty and tool_col:
            # 資料預處理
            df_plot = df.copy()
            df_plot['point_val'] = pd.to_numeric(df_plot['point_val'], errors='coerce')
            if 'point_time' in df_plot.columns:
                df_plot['point_time'] = pd.to_datetime(df_plot['point_time'], errors='coerce')
                df_plot = df_plot.dropna(subset=['point_val', 'point_time'])
            else:
                df_plot = df_plot.dropna(subset=['point_val'])
            
            # 過濾掉 NaN 的 Tool 資料
            df_plot = df_plot[pd.notna(df_plot[tool_col])].copy()
            
            if df_plot.empty:
                # 若過濾後無數據則顯示空白提示
                for fig in [self.fig_sub2, self.fig_sub3]:
                    ax = fig.add_subplot(111)
                    ax.text(0.5, 0.5, "No Tool Info Found", ha='center', va='center')
                return
            
            tools = sorted(df_plot[tool_col].unique())
            color_cycle = ['#2563eb', '#dc2626', '#16a34a', '#f59e0b', '#7c3aed', '#0891b2']
            tool_colors = {t: color_cycle[i % len(color_cycle)] for i, t in enumerate(tools)}

            # --- 圖 2: Tool-by-Tool Comparison (左下：水平排排站 + 組內連線) ---
            ax2 = self.fig_sub2.add_subplot(111)
            ax2.set_title("Tool Variation Comparison", fontsize=10, fontweight='bold', pad=10)
            ax2.set_ylabel("Measured Value", fontsize=8)
            ax2.set_xlabel("Tool Groups", fontsize=8)

            group_width = 10 
            gap = 5
            for i, t in enumerate(tools):
                subset = df_plot[df_plot[tool_col] == t]
                if 'point_time' in subset.columns:
                    subset = subset.sort_values('point_time').reset_index(drop=True)
                else:
                    subset = subset.reset_index(drop=True)
                    
                start_x = i * (group_width + gap)
                x_internal = start_x + subset.index
                color = tool_colors[t]
                
                # 繪製組內連線與點（加上 label 以便顯示在 legend 中）
                ax2.plot(x_internal, subset['point_val'], color=color, alpha=0.3, linewidth=1, zorder=2, label='_nolegend_')
                ax2.scatter(x_internal, subset['point_val'], color=color, s=20, alpha=0.7, zorder=3, edgecolors='white', linewidth=0.5, label=str(t))

            ax2.set_xticks([])  # 隱藏 X 軸數字刻度
            ax2.grid(True, axis='y', linestyle=':', alpha=0.4)
            ax2.tick_params(labelsize=8)
            # 添加 legend
            ax2.legend(fontsize=7, loc='upper left', frameon=True, ncol=3, framealpha=0.9, edgecolor='#d1d5db')
            
            # 同步大圖 Y 軸範圍，確保對比機台 Bias 有意義
            if self.figure.axes:
                try:
                    ax2.set_ylim(self.figure.axes[0].get_ylim())
                except:
                    pass

            # --- 圖 3: Normal Probability Plot (右下：Q-Q 圖標籤與格式優化) ---
            ax3 = self.fig_sub3.add_subplot(111)
            ax3.set_title("Normal Probability Plot", fontsize=10, fontweight='bold', pad=10)
            ax3.set_xlabel("Theoretical Quantiles (σ)", fontsize=8)
            ax3.set_ylabel("Sample Quantiles (Value)", fontsize=8)

            for t in tools:
                tool_data = df_plot[df_plot[tool_col] == t]['point_val']
                if len(tool_data) > 3:
                    (osm, osr), (slope, intercept, r) = stats.probplot(tool_data, dist="norm")
                    # 繪製實測點
                    ax3.scatter(osm, osr, color=tool_colors[t], s=15, alpha=0.5, label=t, edgecolors='none')
                    # 繪製常態分佈基準線（直線）
                    ax3.plot(osm, slope*osm + intercept, color=tool_colors[t], alpha=0.2, linewidth=1)
            
            # 強制 Z-score 刻度，讓你一眼看出 3σ 邊界
            ax3.set_xticks([-3, -2, -1, 0, 1, 2, 3])
            ax3.grid(True, linestyle=':', alpha=0.4)
            ax3.tick_params(labelsize=8)
            ax3.legend(fontsize=7, loc='upper left', frameon=True, ncol=3, framealpha=0.9, edgecolor='#d1d5db')
        else:
            # 若無數據或無 Tool 資訊則顯示空白提示
            for fig in [self.fig_sub2, self.fig_sub3]:
                ax = fig.add_subplot(111)
                ax.text(0.5, 0.5, "No Tool Info Found", ha='center', va='center')

        # 最終佈局調整：確保 Title 和 Label 不會被裁切
        self.fig_sub2.subplots_adjust(left=0.18, bottom=0.22, right=0.95, top=0.85)
        self.fig_sub3.subplots_adjust(left=0.18, bottom=0.22, right=0.95, top=0.85)
        
        self.canvas_sub2.draw()
        self.canvas_sub3.draw()

    def prev_chart(self):
        """切換到上一張圖表"""
        current_idx = self.chart_combo.currentIndex()
        if current_idx > 0:
            self.chart_combo.setCurrentIndex(current_idx - 1)
    
    def next_chart(self):
        """切換到下一張圖表"""
        current_idx = self.chart_combo.currentIndex()
        if current_idx < self.chart_combo.count() - 1:
            self.chart_combo.setCurrentIndex(current_idx + 1)