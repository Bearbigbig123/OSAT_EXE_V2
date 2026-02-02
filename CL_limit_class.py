import pandas as pd
import numpy as np
import os

# === 設定 matplotlib 後端，確保打包環境下的兼容性 ===
import matplotlib
try:
    matplotlib.use('Agg')  # 使用 Agg 後端，適合無顯示環境
except:
    pass

import matplotlib.pyplot as plt
import matplotlib.font_manager as fm

from math import gcd, floor, ceil
from functools import reduce
from scipy.stats import skew, median_abs_deviation, kurtosis, norm, rankdata
import scipy.stats as stats

# 設定中文字體（添加異常處理）
try:
    plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei', 'SimHei', 'DejaVu Sans']
    plt.rcParams['axes.unicode_minus'] = False
except Exception as e:
    print(f"[Warning] 字體設定失敗，使用預設字體: {e}")
    plt.rcParams['font.family'] = 'DejaVu Sans'


def transform_johnson_slifker_shapiro_full(data):
    """
    修正版: 使用 Slifker-Shapiro (1980) 演算法進行 Johnson 轉換
    修正重點: 
    1. SU 使用正確的 arccosh 計算 delta
    2. SL 使用正確的 log(x-xi) 公式 (無分母)
    3. SB/其他情況 使用 Rank-based INT 避免崩潰
    """
    data = np.array(data)
    n = len(data)
    
    if n < 10: # 數據太少不適合用 Slifker
        # 回傳簡單的 Z-score
        return (data - np.mean(data)) / (np.std(data, ddof=1) + 1e-9), "Insufficient_Data"
    
    try:
        # Step 1: 計算 Percentiles (z = 0.524)
        z_val = 0.524
        cdf_vals = norm.cdf([-3*z_val, -z_val, z_val, 3*z_val])
        
        # 使用 linear 插值
        x_quants = np.percentile(data, cdf_vals * 100)
        x_neg_3z, x_neg_z, x_z, x_3z = x_quants[0], x_quants[1], x_quants[2], x_quants[3]
        
        # Step 2: 計算 m, n, p
        m = x_3z - x_z
        n_val = x_neg_z - x_neg_3z
        p = x_z - x_neg_z
        
        if p <= 0 or m <= 0 or n_val <= 0:
             raise ValueError("Invalid quantile distance")
        
        # Step 3: 計算 QR
        QR = (m * n_val) / (p**2)
        
        # Step 4: 策略分支
        
        # === Case A: Johnson SU (Unbounded) ===
        if QR > 1.05:
            system_type = "SU"
            
            # [修正] 使用 arccosh 計算 delta
            cosh_arg = 0.5 * (m/p + n_val/p)
            if cosh_arg < 1: cosh_arg = 1.0001
            
            eta = 2 * z_val
            delta = eta / np.arccosh(cosh_arg)
            
            # 計算 gamma
            sinh_arg = (n_val/p - m/p) / (2 * np.sqrt(QR - 1))
            gamma = delta * np.arcsinh(sinh_arg)
            
            # 計算 lambda 和 xi
            lambda_param = (2 * p * np.sqrt(QR - 1)) / (m/p + n_val/p - 2)
            xi = (x_z + x_neg_z) / 2 - (p * (n_val/p - m/p)) / (2 * (m/p + n_val/p - 2))
            
            if lambda_param <= 0: raise ValueError("Lambda <= 0")
            
            # [公式] SU: arcsinh
            transformed = gamma + delta * np.arcsinh((data - xi) / lambda_param)
            return transformed, system_type

        # === Case B: Johnson SL (Lognormal) ===
        elif 0.95 <= QR <= 1.05:
            system_type = "SL"
            
            # [修正] SL 參數估計公式
            eta = 2 * z_val
            # 這裡用 m/p 近似 n/p
            ratio = m / p 
            if ratio <= 1: ratio = 1.0001
            
            delta = eta / np.log(ratio)
            
            # 計算 xi (下限)
            # Slifker 對 SL 的 xi 估計:
            xi = 0.5 * (x_z + x_neg_z) - 0.5 * p * (ratio + 1) / (ratio - 1)
            
            # 計算 gamma
            # 從 z = gamma + delta * ln(x_z - xi) 反推
            if (x_z - xi) <= 0: raise ValueError("Invalid SL gamma param")
            gamma = z_val - delta * np.log(x_z - xi)
            
            # [安全性檢查] 確保所有數據都大於下限 xi
            safe_data = data - xi
            if np.any(safe_data <= 0):
                # 如果數據違背 SL 假設 (有值 <= 下限)，轉用 Rank
                raise ValueError("Data violates SL lower bound")
                
            # [公式] SL: log(X - xi)  <--- 注意：這裡沒有分母！
            transformed = gamma + delta * np.log(safe_data)
            return transformed, system_type

        else:
            # QR < 0.95 (SB) 或其他情況
            # 為了系統穩定，統一使用 Rank-based INT
            raise ValueError("QR indicates SB or Normal, fallback to Rank INT")
            
    except Exception as e:
        # Fallback: Rank-based Inverse Normal Transformation
        # 這是最安全的兜底方案
        system_type = "Rank_INT"
        ranks = rankdata(data, method='average')
        probabilities = (ranks - 0.375) / (n + 0.25)
        transformed = norm.ppf(probabilities)
        transformed = np.clip(transformed, -6, 6)
        
        return transformed, system_type


class CLTightenCalculator:
    """Control Limit Tighten Calculator - 管制線收緊計算器"""
    
    def __init__(self, chart_info_path=None, raw_data_dir=None, start_date=None, end_date=None):
        """
        初始化 CL Tighten Calculator
        
        Args:
            chart_info_path: Chart 資訊檔案路徑
            raw_data_dir: 原始數據目錄路徑
            start_date: 自訂起始日期 (datetime object)
            end_date: 自訂結束日期 (datetime object)
        """
        self.chart_info_path = chart_info_path
        self.raw_data_dir = raw_data_dir
        self.start_date = start_date
        self.end_date = end_date
        self.results = []
        
    # === Utility Functions ===
    
    def compute_resolution(self, values):
        """
        SOP 1.3: 估計資料解析度 (Resolution) - 工業級抗噪版
        策略：事前清理 (Pre-rounding) -> 整數 GCD -> 事後鎖定 (Post-rounding)
        """
        if len(values) < 2: 
            return None
        
        # 1. 排序並去重
        sorted_vals = sorted(list(set(values)))
        if len(sorted_vals) < 2: 
            return None
        
        # === 【關鍵步驟 1】事前清理：計算差值時直接濾除雜訊 ===
        # 使用 round(x, 10) 將 0.09999999999999987 強制修正為 0.1
        # 這能保證進入 GCD 的數字是乾淨的
        diffs = []
        for i, j in zip(sorted_vals[:-1], sorted_vals[1:]):
            diff = j - i
            # 如果差值極小（可能是浮點數誤差造成的 0），忽略它
            if diff < 1e-9:
                continue
            # 關鍵：在這裡就先修整數字
            diffs.append(round(diff, 10))
            
        # 去重，減少計算量
        unique_diffs = sorted(list(set(diffs)))
        if not unique_diffs: 
            return None
        
        # ========== 動態判斷放大倍率 ==========
        # 找出最小的差值，決定要放大多少倍才能變成整數
        min_val = unique_diffs[0]
        
        # 尋找最佳 scale_factor (讓所有差值變成整數的最小倍率)
        scale_factor = 1
        found_scale = False
        
        # 嘗試從 10^0 到 10^8
        for p in range(9):
            factor = 10 ** p
            # 檢查是否所有差值乘上 factor 後都接近整數 (誤差小於 1e-5)
            # 使用 1e-5 是因為前面已經 round 過了，這裡可以寬鬆一點
            if all(abs(d * factor - round(d * factor)) < 1e-5 for d in unique_diffs):
                scale_factor = factor
                found_scale = True
                break
        
        if not found_scale:
            # 如果找不到完美倍率，使用根據小數位數推算的倍率（最大保底）
            # 找出最多小數位數的數
            max_decimals = 0
            for val in unique_diffs:
                s = f"{val:.10f}".rstrip('0')
                if '.' in s:
                    max_decimals = max(max_decimals, len(s.split('.')[1]))
            scale_factor = 10 ** min(max_decimals, 8)

        # ========== 整數 GCD 計算 ==========
        try:
            # 放大並轉為整數
            scaled_diffs = [int(round(d * scale_factor)) for d in unique_diffs]
            
            if not scaled_diffs:
                return None
            
            # 計算 GCD
            res_scaled = reduce(gcd, scaled_diffs)
            
            # 縮小回浮點數
            resolution = res_scaled / scale_factor
            
            # 防呆：解析度不應大於最小差值
            if resolution > min_val:
                resolution = min_val
            
            # === 【關鍵步驟 2】事後鎖定：再次清理結果 ===
            # 因為除法可能又引入極微小的誤差，最後再一次 Round
            # 這裡就是你原本想要的「找到最接近」的動作
            resolution = round(resolution, 10)
            
            return resolution
            
        except Exception as e:
            print(f"    [Warning] compute_resolution 計算失敗: {e}，返回 None")
            return None

    def robust_zscore_sop2(self, values):
        """SOP 2.4/4.2: 計算 Robust Z-score (zi)"""
        med = np.median(values)
        sd = np.std(values, ddof=1)  # 使用樣本標準差
        if sd == 0: 
            return np.zeros(len(values))
            
        mad = median_abs_deviation(values) 
        
        if mad == 0:
            mean_dev = np.mean(np.abs(values - med))
            if mean_dev == 0: 
                return np.zeros(len(values))
            z = 0.7979 * np.abs(values - med) / mean_dev
        else:
            z = 0.6745 * np.abs(values - med) / mad
        return z

    def compute_CB(self, values):
        """標準 Bimodality Coefficient (BC)"""
        N = len(values)
        if N < 4: 
            return 0 
        sk = skew(values, bias=False)
        ku = kurtosis(values, fisher=True, bias=False)
        return (sk**2 + 1) / (ku+3)

    def compute_robust_sigma(self, values):
        """SOP 4.1: 依資料筆數計算 UR/LR (Robust Sigma)"""
        N = len(values)
        P = np.percentile
        median_val = np.median(values)
        
        if N < 4: 
            return np.std(values, ddof=1), np.std(values, ddof=1)
            
        # SOP 4.1 Robust Sigma 估計邏輯
        if N >= 10000:
            UR = (P(values, 99.9) - median_val)/3.09
            LR = (median_val - P(values, 0.1))/3.09
        elif 3000 <= N < 10000:
            UR = (P(values, 99.5) - median_val)/2.576
            LR = (median_val - P(values, 0.5))/2.576
        elif 300 <= N < 3000:
            UR = (P(values, 99) - median_val)/2.326
            LR = (median_val - P(values, 1))/2.326
        elif 100 <= N < 300:
            UR = (P(values, 97) - median_val)/1.881
            LR = (median_val - P(values, 3))/1.881
        else: 
            UR = (P(values, 95) - median_val)/1.645
            LR = (median_val - P(values, 5))/1.645
            
        # 確保不會返回負值
        return max(0, UR), max(0, LR)

    # === SOP 1: Data Integrity ===

    def data_integrity(self, df, date_col, value_col, oos_col):
        """SOP 1.1 & 1.2: 篩選有效資料並排除 OOS
        
        如果設定了自訂日期範圍 (self.start_date 和 self.end_date)，則使用自訂範圍
        否則使用預設的最近2年
        """
        df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
        
        # 判斷是否使用自訂日期範圍
        if self.start_date is not None and self.end_date is not None:
            # 使用自訂日期範圍
            cutoff_start = pd.Timestamp(self.start_date)
            cutoff_end = pd.Timestamp(self.end_date)
            df_filtered = df[(df[date_col] >= cutoff_start) & (df[date_col] <= cutoff_end)].dropna(subset=[value_col])
            print(f"    使用自訂日期範圍: {cutoff_start.date()} 至 {cutoff_end.date()}")
        else:
            # 使用預設的最近2年
            cutoff = pd.Timestamp.today() - pd.DateOffset(years=2)
            df_filtered = df[df[date_col] >= cutoff].dropna(subset=[value_col])
            print(f"    使用預設日期範圍: 最近2年 (從 {cutoff.date()} 起)")
        
        if oos_col in df_filtered.columns:
            # SOP 1.2: 排除 OOS 點
            df_filtered = df_filtered[~df_filtered[oos_col].astype(bool)]
        
        values = df_filtered[value_col].values
        
        # 檢查是否有有效數據
        if len(values) == 0:
            print(f"    [Warning] 經過篩選後沒有有效數據點")
            return values, None
            
        resolution = self.compute_resolution(values)
        return values, resolution

    def calculate_decimals_from_resolution(self, resolution):
        """✅ 正確計算 resolution 的小數位數（智能處理浮點誤差）"""
        from decimal import Decimal
        
        if resolution == 0:
            return 0
        elif isinstance(resolution, int):
            return 0
        
        try:
            # 🔧 [智能修正] 先檢查是否接近常見的 "乾淨" resolution 值
            # 常見值：1, 0.1, 0.01, 0.001, 0.0001, 5, 0.5, 0.05, 0.005, 0.0005
            common_resolutions = [
                (1, 0), (0.5, 1), (0.1, 1), (0.05, 2), (0.01, 2),
                (0.005, 3), (0.001, 3), (0.0005, 4), (0.0001, 4),
                (0.00005, 5), (0.00001, 5), (0.000001, 6)
            ]
            
            # 檢查是否非常接近某個常見值（容忍度 1e-10）
            for common_val, decimals in common_resolutions:
                if abs(resolution - common_val) < 1e-10:
                    print(f"    [Resolution 修正] {resolution:.18f} → {common_val} (decimals={decimals})")
                    return decimals
            
            # 如果不是常見值，使用 Decimal 進行精確計算
            res_decimal = Decimal(str(resolution))
            # as_tuple() 返回 (sign, digits, exponent)
            # exponent 為負值時表示小數位數
            sign, digits, exponent = res_decimal.as_tuple()
            
            if exponent >= 0:
                # 整數或科學記號，無小數位
                return 0
            else:
                # exponent 為負值，小數位數 = abs(exponent)
                return -exponent
        except Exception as e:
            print(f"Warning: Failed to calculate decimals from {resolution}: {e}")
            return 4  # 預設為 4 位小數

    def apply_resolution_precision(self, value, resolution, value_name="value"):
        """
        ✅ 統一的精度鎖定函數：根據 resolution 將數值對齊到正確的小數位數
        
        Args:
            value: 要處理的數值
            resolution: 資料解析度
            value_name: 數值名稱（用於 debug 輸出）
            
        Returns:
            對齊後的數值
        """
        if value is None or np.isnan(value):
            return value
            
        if resolution is None or resolution <= 0:
            return value  # 無效 resolution，保持原值
        
        try:
            decimals = self.calculate_decimals_from_resolution(resolution)
            old_value = float(value)
            new_value = round(old_value, decimals)
            
            # 如果是 0 位小數，轉為整數
            if decimals == 0:
                new_value = int(new_value)
            
            # Debug 輸出（只在有變化時）
            if abs(old_value - float(new_value)) > 1e-10:
                print(f"    [Precision] {value_name}: {old_value:.10f} → {new_value} (decimals={decimals})")
            
            return new_value
        except Exception as e:
            print(f"    [Warning] Failed to apply precision to {value_name}: {e}")
            return value

    # === Hard Rule Function ===

    def apply_discrete_hard_rules(self, values, resolution, N):
        if N < 4:
            return None, None, False, None 

        values = np.array(values)
        val_counts = pd.Series(values).value_counts()
        
        if val_counts.empty:
            return None, None, False, None

        mode_val = val_counts.idxmax()
        category_num = len(val_counts)
        
        # Rule 1: If Constant or non mode data# = 1
        non_mode_count = N - val_counts.max()
        if non_mode_count == 0 or non_mode_count == 1:
            return mode_val, mode_val, True, "Hard Rule 1: Constant/Near Constant"

        min_val = np.min(values)
        max_val = np.max(values)
        
        # Rule 2: If data value category# = 2
        if category_num == 2:
            return max_val, min_val, True, "Hard Rule 2: Two Categories"

        # Rule 3: If data value category# = 3 and max – min = 2*resolution
        if category_num == 3 and resolution is not None and resolution > 0:
            if abs((max_val - min_val) - 2 * resolution) < 1e-6: 
                return max_val, min_val, True, "Hard Rule 3: Three Categories Spaced by Resolution"
                
        return None, None, False, None

    # === SOP 2 & 3: Pattern Diagnosis ===

    def data_prep_for_pattern(self, values):
            """
            Pattern 識別前的數據預處理。
            修改紀錄: 將 Johnson SB fit (MLE) 替換為 Quantile Regression 以提升穩定性。
            """
            values = np.array(values) # 確保是 array
            N_orig = len(values)
            if N_orig == 0: 
                return values
            if N_orig < 4: 
                return values
                
            val_counts = pd.Series(values).value_counts()
            mode_val = val_counts.idxmax()
            mode_count = val_counts.max()
            non_mode_count = N_orig - mode_count
            
            # SOP 2.1: 數據選擇 (Mode Balancing)
            if non_mode_count / N_orig < 0.25:
                max_mode_use = min(mode_count, non_mode_count * 3)
                
                # 特殊情況：當所有數據都相同時（non_mode_count = 0），直接返回原數據
                if non_mode_count == 0:
                    print(f"    [Debug] data_prep_for_pattern: 所有數據都相同 ({mode_val})，直接返回原數據")
                    return values
                    
                mode_data = values[values == mode_val][:max_mode_use]
                non_mode_data = values[values != mode_val]
                w = np.concatenate((mode_data, non_mode_data))
            else:
                w = values.copy()
                
            # 檢查 w 是否為空
            if len(w) == 0:
                print(f"    [Debug] data_prep_for_pattern: 預處理後數據為空，返回原數據")
                return values
            
            # SOP 2.3: Johnson Transformation using Slifker-Shapiro Method
            try:
                y, transform_type = transform_johnson_slifker_shapiro_full(w)
                print(f"    [Debug] 使用 {transform_type} 轉換方法")
            except Exception as e:
                print(f"    [Debug] 轉換失敗，使用原始數據: {e}")
                y = w  # 若連 Rank 都失敗(極少見)，回退原數據
                
            # SOP 2.4: Robust Z-score
            z = self.robust_zscore_sop2(y)
            
            # SOP 2.5: 邊界過濾法 (使用 SOP 定義的 6.0 閾值)
            # 注意：這裡使用 6.0 (原程式碼設定)，與 outlier_filter 的 4.5 不同
            normal_mask = z <= 6.0
            
            if not np.any(normal_mask):
                return values  # 全被判異常，不過濾
            
            # 用 w 的正常範圍篩選原始 values
            min_normal = np.min(w[normal_mask])
            max_normal = np.max(w[normal_mask])
            
            # 用邊界篩選原始 values
            values_filtered = values[(values >= min_normal) & (values <= max_normal)]
            
            # 檢查濾除比例（用原始 values 計算）
            filter_ratio = (N_orig - len(values_filtered)) / N_orig
            
            if filter_ratio > 0.05:
                # 超過 5%，只濾除具有最大 Z-score 的值（w 中的最極端值）
                max_z_value = np.max(z)
                max_z_mask = (z == max_z_value)
                max_outlier_vals = w[max_z_mask]  # 所有最大 ZI 對應的 w 值
                unique_max_vals = np.unique(max_outlier_vals)  # 去重
                
                # 從原始數據中移除所有這些 w 值
                # 注意：這會移除 values 中所有等於這些值的點，即使同值有多個
                values_filtered = values.copy()
                for outlier_val in unique_max_vals:
                    values_filtered = values_filtered[values_filtered != outlier_val]
            
            return values_filtered

    def pattern_diagnosis(self, values, resolution=None):
        N = len(values)
        if N < 4: 
            return "Insufficient Data", np.nan, np.nan
        
        sigma = np.std(values, ddof=1)
        val_counts = pd.Series(values).value_counts()
        category_num = len(val_counts)
        sk = skew(values, bias=False)
        cb = self.compute_CB(values)
        
        if 4 <= N < 20: 
            if sigma == 0: 
                return "Constant", sk, cb
            
            near_constant_cond = (N - val_counts.max() == 1)
            attribute_cond = (category_num / N <= 1/3 and category_num <= 5)
            
            if near_constant_cond: 
                return "Near Constant", sk, cb
            if attribute_cond: 
                return "Attribute", sk, cb
            if category_num / N >= 1/2:
                if sk > 0.6: 
                    return "Skew-Right", sk, cb
                if sk < -0.6: 
                    return "Skew-Left", sk, cb
            return "Normal", sk, cb
        
        if N >= 20: 
            if sigma == 0: 
                return "Constant", sk, cb
                
            near_constant_cond = (val_counts.max() / N >= 0.9)
            
            attr_cond_A = (category_num / N <= 1/3 and category_num <= 5)
            attr_cond_B = (resolution and sigma != 0 and resolution / (6*sigma) >= 0.1 and category_num < 10)
            attr_cond_C = (N >= 30 and category_num <= 10)
            attribute_cond = (attr_cond_A or attr_cond_B or attr_cond_C)
            
            if near_constant_cond: 
                return "Near Constant", sk, cb
            if attribute_cond: 
                return "Attribute", sk, cb
                
            if cb > 0.6 and -0.6 <= sk <= 0.6: 
                return "Bimodal", sk, cb
                
            # 判斷偏斜
            if sk > 0.6:
                return "Skew-Right", sk, cb
            elif sk < -0.6:
                return "Skew-Left", sk, cb

            # 其餘情況視為正常分布
            return "Normal", sk, cb

    # === SOP 4: Outlier Exclusion ===
    def outlier_filter(self, values, pattern):
            values = np.array(values)
            N_orig = len(values)
            if N_orig == 0 or np.std(values, ddof=1) == 0: 
                return values
            
            print(f"\n    ===== Outlier Filter Debug (Pattern: {pattern}) =====")
            print(f"    原始數據筆數: {N_orig}")
            print(f"    原始數據範圍: [{np.min(values):.6f}, {np.max(values):.6f}]")
            print(f"    原始數據統計: Mean={np.mean(values):.6f}, Median={np.median(values):.6f}, Std={np.std(values, ddof=1):.6f}")
            
            if pattern in ["Skew-Right","Skew-Left"]:
                # (這部分保持原本邏輯)
                print(f"\n    --- Skew Pattern 濾除流程 ---")
                median_val = np.median(values)
                U_R, L_R = self.compute_robust_sigma(values)
                
                print(f"    Median: {median_val:.6f}")
                print(f"    Robust Sigma: U_R={U_R:.6f}, L_R={L_R:.6f}")
                print(f"    Upper Threshold: {median_val + 6*U_R:.6f}")
                print(f"    Lower Threshold: {median_val - 6*L_R:.6f}")
                
                mask = (values <= median_val + 6*U_R) & (values >= median_val - 6*L_R)
                values_pre_filtered = values[mask]
                
                outliers = values[~mask]
                if len(outliers) > 0:
                    print(f"    被 6σ 濾除的點 ({len(outliers)}): {sorted(outliers)[:10]}..." if len(outliers) > 10 else f"    被 6σ 濾除的點 ({len(outliers)}): {sorted(outliers)}")
                else:
                    print(f"    無點被 6σ 濾除")
                
                filter_ratio = (N_orig - len(values_pre_filtered)) / N_orig
                print(f"    濾除比例: {filter_ratio*100:.2f}% ({N_orig - len(values_pre_filtered)}/{N_orig})")
                
                if filter_ratio <= 0.05:  # 5% 以內，使用濾除後的結果
                    print(f"    ✓ 濾除比例 ≤ 5%，使用濾除後的結果")
                    print(f"    最終保留數據筆數: {len(values_pre_filtered)}")
                    print(f"    最終數據範圍: [{np.min(values_pre_filtered):.6f}, {np.max(values_pre_filtered):.6f}]")
                    return values_pre_filtered
                else:  # 超過 5%，只濾除所有最大比例的點
                    print(f"    ✗ 濾除比例 > 5%，改用最大比例濾除法")
                    ratio_upper = (values - median_val) / U_R
                    ratio_lower = (median_val - values) / L_R
                    ratio = np.maximum(ratio_upper, ratio_lower) 
                    max_ratio_value = np.max(ratio)
                    max_ratio_mask = ratio == max_ratio_value
                    max_outliers = values[max_ratio_mask]
                    
                    print(f"    最大比例值: {max_ratio_value:.4f}")
                    print(f"    最大比例對應的點 ({len(max_outliers)}): {sorted(max_outliers)}")
                    
                    values_filtered = values[~max_ratio_mask]
                    print(f"    最終保留數據筆數: {len(values_filtered)}")
                    print(f"    最終數據範圍: [{np.min(values_filtered):.6f}, {np.max(values_filtered):.6f}]")
                    return values_filtered
            
            else:
                # SOP 4.2: Other Pattern - Johnson Transformation + Boundary Filter Method
                print(f"\n    --- Other Pattern Johnson 轉換 + 邊界過濾法 ---")
                
                val_counts = pd.Series(values).value_counts()
                mode_val = val_counts.idxmax()
                mode_count = val_counts.max()
                non_mode_count = N_orig - mode_count
                
                print(f"    Mode 值: {mode_val:.6f} (出現 {mode_count} 次)")
                print(f"    Non-Mode 數量: {non_mode_count}")
                print(f"    Non-Mode 比例: {non_mode_count/N_orig*100:.2f}%")
                
                # ========== Step 1: Mode 平衡處理 (SOP 2.1) ==========
                if non_mode_count / N_orig < 0.25:
                    print(f"\n    Step 1: Mode 平衡處理 (Non-Mode < 25%)")
                    max_mode_use = min(mode_count, non_mode_count * 3) if non_mode_count > 0 else mode_count
                    
                    if non_mode_count == 0:
                        print(f"    ⚠ 所有數據相同，無需濾除")
                        return values
                    
                    print(f"    Mode 使用數量: {max_mode_use} (原 {mode_count})")
                    print(f"    Non-Mode 使用數量: {non_mode_count}")
                    
                    mode_data = values[values == mode_val][:max_mode_use]
                    non_mode_data = values[values != mode_val]
                    w = np.concatenate((mode_data, non_mode_data))
                    
                    print(f"    平衡後 w 數量: {len(w)} (Mode: {len(mode_data)}, Non-Mode: {len(non_mode_data)})")
                else:
                    print(f"\n    Step 1: 無需 Mode 平衡 (Non-Mode ≥ 25%)")
                    w = values.copy()
                
                if len(w) == 0:
                    print(f"    ⚠ w 為空，無需濾除")
                    return values
                
                print(f"    w 數據範圍: [{np.min(w):.6f}, {np.max(w):.6f}]")
                
                # ========== Step 2: [修改] 移除 Min-Max 標準化，直接使用 Slifker-Shapiro ==========
                print(f"\n    Step 2: Johnson Transformation using Slifker-Shapiro Method")
                
                try:
                    y, transform_type = transform_johnson_slifker_shapiro_full(w)
                    print(f"    [Debug-Johnson] 使用 {transform_type} 轉換方法")
                    print(f"    [Debug-Johnson] y (Z-score) 範圍 = [{np.min(y):.4f}, {np.max(y):.4f}]")
                except Exception as e:
                    print(f"    [Error-Johnson] Johnson 轉換失敗: {e}，回退到原始 w")
                    y = w
                
                # ========== Step 3: Robust Z-score (SOP 2.4) ==========
                print(f"\n    Step 3: Robust Z-score 計算")
                z = self.robust_zscore_sop2(y)
                print(f"    Robust Z-score 範圍: [{np.min(z):.4f}, {np.max(z):.4f}]")
                print(f"    Z-score 前 10 個值: {z[:10]}")
                
                max_z_idx = np.argmax(z)
                print(f"    最大 Z-score: {z[max_z_idx]:.4f} (索引 {max_z_idx})")
                print(f"    對應 w 值: {w[max_z_idx]:.6f}")
                
                # ========== Step 4: 邊界過濾法 (SOP 2.5) ==========
                print(f"\n    Step 4: 邊界過濾法")
                normal_mask = z <= 4.5
                outlier_count = (~normal_mask).sum()
                
                print(f"    Z > 4.5 的異常點數量: {outlier_count}")
                
                if outlier_count > 0:
                    outlier_z = z[~normal_mask]
                    outlier_w = w[~normal_mask]
                    print(f"    異常點 Z-scores: {outlier_z}")
                    print(f"    異常點對應 w 值: {outlier_w}")
                
                if not np.any(normal_mask):
                    print(f"    ⚠ 所有點都被判為異常，不進行濾除")
                    return values 
                
                # ✅ 用 w 的正常範圍來過濾原始 values
                min_normal = np.min(w[normal_mask])
                max_normal = np.max(w[normal_mask])
                
                print(f"    正常數據邊界: [{min_normal:.6f}, {max_normal:.6f}]")
                
                values_filtered = values[(values >= min_normal) & (values <= max_normal)]
                filtered_out = values[(values < min_normal) | (values > max_normal)]
                
                print(f"    原始數據中被濾除的點 ({len(filtered_out)}): {sorted(filtered_out)}")
                
                filter_ratio = (N_orig - len(values_filtered)) / N_orig
                print(f"    濾除比例: {filter_ratio*100:.2f}% ({N_orig - len(values_filtered)}/{N_orig})")
                
                if filter_ratio > 0.05:
                    print(f"    ✗ 濾除比例 > 5%，改用最大 Z-score 濾除法")
                    # 找出所有具有最大 Z-score 的 w 值
                    max_z_value = np.max(z)
                    max_z_mask = (z == max_z_value)
                    max_outlier_vals = w[max_z_mask]  # 所有最大 ZI 對應的 w 值
                    unique_max_vals = np.unique(max_outlier_vals)  # 去重
                    
                    print(f"    最大 Z-score: {max_z_value:.4f}")
                    print(f"    具有最大 ZI 的 w 值 ({len(unique_max_vals)}): {unique_max_vals}")
                    
                    # 從原始數據中移除所有這些 w 值
                    values_filtered = values.copy()
                    for outlier_val in unique_max_vals:
                        max_outlier_mask = values_filtered == outlier_val
                        max_outlier_points = values_filtered[max_outlier_mask]
                        if len(max_outlier_points) > 0:
                            print(f"    濾除 w={outlier_val:.6f} 的點 ({len(max_outlier_points)} 個)")
                        values_filtered = values_filtered[~max_outlier_mask]
                else:
                    print(f"    ✓ 濾除比例 ≤ 5%，使用邊界濾除結果")
                
                print(f"    最終保留數據筆數: {len(values_filtered)}")
                print(f"    最終數據範圍: [{np.min(values_filtered):.6f}, {np.max(values_filtered):.6f}]")
                print(f"    ===== Outlier Filter 完成 =====\n")
                
                return values_filtered

    # === SOP 5: Control Limit Calculation ===

    def get_k_value(self, N, characteristic, pattern='Normal', kurtosis_value=None):
        """
        SOP 5.2: 依資料筆數和特性決定 k 值
        
        ✅ 特殊邏輯：當 Kurtosis > 1 且 Pattern = Normal 時，各加 1 sigma
        - N >= 30: 3σ → 4σ
        - 16 <= N <= 29: 4σ → 5σ
        - 4 <= N <= 15: 5σ → 6σ
        """
        # 基礎 k 值
        if N >= 30:
            base_k = 3.0
        elif 16 <= N <= 29: 
            base_k = 4.0
        elif 4 <= N <= 15: 
            base_k = 5.0
        else:
            base_k = 3.0
        
        # 🔥 特殊邏輯：Kurtosis > 1 且 Pattern = Normal 時，加 1σ
        if pattern == 'Normal' and kurtosis_value is not None and kurtosis_value > 1:
            final_k = base_k + 1.0  # 各加 1σ：3→4, 4→5, 5→6
            print(f"    [Kurtosis +1σ] Pattern={pattern}, Kurtosis={kurtosis_value:.3f} > 1")
            print(f"    [Kurtosis +1σ] N={N}, Base k={base_k:.1f} → Final k={final_k:.1f}")
            return final_k
        
        return base_k

    def calc_CL(self, values, pattern, resolution=None, characteristic='Nominal', kurtosis_value=None):
        """SOP 5.1 & 5.3: 計算 UCL/LCL"""
        N = len(values)
        if N < 4: 
            return np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan
            
        median_val = np.median(values)
        mean_val = np.mean(values)
        std_val = np.std(values, ddof=1)
        k = self.get_k_value(N, characteristic, pattern, kurtosis_value)
        
        # 預先計算 Robust Sigma 和 ECDF 3-sigma 備用值
        UR_robust, LR_robust = self.compute_robust_sigma(values)
        UCL3_ecdf, LCL3_ecdf = np.nan, np.nan 
        
        # --- 1. 處理 Normal 模式 (Mean +/- k*Std) ---
        if pattern == "Normal": 
            UCL = mean_val + k * std_val
            LCL = mean_val - k * std_val
            center_line = mean_val
            sigma_for_spec = std_val
        
        # --- 2. 處理 Skew 模式 (Median +/- k*RobustSigma) ---
        elif pattern in ["Skew-Right", "Skew-Left"]: 
            UCL = median_val + k * UR_robust
            LCL = median_val - k * LR_robust
            center_line = median_val
            # Skew 模式：上下分開使用各自的 sigma（不取最大值）
            sigma_upper = UR_robust
            sigma_lower = LR_robust
            
        # --- 3. 處理 ECDF 相關模式 (Bimodal, Attribute, Constant, Near Constant) ---
        else:
            # 3.1 ECDF: 強制使用 ECDF 確定基礎 3-sigma 水位 (UCL3, LCL3)
            p_low_3sigma = 0.135
            p_high_3sigma = 99.865
            
            try:
                UCL3_ecdf = np.percentile(values, p_high_3sigma)
                LCL3_ecdf = np.percentile(values, p_low_3sigma)
            except Exception:
                UCL3_ecdf = median_val + 3 * UR_robust
                LCL3_ecdf = median_val - 3 * LR_robust

            # 3.2 Tolerance 擴展邏輯 (修改：上下分別記住各自的 tolerance)
            T_upper = UCL3_ecdf - median_val
            T_lower = median_val - LCL3_ecdf
            
            # 上下各自根據 k 倍數放大
            if T_upper > 0:
                T_upper_new = T_upper * (k / 3.0)
                UCL = median_val + T_upper_new
            else:
                UCL = UCL3_ecdf
            
            if T_lower > 0:
                T_lower_new = T_lower * (k / 3.0)
                LCL = median_val - T_lower_new
            else:
                LCL = LCL3_ecdf
            
            center_line = median_val
            # 對於ECDF模式，使用T_upper和T_lower的平均值除以3作為sigma
            sigma_for_spec = (T_upper + T_lower) / 6.0 if (T_upper > 0 and T_lower > 0) else max(UR_robust, LR_robust)
        
        # === 新需求1: 計算 Sug USL/LSL (center line ± 6*sigma) ===
        # Skew 模式：上下分開計算（使用各自的 sigma）
        if pattern in ["Skew-Right", "Skew-Left"]:
            Sug_USL = center_line + 6 * sigma_upper
            Sug_LSL = center_line - 6 * sigma_lower
        # Normal 和 ECDF 模式：使用統一的 sigma
        else:
            Sug_USL = center_line + 6 * sigma_for_spec
            Sug_LSL = center_line - 6 * sigma_for_spec
        
        # 返回原始計算值（不在此處做 resolution 調整）
        return UCL, LCL, UR_robust, LR_robust, UCL3_ecdf, LCL3_ecdf, Sug_USL, Sug_LSL

    # === SOP 6: Control Limit Adjustment & Tighten ===

    def adjust_CL_based_on_OOC(self, values, UCL, LCL, pattern, resolution, sigma_est_u, sigma_est_l, max_adj_units=2, characteristic='Nominal'):
        values = np.array(values)  # 确保 values 是 numpy array
        current_UCL, current_LCL = UCL, LCL
        N = len(values)
        
        print(f"    [Debug OOC] === 開始 OOC 退格調整 ===")
        print(f"    [Debug OOC] 輸入數據長度 N: {N}")
        print(f"    [Debug OOC] 初始 UCL: {UCL:.6f}, LCL: {LCL:.6f}")
        print(f"    [Debug OOC] Pattern: {pattern}, Characteristic: {characteristic}")
        print(f"    [Debug OOC] Resolution: {resolution}")
        print(f"    [Debug OOC] Sigma_est_u: {sigma_est_u:.6f}, Sigma_est_l: {sigma_est_l:.6f}")
        print(f"    [Debug OOC] 數據範圍: min={np.min(values):.6f}, max={np.max(values):.6f}")
        
        sigma_u = sigma_est_u if sigma_est_u > 1e-9 else 1e-9
        sigma_l = sigma_est_l if sigma_est_l > 1e-9 else 1e-9

        if pattern == "Constant":
            adj_u = adj_l = 0.0
        elif pattern in ["Near Constant", "Attribute"]:
            # Discrete 模式：每次加 1 resolution，無固定迭代次數上限
            if resolution is not None and resolution > 0:
                adj_u = adj_l = resolution
            else:
                adj_u = adj_l = 0.0
        else:
            # Continuous 模式 (Normal, Skew-Right, Skew-Left, Bimodal)：每次 0.25σ
            adj_u = 0.25 * sigma_u
            adj_l = 0.25 * sigma_l
        
        print(f"    [Debug OOC] 調整量 adj_u: {adj_u:.6f}, adj_l: {adj_l:.6f}")
            
        # 根據特性類型決定是否只調整單邊
        adjust_ucl = True
        adjust_lcl = True
        
        if characteristic == 'Smaller':
            # Smaller 只需要 tighten UCL (上限)
            adjust_lcl = False
            adj_l = 0.0  # 不調整 LCL
        elif characteristic == 'Bigger':
            # Bigger 只需要 tighten LCL (下限)
            adjust_ucl = False
            adj_u = 0.0  # 不調整 UCL
        
        print(f"    [Debug OOC] adjust_ucl: {adjust_ucl}, adjust_lcl: {adjust_lcl}")
            
        # 設定最大迭代次數（安全上限，防止無限循環）
        max_iterations = 100
        
        initial_ooc_count = 0
        final_ooc_count = 0
        total_adj_units = 0
        
        # 記錄退格前的初始值（用於計算累積退格量）
        initial_UCL = current_UCL
        initial_LCL = current_LCL

        for i in range(max_iterations):
            upper_ooc_mask = (values > current_UCL)
            lower_ooc_mask = (values < current_LCL)
            
            upper_ooc_count = np.sum(upper_ooc_mask)
            lower_ooc_count = np.sum(lower_ooc_mask)
            
            #String --- [邏輯修正] 根據特性只計算「有效」的 OOC ---
            # 原因：避免 Smaller 特性時，因無法調整 LCL，導致下界 OOC 讓迴圈無法滿足停止條件而空轉 100 次
            if characteristic == 'Smaller':
                # 望小特性：只關注 Upper OOC (過大)，忽略 Lower OOC
                total_ooc_count = upper_ooc_count
                
                # Debug 提示：如果有很多 Lower OOC 但被忽略
                if lower_ooc_count > 0:
                    print(f"    [Debug OOC] Smaller 特性忽略 {lower_ooc_count} 個 Lower OOC")
                    
            elif characteristic == 'Bigger':
                # 望大特性：只關注 Lower OOC (過小)，忽略 Upper OOC
                total_ooc_count = lower_ooc_count
                
                # Debug 提示
                if upper_ooc_count > 0:
                    print(f"    [Debug OOC] Bigger 特性忽略 {upper_ooc_count} 個 Upper OOC")
                    
            else:
                # Nominal (望目) 或其他：上下界 OOC 都算
                total_ooc_count = upper_ooc_count + lower_ooc_count
            # ----------------------------------------------------
            
            if i == 0: 
                initial_ooc_count = total_ooc_count
                print(f"    [Debug OOC] 初始 OOC count: {initial_ooc_count} (upper: {upper_ooc_count}, lower: {lower_ooc_count})")
            
            ooc_percent = total_ooc_count / N
            
            print(f"    [Debug OOC] 迭代 {i+1}: UCL={current_UCL:.6f}, LCL={current_LCL:.6f}, OOC={total_ooc_count} ({ooc_percent*100:.2f}%)")
            
            # 停止條件 1：OOC% ≤ 0.3% 或 OOC 點數 < 2
            if ooc_percent <= 0.003 or total_ooc_count < 2: 
                final_ooc_count = total_ooc_count
                print(f"    [Debug OOC] 迭代 {i+1}: 達到停止條件 (OOC≤0.3% 或 <2點)，停止調整")
                break
            
            # 執行退格前，檢查累積退格量是否會超過 ±2σ 限制
            should_stop = False
            
            if upper_ooc_count > 0 and adj_u > 0 and adjust_ucl:
                # 計算退格後的累積量
                cumulative_adj_u = (current_UCL + adj_u) - initial_UCL
                
                # 檢查是否超過 +2σ
                if cumulative_adj_u > 2 * sigma_est_u:
                    print(f"    [Debug] 迭代 {i+1}: UCL 累積退格 {cumulative_adj_u:.4f} 超過 +2σ ({2*sigma_est_u:.4f})，停止調整")
                    should_stop = True
                else:
                    current_UCL += adj_u
                    if pattern in ["Near Constant", "Attribute"]:
                        total_adj_units += 1
                    else:
                        total_adj_units += adj_u / sigma_u
            
            if lower_ooc_count > 0 and adj_l > 0 and adjust_lcl:
                # 計算退格後的累積量
                cumulative_adj_l = initial_LCL - (current_LCL - adj_l)
                
                # 檢查是否超過 -2σ
                if cumulative_adj_l > 2 * sigma_est_l:
                    print(f"    [Debug] 迭代 {i+1}: LCL 累積退格 {cumulative_adj_l:.4f} 超過 -2σ ({2*sigma_est_l:.4f})，停止調整")
                    should_stop = True
                else:
                    current_LCL -= adj_l
                    if pattern in ["Near Constant", "Attribute"]:
                        total_adj_units += 1
                    else:
                        total_adj_units += adj_l / sigma_l
            
            # 停止條件 2：達到 ±2σ 上限
            if should_stop:
                final_ooc_count = total_ooc_count
                break
            
            # 🌀 [Bug Fix] 停止條件 3：沒有「有效的」OOC 了（修正單邊規格死邏輯）
            # 原邏輯：if upper_ooc_count == 0 and lower_ooc_count == 0 在單邊規格下永遠無法滿足
            # 新邏輯：檢查 total_ooc_count（已根據特性過濾過）
            if total_ooc_count == 0:
                final_ooc_count = 0
                print(f"    [Debug OOC] 迭代 {i+1}: 有效 OOC = 0，停止調整")
                break
        
        # 最終硬性限制：確保 Suggest CL 不超過 Static CL ± 2σ
        max_ucl_allowed = initial_UCL + 2 * sigma_est_u
        min_lcl_allowed = initial_LCL - 2 * sigma_est_l
        
        if current_UCL > max_ucl_allowed:
            print(f"    [Warning] UCL 超過 Static+2σ 上限 ({current_UCL:.4f} > {max_ucl_allowed:.4f})，強制限制")
            current_UCL = max_ucl_allowed
        
        if current_LCL < min_lcl_allowed:
            print(f"    [Warning] LCL 超過 Static-2σ 下限 ({current_LCL:.4f} < {min_lcl_allowed:.4f})，強制限制")
            current_LCL = min_lcl_allowed

        # SOP 6.3: 最終 CL 根據 resolution 修正
        if resolution is not None and not np.isnan(current_UCL) and not np.isnan(current_LCL):
            # ✅ 使用新的方法精確計算 decimals（避免 rstrip('0') 的問題）
            decimals = self.calculate_decimals_from_resolution(resolution)
            
            if decimals >= 0:
                power_of_10 = 10**decimals
                print(f"    [DEBUG SOP 6.3] resolution={resolution}, decimals={decimals}, power_of_10={power_of_10}")
                print(f"    [DEBUG SOP 6.3] Before: UCL={current_UCL:.8f}, LCL={current_LCL:.8f}")
                
                # 🛑 [Bug Fix] 加入 epsilon 避免浮點精度陷阱
                # 防止 10.4999...99 被 floor 無條件捨去成 10.4
                epsilon = 1e-9
                current_UCL = floor(current_UCL * power_of_10 + epsilon) / power_of_10
                current_LCL = ceil(current_LCL * power_of_10 - epsilon) / power_of_10
                
                # ✅ [關鍵鎖定] 再次 round 以徹底消除 Python 浮點數微幅雜訊 (如 20.22000000001)
                current_UCL = round(float(current_UCL), decimals)
                current_LCL = round(float(current_LCL), decimals)
                
                print(f"    [DEBUG SOP 6.3] After: UCL={current_UCL:.8f}, LCL={current_LCL:.8f}")

            if current_LCL > current_UCL: 
                current_LCL = current_UCL
        
        print(f"    [Debug OOC] === OOC 退格調整結束 ===")
        print(f"    [Debug OOC] Raw UCL (退格後，capping 前): {current_UCL:.6f}, LCL: {current_LCL:.6f}")
        print(f"    [Debug OOC] Static OOC count: {initial_ooc_count}")
        print(f"    [Debug OOC] ⚠️ Final OOC count 將在 capping 後重新計算")
        print(f"    [Debug OOC] Total adjustment units: {total_adj_units:.4f}")

        # ⚠️ 注意：這裡返回的 final_ooc_count 是暫時值，會在 capping 後重新計算
        final_ooc_count = 0  # 暫時設為 0，實際值將在主函數中重新計算
        return current_UCL, current_LCL, initial_ooc_count, final_ooc_count, total_adj_units

    def check_tighten(self, original_tol, new_tol, data_count):
        """SOP 6.4: 判斷是否需要 tighten (容差比對)"""
        tighten_flag, _, _ = self.check_tighten_with_details(original_tol, new_tol, data_count)
        return tighten_flag
    
    def check_tighten_with_details(self, original_tol, new_tol, data_count):
        """SOP 6.4: 判斷是否需要 tighten (容差比對) - 返回詳細資訊"""
        N_pct_table = [(125,15),(70,18),(45,20),(30,25),(15,30),(10,35),(0,40)]
        N_pct = 40 
        for n_threshold, pct in N_pct_table:
            if data_count > n_threshold:
                N_pct = pct
                break
        
        if original_tol <= 0: 
            return False, np.nan, N_pct
        
        # 計算變化率（正值表示收緊，負值表示放寬）
        diff_ratio = (original_tol - new_tol) / original_tol * 100
        
        # 邏輯：只有當容差收緊（new_tol < original_tol）且變化率 > N% 時，才判定為 TightenNeeded
        # 如果放寬（new_tol > original_tol），則 diff_ratio 為負值，不會觸發 tighten
        tighten_flag = diff_ratio > N_pct
        
        return tighten_flag, diff_ratio, N_pct

    # === 核心流程包裝 ===
    
    def process_chart(self, df, value_col, date_col, oos_col, characteristic):
        """主體 SOP 流程 (SOP 1-6)"""
        
        # 1. Data Integrity (SOP 1)
        values_orig, resolution = self.data_integrity(df.copy(), date_col, value_col, oos_col)
        
        # 檢查是否有有效數據
        if len(values_orig) == 0:
            return {
                "Pattern": "No Valid Data",
                "Skew": np.nan,
                "CB": np.nan,
                "Resolution_Estimated": None,
                "Suggest UCL": np.nan,
                "Suggest LCL": np.nan,
                "Static UCL": np.nan,
                "Static LCL": np.nan,
                "Raw UCL After Resolution": np.nan,
                "Raw LCL After Resolution": np.nan,
                "Sug USL": np.nan,
                "Sug LSL": np.nan,
                "TightenNeeded": False,
                "TotalDataCount": 0,
                "DataCountUsed": 0,
                "HardRule": "None",
                "DetectionLimit": np.nan,
                "CL_Center": np.nan,
                "Sigma_Est": 0.0,
                "Sigma_Est_Upper": 0.0,
                "Sigma_Est_Lower": 0.0,
                "Original_UCL_K_Set": np.nan,
                "Original_LCL_K_Set": np.nan,
                "Suggest_UCL_K_Set": np.nan,
                "Suggest_LCL_K_Set": np.nan,
                "Ori_K_Set": np.nan,
                "Sug_K_Set": np.nan,
                "Total_Adj_Units": 0.0,
                "Initial_OOC_Count": 0,
                "Final_OOC_Count": 0,
                "Original_Tolerance": np.nan,
                "New_Tolerance": np.nan,
                "Diff_Ratio_%": np.nan,
                "Tighten_Threshold_%": np.nan
            }
        
        # 檢查數據點是否 < 4 個
        if len(values_orig) < 4:
            print(f"    [Warning] 數據點不足 ({len(values_orig)} < 4)，跳過計算")
            
            # 讀取必要參數
            detection_limit = df['DetectionLimit'].iloc[0] if 'DetectionLimit' in df.columns and len(df)>0 else np.nan
            original_ucl = df['UCL'].iloc[0] if 'UCL' in df.columns and len(df)>0 else np.nan
            original_lcl = df['LCL'].iloc[0] if 'LCL' in df.columns and len(df)>0 else np.nan
            
            # 計算 Ori OOC Count
            ori_ooc_count = 0
            if not (pd.isna(original_ucl) or pd.isna(original_lcl)):
                ori_upper_ooc = np.sum(values_orig > original_ucl)
                ori_lower_ooc = np.sum(values_orig < original_lcl)
                ori_ooc_count = ori_upper_ooc + ori_lower_ooc
            
            return {
                "Pattern": "Insufficient Data",
                "Skew": np.nan,
                "CB": np.nan,
                "Resolution_Estimated": resolution,
                "Suggest UCL": np.nan,
                "Suggest LCL": np.nan,
                "Static UCL": np.nan,
                "Static LCL": np.nan,
                "Raw UCL After Resolution": np.nan,
                "Raw LCL After Resolution": np.nan,
                "Sug USL": np.nan,
                "Sug LSL": np.nan,
                "TightenNeeded": False,
                "TotalDataCount": len(values_orig),
                "DataCountUsed": len(values_orig),
                "HardRule": "Insufficient Data",
                "DetectionLimit": detection_limit,
                "CL_Center": np.nan,
                "Sigma_Est": 0.0,
                "Sigma_Est_Upper": 0.0,
                "Sigma_Est_Lower": 0.0,
                "Original_UCL_K_Set": np.nan,
                "Original_LCL_K_Set": np.nan,
                "Suggest_UCL_K_Set": np.nan,
                "Suggest_LCL_K_Set": np.nan,
                "Ori_K_Set": np.nan,
                "Sug_K_Set": np.nan,
                "Total_Adj_Units": 0.0,
                "Ori_OOC_Count": ori_ooc_count,
                "Static_OOC_Count": 0,
                "Final_OOC_Count": 0,
                "Original_Tolerance": np.nan,
                "New_Tolerance": np.nan,
                "Diff_Ratio_%": np.nan,
                "Tighten_Threshold_%": np.nan
            }
        
        # 1.5. Hard Rule Check (優先檢查，在 Pattern Diagnosis 之前)
        print(f"    [Debug] 原始數據點數量: {len(values_orig)}")
        print(f"    [Debug] 原始數據範圍: {np.min(values_orig):.4f} ~ {np.max(values_orig):.4f}")
        
        UCL_hr, LCL_hr, rule_satisfied, rule_applied_name = self.apply_discrete_hard_rules(
            values_orig, resolution, len(values_orig))
        
        # 讀取額外參數
        detection_limit = df['DetectionLimit'].iloc[0] if 'DetectionLimit' in df.columns and len(df)>0 else np.nan
        original_ucl = df['UCL'].iloc[0] if 'UCL' in df.columns and len(df)>0 else np.nan
        original_lcl = df['LCL'].iloc[0] if 'LCL' in df.columns and len(df)>0 else np.nan
        target_val = df['Target'].iloc[0] if 'Target' in df.columns and len(df)>0 and pd.notna(df['Target'].iloc[0]) else np.nan

        if rule_satisfied:
            print(f"    [Debug] Hard Rule 滿足: {rule_applied_name}")
            
            # 根據特性類型決定使用 Hard Rule 的哪些管制線
            if characteristic == 'Bigger':
                # Bigger: 只使用 Hard Rule 的 LCL，UCL 保持原始值
                suggest_ucl_hr = original_ucl if not pd.isna(original_ucl) else UCL_hr
                suggest_lcl_hr = LCL_hr
                print(f"    [Debug] Bigger 特性: 只調整 LCL = {LCL_hr:.4f}, UCL 保持原始 = {suggest_ucl_hr:.4f}")
            elif characteristic == 'Smaller':
                # Smaller: 只使用 Hard Rule 的 UCL，LCL 保持原始值
                suggest_ucl_hr = UCL_hr
                suggest_lcl_hr = original_lcl if not pd.isna(original_lcl) else LCL_hr
                print(f"    [Debug] Smaller 特性: 只調整 UCL = {UCL_hr:.4f}, LCL 保持原始 = {suggest_lcl_hr:.4f}")
            else:
                # Nominal: 雙邊都使用 Hard Rule
                suggest_ucl_hr = UCL_hr
                suggest_lcl_hr = LCL_hr
                print(f"    [Debug] Nominal 特性: 雙邊都調整 UCL = {UCL_hr:.4f}, LCL = {LCL_hr:.4f}")
            
            # Hard Rule 成立時，CL_Center 應為最終 UCL 和 LCL 的中點
            cl_center_hr = (suggest_ucl_hr + suggest_lcl_hr) / 2
            
            # ========== 簡化版：只為了 Tighten Check 計算 Tolerance (使用錨點邏輯) ==========
            # 決定一個固定的「錨點 (Anchor)」
            # Hard Rule 1: 使用 Target 或預設值（因為是常數）
            # Hard Rule 2 和 3: 使用 median
            if rule_applied_name == "Hard Rule 1: Constant/Near Constant":
                # Hard Rule 1 使用 Target 或預設值
                if not pd.isna(target_val):
                    anchor = target_val
                else:
                    anchor = 0.0  # 對於 Smaller 預設為 0
                    # 對於 Bigger 若無 Target，使用數據最大值
                    if characteristic == 'Bigger' and len(values_orig) > 0:
                        anchor = np.max(values_orig)
            else:
                # Hard Rule 2 和 3 使用 median
                anchor = np.median(values_orig)
            
            tighten_needed_hr = False
            original_tol = np.nan
            new_tol = np.nan
            diff_ratio = np.nan
            tighten_threshold = np.nan
            
            print(f"    [Debug] Hard Rule Tighten 分析開始... (特性={characteristic}, 錨點={anchor:.4f})")
            
            # Hard Rule 1 (常數/近常數) 的特殊處理：只要管制線比原本緊就 tighten
            if rule_applied_name == "Hard Rule 1: Constant/Near Constant":
                print(f"    [Debug] Hard Rule 1: 常數/近常數模式")
                if not (pd.isna(original_ucl) or pd.isna(original_lcl)):
                    # 浮點數精度修正：round 到合理位數避免微小誤差
                    decimals = 10  # 保留 10 位小數
                    suggest_ucl_hr_rounded = round(suggest_ucl_hr, decimals)
                    suggest_lcl_hr_rounded = round(suggest_lcl_hr, decimals)
                    original_ucl_rounded = round(original_ucl, decimals)
                    original_lcl_rounded = round(original_lcl, decimals)
                    
                    # 檢查是否比原本更緊（UCL 更小或 LCL 更大）
                    is_tighter = False
                    if characteristic == 'Nominal':
                        is_tighter = (suggest_ucl_hr_rounded <= original_ucl_rounded) or (suggest_lcl_hr_rounded >= original_lcl_rounded)
                    elif characteristic == 'Smaller':
                        is_tighter = (suggest_ucl_hr_rounded <= original_ucl_rounded)
                    elif characteristic == 'Bigger':
                        is_tighter = (suggest_lcl_hr_rounded >= original_lcl_rounded)
                    
                    tighten_needed_hr = is_tighter
                    print(f"    [Debug] Hard Rule 1: Control Limit 比較結果 = {'更緊' if is_tighter else '未更緊'}，TightenNeeded = {tighten_needed_hr}")
                    print(f"    [Debug] Hard Rule 1: UCL {suggest_ucl_hr_rounded:.10f} vs {original_ucl_rounded:.10f}, LCL {suggest_lcl_hr_rounded:.10f} vs {original_lcl_rounded:.10f}")
                else:
                    # 無原始管制線，直接 tighten
                    tighten_needed_hr = True
                    print(f"    [Debug] Hard Rule 1: 無原始管制線，TightenNeeded = Yes")
            
            # Hard Rule 2 和 3: 需要進行 tolerance 判斷
            elif rule_applied_name in ["Hard Rule 2: Two Categories", "Hard Rule 3: Three Categories Spaced by Resolution"]:
                print(f"    [Debug] {rule_applied_name}: 使用 Tolerance 判斷")
                
                if not (pd.isna(original_ucl) or pd.isna(original_lcl)):
                    # 根據特性計算 Tolerance (寬度)
                    if characteristic == 'Nominal':
                        # 雙邊：直接算距離
                        original_tol = original_ucl - original_lcl
                        new_tol = suggest_ucl_hr - suggest_lcl_hr
                        print(f"    [Debug] Nominal 特性: 原始 tolerance={original_tol:.6f}, 新 tolerance={new_tol:.6f}")
                        
                    elif characteristic == 'Smaller':
                        # 望小：只看 UCL 到錨點的距離
                        original_tol = original_ucl - anchor
                        new_tol = suggest_ucl_hr - anchor
                        print(f"    [Debug] Smaller 特性 (錨點={anchor:.4f}): 原始 tolerance={original_tol:.6f}, 新 tolerance={new_tol:.6f}")
                        
                    elif characteristic == 'Bigger':
                        # 望大：只看 錨點 到 LCL 的距離
                        original_tol = anchor - original_lcl
                        new_tol = anchor - suggest_lcl_hr
                        print(f"    [Debug] Bigger 特性 (錨點={anchor:.4f}): 原始 tolerance={original_tol:.6f}, 新 tolerance={new_tol:.6f}")
                        
                    else:
                        # 預設 Nominal
                        original_tol = original_ucl - original_lcl
                        new_tol = suggest_ucl_hr - suggest_lcl_hr
                        print(f"    [Debug] 未知特性，預設 Nominal: 原始 tolerance={original_tol:.6f}, 新 tolerance={new_tol:.6f}")
                    
                    # 使用 SOP 6.4 邏輯進行 Tighten 檢查（容差比對）
                    if original_tol > 1e-9 and new_tol > 1e-9:
                        tighten_needed_hr, diff_ratio, tighten_threshold = self.check_tighten_with_details(
                            original_tol, new_tol, len(values_orig)
                        )
                        
                        print(f"    [Debug] Hard Rule Tolerance 比對結果:")
                        print(f"    [Debug]   原始 tolerance: {original_tol:.6f}")
                        print(f"    [Debug]   新 tolerance: {new_tol:.6f}")
                        print(f"    [Debug]   變化率: {diff_ratio:.2f}%")
                        print(f"    [Debug]   Tighten 閾值 (N%): {tighten_threshold:.0f}%")
                        print(f"    [Debug]   TightenNeeded: {'Yes' if tighten_needed_hr else 'No'}")
                        
                        if tighten_needed_hr:
                            print(f"    [Debug] ✓ Hard Rule: Tolerance 收緊 {diff_ratio:.2f}% > {tighten_threshold:.0f}% 閾值，TightenNeeded = Yes")
                        else:
                            print(f"    [Debug] ✗ Hard Rule: Tolerance 變化 {diff_ratio:.2f}% ≤ {tighten_threshold:.0f}% 閾值，TightenNeeded = No")
                    else:
                        print(f"    [Debug] Hard Rule: Tolerance 無效 (original_tol={original_tol:.6f}, new_tol={new_tol:.6f})")
                else:
                    # 沒有原始管制線時，Hard Rule 觸發即為需要 tighten
                    tighten_needed_hr = True
                    print(f"    [Debug] Hard Rule: 無原始管制線，TightenNeeded = Yes")
            else:
                # 其他未知的 Hard Rule（保險起見）
                if not (pd.isna(original_ucl) or pd.isna(original_lcl)):
                    tighten_needed_hr = False
                    print(f"    [Debug] 未知 Hard Rule 類型，預設 TightenNeeded = No")
                else:
                    tighten_needed_hr = True
                    print(f"    [Debug] 未知 Hard Rule 類型但無原始管制線，TightenNeeded = Yes")
            
            # 計算 Ori OOC Count (使用原始管制線)
            ori_ooc_count = 0
            if not (pd.isna(original_ucl) or pd.isna(original_lcl)):
                ori_upper_ooc = np.sum(values_orig > original_ucl)
                ori_lower_ooc = np.sum(values_orig < original_lcl)
                ori_ooc_count = ori_upper_ooc + ori_lower_ooc
            
            # 計算 Static/Final OOC Count (使用 Hard Rule 的管制線)
            static_ooc_count = 0
            static_upper_ooc = np.sum(values_orig > suggest_ucl_hr)
            static_lower_ooc = np.sum(values_orig < suggest_lcl_hr)
            static_ooc_count = static_upper_ooc + static_lower_ooc
            
            # 計算 Ori OOC Count (使用原始管制線)
            ori_ooc_count = 0
            if not (pd.isna(original_ucl) or pd.isna(original_lcl)):
                ori_upper_ooc = np.sum(values_orig > original_ucl)
                ori_lower_ooc = np.sum(values_orig < original_lcl)
                ori_ooc_count = ori_upper_ooc + ori_lower_ooc
            
            # 計算 Static/Final OOC Count (使用 Hard Rule 的管制線)
            static_ooc_count = 0
            static_upper_ooc = np.sum(values_orig > suggest_ucl_hr)
            static_lower_ooc = np.sum(values_orig < suggest_lcl_hr)
            static_ooc_count = static_upper_ooc + static_lower_ooc
            
            # 📐 [Bug Fix] Hard Rule 與 SOP 6.6 的邏輯衝突檢查
            # 確保 Hard Rule 也遵守「只能收緊、不能放寬」原則
            clamped = False  # 追蹤是否發生 clamp
            if not (pd.isna(original_ucl) or pd.isna(original_lcl)):
                # 檢查是否違反 Tighten-only 原則（使用 round 避免浮點數精度問題）
                decimals = 10
                suggest_ucl_hr_rounded = round(suggest_ucl_hr, decimals)
                suggest_lcl_hr_rounded = round(suggest_lcl_hr, decimals)
                original_ucl_rounded = round(original_ucl, decimals)
                original_lcl_rounded = round(original_lcl, decimals)
                
                if characteristic == 'Nominal' or characteristic == 'Smaller':
                    if suggest_ucl_hr_rounded > original_ucl_rounded:
                        print(f"    [Warning] Hard Rule UCL ({suggest_ucl_hr:.10f}) > Original UCL ({original_ucl:.10f})，強制 clamp 至原始值")
                        suggest_ucl_hr = original_ucl
                        clamped = True
                
                if characteristic == 'Nominal' or characteristic == 'Bigger':
                    if suggest_lcl_hr_rounded < original_lcl_rounded:
                        print(f"    [Warning] Hard Rule LCL ({suggest_lcl_hr:.10f}) < Original LCL ({original_lcl:.10f})，強制 clamp 至原始值")
                        suggest_lcl_hr = original_lcl
                        clamped = True
                
                # ⚠️ 關鍵：如果發生了 clamp（界限變寬），強制 tighten_needed = False
                if clamped:
                    tighten_needed_hr = False
                    print(f"    [Warning] 檢測到管制線變寬，強制設定 TightenNeeded = False")
                
                # 重新計算 cl_center 和 tolerance（因可能被 clamp 過）
                cl_center_hr = (suggest_ucl_hr + suggest_lcl_hr) / 2
                
                if characteristic == 'Nominal':
                    new_tol = suggest_ucl_hr - suggest_lcl_hr
                elif characteristic == 'Smaller':
                    new_tol = suggest_ucl_hr - anchor
                elif characteristic == 'Bigger':
                    new_tol = anchor - suggest_lcl_hr
            
            # ⚠️ Hard Rule 返回前：特別說明不套用 Resolution 精度修正
            print(f"\n    [Hard Rule Return] Pattern: {rule_applied_name}")
            print(f"    [Hard Rule Return] ⚠️ Hard Rule 不套用 Resolution 精度修正（避免造成 OOC）")
            print(f"    [Hard Rule Return] Suggest UCL: {suggest_ucl_hr:.10f}")
            print(f"    [Hard Rule Return] Suggest LCL: {suggest_lcl_hr:.10f}\n")
            
            # Hard Rule 滿足時，返回結果
            return {
                "Pattern": rule_applied_name,  # 直接顯示具體的 Hard Rule
                "Skew": np.nan,  # Hard Rule 不需要 Skew/CB
                "CB": np.nan,
                "Resolution_Estimated": resolution,
                "Suggest UCL": suggest_ucl_hr,  # ⚠️ Hard Rule：保持原始精度，不套用 Resolution
                "Suggest LCL": suggest_lcl_hr,  # ⚠️ Hard Rule：保持原始精度，不套用 Resolution
                "Static UCL": suggest_ucl_hr,
                "Static LCL": suggest_lcl_hr, 
                "TightenNeeded": tighten_needed_hr,  # 基於 SOP 6.4 Tolerance 比對判定
                "TotalDataCount": len(values_orig),
                "DataCountUsed": len(values_orig),
                "HardRule": rule_applied_name,
                "DetectionLimit": detection_limit,
                "CL_Center": cl_center_hr,
                "Sigma_Est": 0.0, 
                "Sigma_Est_Upper": 0.0,
                "Sigma_Est_Lower": 0.0,
                "Original_UCL_K_Set": np.nan,  # Hard Rule 無sigma概念
                "Original_LCL_K_Set": np.nan,
                "Suggest_UCL_K_Set": np.nan,
                "Suggest_LCL_K_Set": np.nan,
                "Ori_K_Set": np.nan,
                "Sug_K_Set": np.nan,
                "Total_Adj_Units": 0.0,
                "Ori_OOC_Count": ori_ooc_count,
                "Static_OOC_Count": static_ooc_count,
                "Final_OOC_Count": static_ooc_count,  # Hard Rule 沒有退格，Static = Final
                "Original_Tolerance": original_tol,  # ✅ 新增：原始 tolerance
                "New_Tolerance": new_tol,  # ✅ 新增：新 tolerance
                "Diff_Ratio_%": diff_ratio,  # ✅ 新增：變化率百分比
                "Tighten_Threshold_%": tighten_threshold  # ✅ 新增：Tighten 閾值
            }
        
        print(f"    [Debug] Hard Rule 不滿足，繼續進行 Pattern Diagnosis")
        
        # 2-4. Pattern Diagnosis & Outlier Filter
        print(f"    [Debug] 原始數據點數量: {len(values_orig)}")
        print(f"    [Debug] 原始數據範圍: {np.min(values_orig):.4f} ~ {np.max(values_orig):.4f}")
        
        values_for_pattern = self.data_prep_for_pattern(values_orig)
        print(f"    [Debug] 預處理後數據點數量: {len(values_for_pattern)}")
        
        pattern, skew_value, cb_value = self.pattern_diagnosis(values_for_pattern, resolution)
        print(f"    [Debug] 診斷出的模式: {pattern}, Skew: {skew_value:.4f}, CB: {cb_value:.4f}")
        
        values_filtered = self.outlier_filter(values_orig, pattern)
        N_filtered = len(values_filtered)
        print(f"    [Debug] 過濾後數據點數量: {N_filtered}")
        if N_filtered > 0:
            print(f"    [Debug] 過濾後數據範圍: {np.min(values_filtered):.4f} ~ {np.max(values_filtered):.4f}")
        else:
            print(f"    [Debug] 過濾後數據為空!")
        
        # 檢查過濾後的數據是否為空
        if N_filtered == 0:
            print(f"    [Warning] 過濾後沒有有效數據點，跳過計算")
            
            # 讀取必要參數
            detection_limit = df['DetectionLimit'].iloc[0] if 'DetectionLimit' in df.columns and len(df)>0 else np.nan
            original_ucl = df['UCL'].iloc[0] if 'UCL' in df.columns and len(df)>0 else np.nan
            original_lcl = df['LCL'].iloc[0] if 'LCL' in df.columns and len(df)>0 else np.nan
            
            # 計算 Ori OOC Count (使用原始數據，即使過濾後為空)
            ori_ooc_count = 0
            if not (pd.isna(original_ucl) or pd.isna(original_lcl)):
                ori_upper_ooc = np.sum(values_orig > original_ucl)
                ori_lower_ooc = np.sum(values_orig < original_lcl)
                ori_ooc_count = ori_upper_ooc + ori_lower_ooc
            
            return {
                "Pattern": "No Data After Filter",
                "Skew": np.nan,
                "CB": np.nan,
                "Resolution_Estimated": resolution,
                "Suggest UCL": np.nan,
                "Suggest LCL": np.nan,
                "Static UCL": np.nan,
                "Static LCL": np.nan,
                "Raw UCL After Resolution": np.nan,
                "Raw LCL After Resolution": np.nan,
                "Sug USL": np.nan,
                "Sug LSL": np.nan,
                "TightenNeeded": False,
                "TotalDataCount": len(values_orig),
                "DataCountUsed": N_filtered,
                "HardRule": "No Data After Filter",
                "DetectionLimit": detection_limit,
                "CL_Center": np.nan,
                "Sigma_Est": 0.0,
                "Sigma_Est_Upper": 0.0,
                "Sigma_Est_Lower": 0.0,
                "Original_UCL_K_Set": np.nan,
                "Original_LCL_K_Set": np.nan,
                "Suggest_UCL_K_Set": np.nan,
                "Suggest_LCL_K_Set": np.nan,
                "Ori_K_Set": np.nan,
                "Sug_K_Set": np.nan,
                "Total_Adj_Units": 0.0,
                "Ori_OOC_Count": ori_ooc_count,
                "Static_OOC_Count": 0,  # 過濾後無數據，無法計算 Static
                "Final_OOC_Count": 0,
                "Original_Tolerance": np.nan,
                "New_Tolerance": np.nan,
                "Diff_Ratio_%": np.nan,
                "Tighten_Threshold_%": np.nan
            }
        
        # 讀取額外參數（用於後續計算）
        detection_limit = df['DetectionLimit'].iloc[0] if 'DetectionLimit' in df.columns and len(df)>0 else np.nan
        original_ucl = df['UCL'].iloc[0] if 'UCL' in df.columns and len(df)>0 else np.nan
        original_lcl = df['LCL'].iloc[0] if 'LCL' in df.columns and len(df)>0 else np.nan
        target_val = df['Target'].iloc[0] if 'Target' in df.columns and len(df)>0 and pd.notna(df['Target'].iloc[0]) else np.nan

        # 🔥 計算 Kurtosis（用於判斷是否需要加倍 sigma）
        kurtosis_value = kurtosis(values_filtered, fisher=True, bias=False)  # 轉換為非 Fisher 形式
        print(f"    [Debug] Kurtosis: {kurtosis_value:.4f}")

        # 5. Statistical Model Fitting (SOP 5)
        UCL_static, LCL_static, UR_robust, LR_robust, UCL3_ecdf, LCL3_ecdf, Sug_USL, Sug_LSL = self.calc_CL(
            values_filtered, pattern, resolution, characteristic, kurtosis_value
        )
        
        # 應用 resolution 調整到 Sug USL/LSL
        if resolution is not None and resolution > 0:
            Sug_USL = self.apply_resolution_precision(Sug_USL, resolution, "Sug_USL")
            Sug_LSL = self.apply_resolution_precision(Sug_LSL, resolution, "Sug_LSL")
        
        # 決定 CL Center 和 Sigma Est
        cl_center = np.nan
        sigma_est_u = np.nan
        sigma_est_l = np.nan
        
        if pattern == "Constant":
            cl_center = np.median(values_filtered)
            sigma_est_u = sigma_est_l = 0.0
        elif pattern == "Near Constant":
            # Near Constant 改用 ECDF 計算 sigma（與 Attribute 相同邏輯）
            cl_center = np.median(values_filtered)
            T_upper_ecdf = UCL3_ecdf - cl_center
            T_lower_ecdf = cl_center - LCL3_ecdf
            if T_upper_ecdf > 1e-9 and T_lower_ecdf > 1e-9:
                sigma_est_u = T_upper_ecdf / 3.0
                sigma_est_l = T_lower_ecdf / 3.0
            else:
                sigma_est_u = UR_robust if UR_robust > 0 else 0.0
                sigma_est_l = LR_robust if LR_robust > 0 else 0.0
        elif pattern == "Normal":
            cl_center = np.mean(values_filtered)
            sigma_est_u = sigma_est_l = np.std(values_filtered, ddof=1)
        elif pattern in ["Skew-Right", "Skew-Left"]:
            cl_center = np.median(values_filtered)
            sigma_est_u = UR_robust
            sigma_est_l = LR_robust
        else: # ECDF 相關模式 (Bimodal, Attribute)
            cl_center = np.median(values_filtered)
            T_upper_ecdf = UCL3_ecdf - cl_center
            T_lower_ecdf = cl_center - LCL3_ecdf
            if T_upper_ecdf > 1e-9 and T_lower_ecdf > 1e-9:
                sigma_est_u = T_upper_ecdf / 3.0
                sigma_est_l = T_lower_ecdf / 3.0
            else:
                sigma_est_u = UR_robust if UR_robust > 0 else 0.0
                sigma_est_l = LR_robust if LR_robust > 0 else 0.0

        sigma_est_output = max(sigma_est_u, sigma_est_l)
        if not np.isfinite(sigma_est_output): 
            sigma_est_output = 0.0 

        # 計算 Ori OOC Count (使用原始管制線)
        ori_ooc_count = 0
        if not (pd.isna(original_ucl) or pd.isna(original_lcl)):
            ori_upper_ooc = np.sum(values_orig > original_ucl)
            ori_lower_ooc = np.sum(values_orig < original_lcl)
            ori_ooc_count = ori_upper_ooc + ori_lower_ooc

        # 6. Control Limit Adjustment (SOP 6.1/6.2/6.3/6.5)
        # 使用原始數據 (values_orig) 來計算 OOC count，而不是過濾後的數據
        UCL_suggest, LCL_suggest, static_ooc_count, final_ooc_count, total_adj_units = self.adjust_CL_based_on_OOC(
            values_orig, UCL_static, LCL_static, pattern, resolution, sigma_est_u, sigma_est_l, 2, characteristic
        )
        
        # === 保存 Raw UCL/LCL（已經過 OOC 退格 + resolution 調整，未經過 capping）===
        # 這些值可以超出原始管制線，反映統計上的真實建議（包括放寬）
        Raw_UCL_After_Resolution = UCL_suggest
        Raw_LCL_After_Resolution = LCL_suggest

        # 6.5. Detection Limit Rule (僅適用於 Smaller 特性)
        if characteristic == 'Smaller' and not np.isnan(detection_limit):
            if UCL_suggest < detection_limit:
                UCL_suggest = detection_limit

        # 6.6. 管制線寬度約束 - 不允許建議管制線比原始管制線更寬
        if not (pd.isna(original_ucl) or pd.isna(original_lcl)):
            # Static UCL/LCL 約束：不超出原始管制線
            if not np.isnan(UCL_static) and UCL_static > original_ucl:
                UCL_static = original_ucl
            if not np.isnan(LCL_static) and LCL_static < original_lcl:
                LCL_static = original_lcl
                
            # Suggest UCL/LCL 約束：不超出原始管制線  
            if not np.isnan(UCL_suggest) and UCL_suggest > original_ucl:
                UCL_suggest = original_ucl
            if not np.isnan(LCL_suggest) and LCL_suggest < original_lcl:
                LCL_suggest = original_lcl
        
        # 6.6b. 根據特性類型設定 Static UCL/LCL 為原始值
        if characteristic == 'Smaller':
            # Smaller 特性：Static LCL = 原始 LCL (只需要 USL)
            if not pd.isna(original_lcl):
                LCL_static = original_lcl
        elif characteristic == 'Bigger':
            # Bigger 特性：Static UCL = 原始 UCL (只需要 LSL)
            if not pd.isna(original_ucl):
                UCL_static = original_ucl
        
        # 6.7a. Attribute/Near Constant 特殊邏輯：避免 UCL/LCL 卡在 max/min
        if pattern in ["Attribute", "Near Constant"] and resolution is not None and resolution > 0 and len(values_orig) > 0:
            print(f"    [Debug] 進入 Attribute/Near Constant 特殊邏輯")
            print(f"    [Debug] values_orig 長度: {len(values_orig)}")
            print(f"    [Debug] resolution: {resolution}")
            
            max_val = np.max(values_orig)
            min_val = np.min(values_orig)
            print(f"    [Debug] max_val: {max_val:.6f}, min_val: {min_val:.6f}")
            
            # 暫存調整前的值
            temp_UCL_suggest = UCL_suggest
            temp_LCL_suggest = LCL_suggest
            print(f"    [Debug] 調整前 UCL_suggest: {UCL_suggest:.6f}, LCL_suggest: {LCL_suggest:.6f}")
            
            # 檢查是否需要調整（UCL <= max 或 LCL >= min）
            need_adjust_ucl = not np.isnan(UCL_suggest) and UCL_suggest <= max_val
            need_adjust_lcl = not np.isnan(LCL_suggest) and LCL_suggest >= min_val
            print(f"    [Debug] need_adjust_ucl: {need_adjust_ucl}, need_adjust_lcl: {need_adjust_lcl}")
            
            if need_adjust_ucl:
                temp_UCL_suggest = UCL_suggest + resolution
                print(f"    [Debug] UCL <= max ({UCL_suggest:.6f} <= {max_val:.6f})，調整為 {temp_UCL_suggest:.6f}")
            
            if need_adjust_lcl:
                temp_LCL_suggest = LCL_suggest - resolution
                print(f"    [Debug] LCL >= min ({LCL_suggest:.6f} >= {min_val:.6f})，調整為 {temp_LCL_suggest:.6f}")
            
            # 檢查調整後是否比原始管制線更寬鬆
            if not (pd.isna(original_ucl) or pd.isna(original_lcl)):
                # 計算原始容差
                if characteristic == 'Nominal':
                    ori_tolerance = original_ucl - original_lcl
                    new_tolerance = temp_UCL_suggest - temp_LCL_suggest
                elif characteristic == 'Smaller':
                    ori_tolerance = original_ucl - cl_center
                    new_tolerance = temp_UCL_suggest - cl_center
                elif characteristic == 'Bigger':
                    ori_tolerance = cl_center - original_lcl
                    new_tolerance = cl_center - temp_LCL_suggest
                else:
                    ori_tolerance = original_ucl - original_lcl
                    new_tolerance = temp_UCL_suggest - temp_LCL_suggest
                
                # 如果調整後更寬鬆，不進行調整（保持調整前的值）
                if new_tolerance > ori_tolerance:
                    print(f"    [Debug] 調整後容差 ({new_tolerance:.4f}) 比原始容差 ({ori_tolerance:.4f}) 更寬，不進行調整")
                    # 不更新 UCL_suggest 和 LCL_suggest，保持調整前的值
                    # UCL_suggest 和 LCL_suggest 維持原本在第1558-1564行約束後的值
                else:
                    print(f"    [Debug] 調整後容差 ({new_tolerance:.4f}) 未比原始容差 ({ori_tolerance:.4f}) 更寬，採用調整值")
                    UCL_suggest = temp_UCL_suggest
                    LCL_suggest = temp_LCL_suggest
            else:
                # 沒有原始管制線時，直接採用調整值
                print(f"    [Debug] 無原始管制線，直接採用調整值")
                UCL_suggest = temp_UCL_suggest
                LCL_suggest = temp_LCL_suggest
            
        if characteristic == 'Smaller':
            # Smaller 只允許 UCL 收緊，LCL 直接使用原始值 (只需要 USL)
            if not pd.isna(original_lcl):
                LCL_suggest = original_lcl
        elif characteristic == 'Bigger':
            # Bigger 只允許 LCL 收緊，UCL 直接使用原始值 (只需要 LSL)
            if not pd.isna(original_ucl):
                UCL_suggest = original_ucl

        # 🔧 [修正] 在所有 capping/adjustment 完成後，使用最終的 Suggest UCL/LCL 重新計算 Final_OOC_Count
        print(f"    [Debug Final OOC] === 重新計算 Final_OOC_Count (基於最終 Suggest UCL/LCL) ===")
        print(f"    [Debug Final OOC] 最終 Suggest UCL: {UCL_suggest:.6f}, LCL: {LCL_suggest:.6f}")
        
        final_upper_ooc = np.sum(values_orig > UCL_suggest)
        final_lower_ooc = np.sum(values_orig < LCL_suggest)
        
        # 根據特性計算最終有效的 OOC
        if characteristic == 'Smaller':
            final_ooc_count = final_upper_ooc
            print(f"    [Debug Final OOC] Smaller 特性：只計算上界 OOC = {final_ooc_count}")
        elif characteristic == 'Bigger':
            final_ooc_count = final_lower_ooc
            print(f"    [Debug Final OOC] Bigger 特性：只計算下界 OOC = {final_ooc_count}")
        else:
            final_ooc_count = final_upper_ooc + final_lower_ooc
            print(f"    [Debug Final OOC] Nominal 特性：上界 {final_upper_ooc} + 下界 {final_lower_ooc} = {final_ooc_count}")

        # 7. Tighten 判定 (SOP 6.4) 
        tighten_flag = False
        diff_ratio = np.nan
        tighten_threshold = np.nan
        original_tol = np.nan
        new_tol = np.nan
        
        if not (pd.isna(original_ucl) or pd.isna(original_lcl)):
            center_val = cl_center 
            
            if characteristic == 'Nominal':
                # Nominal: tolerance = UCL – LCL
                original_tol = original_ucl - original_lcl
                new_tol = UCL_suggest - LCL_suggest
            elif characteristic == 'Smaller':
                # Smaller: tolerance = UCL – CL_Center
                original_tol = original_ucl - center_val
                new_tol = UCL_suggest - center_val
                
            elif characteristic == 'Bigger':
                # Bigger: tolerance = CL_Center – LCL
                original_tol = center_val - original_lcl
                new_tol = center_val - LCL_suggest
            else:
                # 預設為 Nominal
                original_tol = original_ucl - original_lcl
                new_tol = UCL_suggest - LCL_suggest

            # 進行 Tighten 檢查，並取得詳細資訊
            if original_tol > 1e-9 and new_tol > 1e-9:
                tighten_flag, diff_ratio, tighten_threshold = self.check_tighten_with_details(
                    original_tol, new_tol, N_filtered
                )
        
        # 8. 計算現行和建議管制線的K倍數
        # 現行管制線的K倍數
        original_ucl_k_set = np.nan
        original_lcl_k_set = np.nan
        suggest_ucl_k_set = np.nan
        suggest_lcl_k_set = np.nan
        
        if not (pd.isna(original_ucl) or pd.isna(original_lcl)) and not pd.isna(cl_center):
            if sigma_est_u > 1e-9:
                original_ucl_k_set = (original_ucl - cl_center) / sigma_est_u
                suggest_ucl_k_set = (UCL_suggest - cl_center) / sigma_est_u
            if sigma_est_l > 1e-9:
                original_lcl_k_set = (cl_center - original_lcl) / sigma_est_l
                suggest_lcl_k_set = (cl_center - LCL_suggest) / sigma_est_l

        # 計算Ori_k_set和Sug_k_set（根據特性類型決定計算方式）
        ori_k_set = np.nan
        sug_k_set = np.nan
        
        if characteristic == 'Bigger':
            # Bigger 特性：只考慮 LCL 的 K 值
            ori_k_set = original_lcl_k_set
            sug_k_set = suggest_lcl_k_set
        elif characteristic == 'Smaller':
            # Smaller 特性：只考慮 UCL 的 K 值
            ori_k_set = original_ucl_k_set
            sug_k_set = suggest_ucl_k_set
        else:
    
            # Nominal 特性（預設）：取 UCL 和 LCL 的 K 值最大值
            if not pd.isna(original_ucl_k_set) and not pd.isna(original_lcl_k_set):
                ori_k_set = max(original_ucl_k_set, original_lcl_k_set)
            elif not pd.isna(original_ucl_k_set):
                ori_k_set = original_ucl_k_set
            elif not pd.isna(original_lcl_k_set):
                ori_k_set = original_lcl_k_set
                
            if not pd.isna(suggest_ucl_k_set) and not pd.isna(suggest_lcl_k_set):
                sug_k_set = max(suggest_ucl_k_set, suggest_lcl_k_set)
            elif not pd.isna(suggest_ucl_k_set):
                sug_k_set = suggest_ucl_k_set
            elif not pd.isna(suggest_lcl_k_set):
                sug_k_set = suggest_lcl_k_set

        # ==========================================================
        # 10. 最終守門員檢查 (Final Gatekeeper)
        # 確保 Suggest 永遠不比 Original 寬，且狀態一致
        # ==========================================================
        
        if not (pd.isna(original_ucl) or pd.isna(original_lcl)):
            # 建立微量容差避免浮點數誤差 (例如 10.0000000001)
            eps = 1e-11
            modified = False

            # --- UCL 檢查 (針對 Nominal 與 Smaller) ---
            if characteristic in ['Nominal', 'Smaller']:
                if UCL_suggest > original_ucl + eps:
                    print(f"    [Final Check] 警告：UCL_suggest ({UCL_suggest}) 寬於 Ori ({original_ucl})，強制縮回。")
                    UCL_suggest = original_ucl
                    modified = True

            # --- LCL 檢查 (針對 Nominal 與 Bigger) ---
            if characteristic in ['Nominal', 'Bigger']:
                if LCL_suggest < original_lcl - eps:
                    print(f"    [Final Check] 警告：LCL_suggest ({LCL_suggest}) 寬於 Ori ({original_lcl})，強制縮回。")
                    LCL_suggest = original_lcl
                    modified = True

            # --- 如果有被縮回，重新計算 OOC 點數 ---
            if modified:
                final_ooc_count = np.sum((values_orig > UCL_suggest + eps) | (values_orig < LCL_suggest - eps))
                
                # 重新計算 Tolerance 比對
                if characteristic == 'Nominal':
                    new_tol = UCL_suggest - LCL_suggest
                    original_tol = original_ucl - original_lcl
                elif characteristic == 'Smaller':
                    new_tol = UCL_suggest - cl_center
                    original_tol = original_ucl - cl_center
                else: # Bigger
                    new_tol = cl_center - LCL_suggest
                    original_tol = cl_center - original_lcl
                
                # 重新判定是否需要 Tighten
                if original_tol > 1e-9:
                    tighten_flag, diff_ratio, tighten_threshold = self.check_tighten_with_details(
                        original_tol, new_tol, N_filtered
                    )
                else:
                    tighten_flag = False

        # 9. 輸出結果
        
        # 🔒 [關鍵] 統一精度鎖定：所有 CL 相關數值都套用 Resolution 修正
        print(f"\n    [Precision Lock] === 開始精度鎖定 ===")
        print(f"    [Precision Lock] Pattern: {pattern}")
        print(f"    [Precision Lock] Resolution: {resolution}")
        
        if resolution is not None and resolution > 0:
            decimals = self.calculate_decimals_from_resolution(resolution)
            print(f"    [Precision Lock] Target Decimals: {decimals}")
            
            # 套用精度鎖定到所有 CL 相關數值
            UCL_suggest = self.apply_resolution_precision(UCL_suggest, resolution, "UCL_suggest")
            LCL_suggest = self.apply_resolution_precision(LCL_suggest, resolution, "LCL_suggest")
            UCL_static = self.apply_resolution_precision(UCL_static, resolution, "UCL_static")
            LCL_static = self.apply_resolution_precision(LCL_static, resolution, "LCL_static")
            cl_center = self.apply_resolution_precision(cl_center, resolution, "cl_center")
        else:
            print(f"    [Precision Lock] 跳過（Resolution 無效）")
        
        print(f"    [Precision Lock] === 精度鎖定完成 ===\n")
        
        return {
            "Pattern": pattern,
            "Skew": skew_value,
            "CB": cb_value,
            "Resolution_Estimated": resolution,
            "Suggest UCL": UCL_suggest,
            "Suggest LCL": LCL_suggest,
            "Static UCL": UCL_static,
            "Static LCL": LCL_static,
            "Raw UCL After Resolution": Raw_UCL_After_Resolution,
            "Raw LCL After Resolution": Raw_LCL_After_Resolution,
            "Sug USL": Sug_USL,
            "Sug LSL": Sug_LSL,
            "TightenNeeded": tighten_flag,
            "TotalDataCount": len(values_orig),
            "DataCountUsed": N_filtered,
            "HardRule": "None",
            "DetectionLimit": detection_limit,
            "CL_Center": cl_center,
            "Sigma_Est": sigma_est_output,
            "Sigma_Est_Upper": sigma_est_u,
            "Sigma_Est_Lower": sigma_est_l,
            "Original_UCL_K_Set": original_ucl_k_set,
            "Original_LCL_K_Set": original_lcl_k_set,
            "Suggest_UCL_K_Set": suggest_ucl_k_set,
            "Suggest_LCL_K_Set": suggest_lcl_k_set,
            "Ori_K_Set": ori_k_set,
            "Sug_K_Set": sug_k_set,
            "Total_Adj_Units": total_adj_units,
            "Ori_OOC_Count": ori_ooc_count,
            "Static_OOC_Count": static_ooc_count,
            "Final_OOC_Count": final_ooc_count,
            "Original_Tolerance": original_tol,
            "New_Tolerance": new_tol,
            "Diff_Ratio_%": diff_ratio,
            "Tighten_Threshold_%": tighten_threshold
        }

    # === 檔案 I/O 與流程控制 ===

    def load_chart_information(self, filepath):
        """讀取 Chart 設定檔 (Excel)"""
        try:
            df_charts = pd.read_excel(filepath, sheet_name='Chart')
            
            # 確保 Resolution, DetectionLimit, tsmc_ucl, tsmc_lcl 存在
            if 'Resolution' not in df_charts.columns: 
                df_charts['Resolution'] = np.nan
            if 'DetectionLimit' not in df_charts.columns: 
                df_charts['DetectionLimit'] = np.nan
            if 'tsmc_ucl' not in df_charts.columns:
                df_charts['tsmc_ucl'] = np.nan
            if 'tsmc_lcl' not in df_charts.columns:
                df_charts['tsmc_lcl'] = np.nan
                
            # Target, UCL, LCL 為必須欄位
            required_columns = ['Target', 'UCL', 'LCL']
            for col in required_columns:
                if col not in df_charts.columns:
                    raise ValueError(f"缺少必須欄位: '{col}'")
                    
            # 檢查必須欄位是否有數值
            for col in required_columns:
                if df_charts[col].isna().any():
                    missing_rows = df_charts[df_charts[col].isna()].index.tolist()
                    raise ValueError(f"必須欄位 '{col}' 在第 {missing_rows} 行有缺失值")
                
            return df_charts
        except Exception as e:
            print(f"Error loading chart information from {filepath}: {e}")
            return pd.DataFrame()

    def find_matching_file(self, raw_data_directory, group_name, chart_name):
        """根據命名規則匹配對應的 Raw Data CSV (用 _ 分割做精確匹配)"""
        group_name = str(group_name).strip()
        chart_name = str(chart_name).strip()
        
        # 構建精確匹配的前綴模式
        pattern_prefix = f"{group_name}_{chart_name}"
        
        matching_files = []
        try:
            for filename in os.listdir(raw_data_directory):
                if not filename.endswith('.csv'):
                    continue
                
                # 移除 .csv 副檔名
                filename_without_ext = filename[:-4]
                
                # 檢查是否以 pattern_prefix 開頭
                if filename_without_ext.startswith(pattern_prefix):
                    # 檢查後面是否只有 _ 或沒有其他字符（允許尾碼）
                    remainder = filename_without_ext[len(pattern_prefix):]
                    if remainder == '' or remainder.startswith('_'):
                        matching_files.append(filename)
                        print(f"    [Debug] 找到匹配檔案: {filename}")
        
        except Exception as e:
            print(f"    [Warning] 掃描目錄時發生錯誤: {e}")
            return None
        
        if matching_files:
            # 若找到多個符合的檔案，優先選擇最短的（最接近的匹配）
            matching_files.sort(key=len)
            selected_file = matching_files[0]
            
            if len(matching_files) > 1:
                print(f"    [Debug] 找到多個候選檔案: {matching_files}，選擇最短的: {selected_file}")
            
            return os.path.join(raw_data_directory, selected_file)
        
        print(f"    [Warning] 未找到匹配檔案 (GroupName={group_name}, ChartName={chart_name})")
        print(f"    [Debug] 期望的檔案前綴: {pattern_prefix}")
        return None

    def plot_control_chart(self, chart_data, chart_info, suggest_ucl, suggest_lcl,
                        static_ucl, static_lcl, cl_center, pattern, 
                        total_data_count=None, used_data_count=None,
                        output_dir='output_charts', max_x_labels=10,
                        tsmc_ucl=None, tsmc_lcl=None):
        """
        繪製管制圖 (SPC Chart)
        - 固定輸出 800x450 像素，高 DPI 保持清晰
        - 包含所有管制線、建議線、靜態線與標註
        """
        import os
        import numpy as np
        import pandas as pd
        import matplotlib.pyplot as plt
        import matplotlib.dates as mdates

        # 🔥 修復問題2: 檢查空數據，避免閃退
        if chart_data is None or len(chart_data) == 0:
            print(f"    [Warning] chart_data 為空，無法繪製圖表")
            return None
        
        if 'value' not in chart_data.columns:
            print(f"    [Warning] chart_data 缺少 'value' 欄位，無法繪製圖表")
            return None
        
        # 確保有有效數據點
        valid_data = chart_data['value'].dropna()
        if len(valid_data) == 0:
            print(f"    [Warning] chart_data 沒有有效的數據點，無法繪製圖表")
            return None

        os.makedirs(output_dir, exist_ok=True)

        # 🎯 固定最終輸出大小：800x450 像素（不變大）
        target_width_px, target_height_px = 900, 430
        dpi = 100  # 高 DPI 保持線條平滑
        fig = plt.figure(figsize=(target_width_px / dpi, target_height_px / dpi), dpi=dpi)
        ax = fig.add_subplot(111)

        # 全域抗鋸齒設定
        plt.rcParams['lines.antialiased'] = True
        plt.rcParams['patch.antialiased'] = True
        fig.patch.set_antialiased(True)

        # ========== 繪圖主體 ==========
        # X 軸資料：使用等距位置，但顯示實際時間標籤
        x = range(len(chart_data))  # 等距位置
        
        if 'date' in chart_data.columns:
            # 轉換日期並使用統一格式 yyyy/m/d hh:mm
            dates = pd.to_datetime(chart_data['date'])
            date_format = '%Y/%m/%d %H:%M'

            # 設定 X 軸標籤為實際時間，但位置是等距的
            date_labels = [d.strftime(date_format) for d in dates]
            
            # 自動計算最佳標籤數量（垂直顯示，主要考慮視覺密度）
            n_points = len(chart_data)
            
            # 垂直顯示允許更多標籤，盡量多顯示時間資訊
            if n_points <= 15:
                optimal_labels = n_points  # 少於 15 個點，全部顯示
            elif n_points <= 50:
                optimal_labels = min(n_points, int(n_points / 1))  # 顯示約 2/3 的點
            else:
                optimal_labels = min(n_points, 30)  # 最多顯示 30 個標籤
            
            # 確保至少顯示 2 個標籤（首尾）
            optimal_labels = max(2, optimal_labels)
            
            # 總是確保刻度位置在視覺上均勻分布
            if n_points <= optimal_labels:
                # 點數少於最佳標籤數，全部顯示
                tick_positions = list(x)  # x 本身就是 range(n_points)，已經等距
                tick_labels = date_labels
            else:
                # 在等距的 X 軸位置上均勻選擇刻度
                # x 是 range(n_points)，即 [0, 1, 2, ..., n_points-1]
                # 我們要在這個等距序列上均勻選擇 optimal_labels 個位置
                
                # 計算均勻間隔
                x_min = 0
                x_max = n_points - 1
                uniform_positions = np.linspace(x_min, x_max, optimal_labels)
                
                # 將浮點位置四捨五入到最近的整數位置
                tick_indices = [int(round(pos)) for pos in uniform_positions]
                
                # 去除重複並確保在有效範圍內
                tick_indices = sorted(list(set(tick_indices)))
                tick_indices = [idx for idx in tick_indices if 0 <= idx < n_points]
                
                # 生成對應的位置和標籤
                tick_positions = tick_indices  # 直接使用整數位置，這些已經是等距的
                tick_labels = [date_labels[i] for i in tick_indices]
            
            ax.set_xticks(tick_positions)
            ax.set_xticklabels(tick_labels, rotation=90, ha='center', fontsize=9)  # rotation=90 垂直顯示

        y = chart_data['value'].values
        
        # 檢查是否有 ByTool 欄位並繪製分色圖
        if 'ByTool' in chart_data.columns:
            # 過濾掉 NaN 的 Tool 資料
            chart_data_with_tool = chart_data[pd.notna(chart_data['ByTool'])].copy()
            
            if not chart_data_with_tool.empty:
                # 準備顏色映射
                tools = sorted([str(t) for t in chart_data_with_tool['ByTool'].unique() if pd.notna(t)])
                color_cycle = ['#2563eb', '#dc2626', '#16a34a', '#f59e0b', '#7c3aed', '#0891b2']
                tool_colors = {t: color_cycle[i % len(color_cycle)] for i, t in enumerate(tools)}
                
                # 先繪製連線（灰色底線）
                ax.plot(x, y, '-', color="#808182", linewidth=1, alpha=0.9, antialiased=True, zorder=1)
                
                # 按 Tool 分色繪製點
                for tool in tools:
                    mask = chart_data['ByTool'].astype(str) == tool
                    if mask.any():
                        x_tool = [x[i] for i in range(len(x)) if mask.iloc[i]]
                        y_tool = [y[i] for i in range(len(y)) if mask.iloc[i]]
                        ax.scatter(x_tool, y_tool, color=tool_colors[tool], s=45, alpha=0.8, 
                                   label=tool, edgecolors='white', linewidth=0.5, zorder=3)
                
                # 添加 legend
                ax.legend(loc='upper left', fontsize=7, frameon=True, ncol=3, 
                          framealpha=0.9, edgecolor='#d1d5db')
            else:
                # 沒有有效的 Tool 資料，使用原始繪製方式
                ax.plot(x, y, 'bo-', markersize=3, linewidth=1, alpha=0.8, antialiased=True)
        else:
            # 沒有 ByTool 欄位，使用原始繪製方式
            ax.plot(x, y, 'bo-', markersize=3, linewidth=1, alpha=0.8, antialiased=True)

        # 判斷是否只顯示 Sug (當 Ori 與 Sug 差距太大時)
        show_only_sug = False
        if not np.isnan(suggest_ucl) and not np.isnan(suggest_lcl):
            ori_range = abs(chart_info['UCL'] - chart_info['LCL'])
            sug_range = abs(suggest_ucl - suggest_lcl)
            # 如果 Ori 範圍是 Sug 範圍的 3 倍以上，只顯示 Sug
            if ori_range > sug_range * 3:
                show_only_sug = True

        # 各種管制線
        # 判斷 Target 是否距離 Suggest 管制線太遠（避免影響圖表 scale）
        show_target = True
        if "Hard Rule" not in pattern and not np.isnan(suggest_ucl) and not np.isnan(suggest_lcl):
            target_val = chart_info['Target']
            sug_range = suggest_ucl - suggest_lcl
            
            # 如果 Target 超出 Suggest 管制線範圍的 1.5 倍距離，不顯示
            if target_val > suggest_ucl + sug_range * 1.5 or target_val < suggest_lcl - sug_range * 1.5:
                show_target = False
                print(f"    [Debug] Target ({target_val:.3f}) 距離 Suggest 管制線太遠，不顯示於圖表")
        
        # Hard Rule 時不顯示 Target 線，避免視覺混亂
        if show_target and "Hard Rule" not in pattern:
            ax.axhline(y=chart_info['Target'], color='gray', linestyle='-', linewidth=1)
        
        # 只在差距不大時顯示 Ori UCL/LCL
        if not show_only_sug:
            ax.axhline(y=chart_info['UCL'], color='red', linestyle='--', linewidth=2)
            ax.axhline(y=chart_info['LCL'], color='red', linestyle='--', linewidth=2)

        if not np.isnan(suggest_ucl):
            ax.axhline(y=suggest_ucl, color='#555555', linestyle='-', linewidth=1.5)
        if not np.isnan(suggest_lcl):
            ax.axhline(y=suggest_lcl, color='#555555', linestyle='-', linewidth=1.5)

        # TSMC 管制線（綠色虛線）
        if tsmc_ucl is not None and not np.isnan(tsmc_ucl):
            ax.axhline(y=tsmc_ucl, color='green', linestyle='--', linewidth=2, label='TSMC UCL')
        if tsmc_lcl is not None and not np.isnan(tsmc_lcl):
            ax.axhline(y=tsmc_lcl, color='green', linestyle='--', linewidth=2, label='TSMC LCL')

        # Static 線已移除，不再繪製

        # ======= 標題 =======
        # 使用計算結果的 pattern，而非 Excel 的 ExpectedPattern
        sug_ucl_text = f"Sug_UCL: {suggest_ucl}" if not np.isnan(suggest_ucl) else "Sug_UCL: N/A"
        sug_lcl_text = f"Sug_LCL: {suggest_lcl}" if not np.isnan(suggest_lcl) else "Sug_LCL: N/A"
        
        # 將點數統計拆開顯示：Total Cnt (兩年內總點數) 和 Cal Cnt (實際計算點數)
        total_cnt = total_data_count if total_data_count is not None else len(chart_data)
        used_cnt = used_data_count if used_data_count is not None else len(chart_data)
        
        title = (f"{chart_info['GroupName']}@{chart_info['ChartName']}@{chart_info['Characteristics']}\n"
                f"Pattern: {pattern} | Total Cnt: {total_cnt} | Cal Cnt: {used_cnt} | {sug_ucl_text} | {sug_lcl_text}")
        ax.set_title(title, fontsize=11)
        ax.grid(False)

        # ======= 調整 Y 軸範圍（特殊處理 Constant/Near Constant/Hard Rule）=======
        if pattern in ["Constant", "Near Constant"] or "Hard Rule" in pattern:
            # Constant/Near Constant/Hard Rule 模式：包含所有點，但合理設定 Y 軸範圍
            y_data_min = y.min()
            y_data_max = y.max()
            
            # 特殊處理：如果範圍太小（所有點都很接近），適當擴展 Y 軸
            data_range = y_data_max - y_data_min
            if data_range < 1e-6:  # 幾乎是常數
                center = (y_data_min + y_data_max) / 2
                margin = max(abs(center) * 0.1, 1.0)  # 至少 1 個單位的範圍
                y_data_min = center - margin
                y_data_max = center + margin
            print(f"[{pattern}] 包含所有點，Y 軸範圍: [{y_data_min:.3f}, {y_data_max:.3f}]")
        else:
            # 一般模式：使用四分位數法排除極端值
            q1 = np.percentile(y, 25)
            q3 = np.percentile(y, 75)
            iqr = q3 - q1
            lower_bound = q1 - 1.5 * iqr
            upper_bound = q3 + 1.5 * iqr
            
            # 過濾掉異常值後的數據範圍
            y_filtered = y[(y >= lower_bound) & (y <= upper_bound)]
            if len(y_filtered) > 0:
                y_data_min = y_filtered.min()
                y_data_max = y_filtered.max()
            else:
                # 如果全部都是異常值，還是使用原始範圍
                y_data_min = y.min()
                y_data_max = y.max()
        
        # 初始化 all_lines，根據 show_target 決定是否包含 Target
        all_lines = [y_data_min, y_data_max]
        if show_target and "Hard Rule" not in pattern:
            all_lines.append(chart_info['Target'])
        
        # 根據是否只顯示 Sug 來決定包含哪些線
        if show_only_sug:
            # 只顯示 Sug 時，不包含 Ori UCL/LCL
            for v in [suggest_ucl, suggest_lcl]:
                if not np.isnan(v):
                    all_lines.append(v)
        else:
            # 正常情況，包含所有線
            all_lines.extend([chart_info['UCL'], chart_info['LCL']])
            for v in [suggest_ucl, suggest_lcl]:
                if not np.isnan(v):
                    all_lines.append(v)
        
        # 計算 Y 軸範圍，處理 Hard Rule 情況下所有線重疊的問題
        if max(all_lines) == min(all_lines):
            # 所有線重疊時，使用數據範圍設定 Y 軸
            center_line = max(all_lines)
            data_range = y_data_max - y_data_min
            if data_range == 0:
                # 數據也是常數時，使用固定範圍
                y_margin = abs(center_line) * 0.1 if center_line != 0 else 1.0
            else:
                y_margin = data_range * 0.2
            ax.set_ylim(center_line - y_margin, center_line + y_margin)
        else:
            y_margin = (max(all_lines) - min(all_lines)) * 0.1
            ax.set_ylim(min(all_lines) - y_margin, max(all_lines) + y_margin)

        plt.tight_layout(rect=[0, 0, 0.85, 1])

        # ======= 外側標註 =======
        annotations = []
        
        # 根據 show_target 決定是否顯示 Target 標註
        if show_target and "Hard Rule" not in pattern:
            annotations.append((chart_info['Target'], f"Target = {chart_info['Target']}", 'gray', 'normal'))

        # 檢查重疊線並合併標註
        tolerance = 1e-6  # 判斷線重疊的容差值
        
        # 收集所有UCL線的資訊
        ucl_lines = []
        if not show_only_sug and not np.isnan(chart_info['UCL']):
            ucl_lines.append(('Ori', chart_info['UCL'], 'red', 'normal'))
        if not np.isnan(suggest_ucl):
            ucl_lines.append(('Sug', suggest_ucl, 'black', 'bold'))
        
        # 收集所有LCL線的資訊
        lcl_lines = []
        if not show_only_sug and not np.isnan(chart_info['LCL']):
            lcl_lines.append(('Ori', chart_info['LCL'], 'red', 'normal'))
        if not np.isnan(suggest_lcl):
            lcl_lines.append(('Sug', suggest_lcl, 'black', 'bold'))
        
        # 處理UCL重疊
        ucl_groups = []
        for line in ucl_lines:
            placed = False
            for group in ucl_groups:
                if abs(line[1] - group[0][1]) < tolerance:
                    group.append(line)
                    placed = True
                    break
            if not placed:
                ucl_groups.append([line])
        
        # 處理LCL重疊
        lcl_groups = []
        for line in lcl_lines:
            placed = False
            for group in lcl_groups:
                if abs(line[1] - group[0][1]) < tolerance:
                    group.append(line)
                    placed = True
                    break
            if not placed:
                lcl_groups.append([line])
        
        # 檢查 UCL 和 LCL 之間的重疊（處理常數/卡定值情況）
        all_cl_groups = []
        
        # 將所有 UCL 和 LCL 線合併到一起檢查
        all_lines = []
        for name, value, color, weight in ucl_lines:
            all_lines.append((name, value, color, weight, 'UCL'))
        for name, value, color, weight in lcl_lines:
            all_lines.append((name, value, color, weight, 'LCL'))
        
        # 重新分組，包括跨 UCL/LCL 的重疊
        for line in all_lines:
            placed = False
            for group in all_cl_groups:
                if abs(line[1] - group[0][1]) < tolerance:
                    group.append(line)
                    placed = True
                    break
            if not placed:
                all_cl_groups.append([line])
        
        # 生成合併後的標註
        for group in all_cl_groups:
            if len(group) == 1:
                name, value, color, weight, cl_type = group[0]
                annotations.append((value, f"{name} {cl_type} = {value}", color, weight))
            else:
                value = group[0][1]
                # 使用最重要的顏色（優先順序：orange > red > purple）
                color = 'orange' if any(line[2] == 'orange' for line in group) else \
                       'red' if any(line[2] == 'red' for line in group) else 'purple'
                weight = 'bold' if any(line[3] == 'bold' for line in group) else 'normal'
                
                # 分別收集 UCL 和 LCL 的名稱
                ucl_names = [line[0] for line in group if line[4] == 'UCL']
                lcl_names = [line[0] for line in group if line[4] == 'LCL']
                
                # 生成合併標籤
                if ucl_names and lcl_names:
                    # UCL 和 LCL 都有重疊，顯示為 Sug_UCL=Sug_LCL 格式
                    ucl_part = '='.join(ucl_names) + '_UCL' if len(ucl_names) > 1 else ucl_names[0] + '_UCL'
                    lcl_part = '='.join(lcl_names) + '_LCL' if len(lcl_names) > 1 else lcl_names[0] + '_LCL'
                    combined_label = f"{ucl_part}={lcl_part}"
                elif ucl_names:
                    # 只有 UCL 重疊
                    combined_label = '='.join(ucl_names) + ' UCL'
                else:
                    # 只有 LCL 重疊
                    combined_label = '='.join(lcl_names) + ' LCL'
                
                annotations.append((value, f"{combined_label} = {value}", color, weight))

        for y_val, text, color, weight in annotations:
            ax.text(1.02, y_val, text, transform=ax.get_yaxis_transform(),
                    color=color, fontsize=9, va='center', fontweight=weight, clip_on=False)

        # ======= 輸出圖檔 - 只保存 PNG 格式 =======
        filename_png = os.path.join(output_dir, f"{chart_info['GroupName']}_{chart_info['ChartName']}.png")
        
        # 📉 [Bug Fix] 修正記憶體洩漏：明確指定 fig 並放在 finally 區塊
        try:
            # 保存 PNG（打包環境下最穩定，且為唯一實際使用的格式）
            plt.savefig(filename_png, dpi=200, bbox_inches='tight',
                       facecolor='white', edgecolor='none')
            print(f"[Info] PNG 圖片已保存: {filename_png}")
                
        except Exception as e:
            print(f"[Error] 繪圖保存失敗: {e}")
            raise e
        finally:
            # 強制清理 matplotlib 資源
            plt.clf()  # 清除當前圖形
            plt.close(fig)  # 明確關閉 fig 物件
            plt.close('all')  # 關閉所有圖形
            
            # 強制垃圾回收
            import gc
            gc.collect()

        return filename_png  # ✅ 返回 PNG 路徑（唯一實際使用的格式）

    def process_single_chart_data(self, chart_info_row, raw_data_df):
        """將 Chart 設定與原始數據合併並運行核心 SOP 計算"""
        
        group_name = chart_info_row.get('GroupName', 'N/A')
        chart_name = chart_info_row.get('ChartName', 'N/A')
        
        print(f"    [Debug] 開始處理 {group_name}_{chart_name}")
        print(f"    [Debug] 原始 CSV 數據 shape: {raw_data_df.shape}")
        print(f"    [Debug] 原始 CSV 欄位: {list(raw_data_df.columns)}")
        
        # 0. 強制轉換 point_val 為數字型別（雙重保險）
        if 'point_val' in raw_data_df.columns:
            raw_data_df['point_val'] = pd.to_numeric(raw_data_df['point_val'], errors='coerce')
            # 檢查轉換後是否有 NaN（代表轉換失敗）
            nan_count = raw_data_df['point_val'].isna().sum()
            if nan_count > 0:
                print(f"    [Warning] point_val 轉換後有 {nan_count} 筆非數字資料被轉為 NaN")
        
        if len(raw_data_df) > 0 and 'point_val' in raw_data_df.columns:
            try:
                valid_vals = raw_data_df['point_val'].dropna()
                if len(valid_vals) > 0:
                    print(f"    [Debug] 數據範圍: {valid_vals.min():.4f} ~ {valid_vals.max():.4f} (有效筆數: {len(valid_vals)}/{len(raw_data_df)})")
                else:
                    print(f"    [Debug] 警告：所有 point_val 都是無效數字！")
            except Exception as e:
                print(f"    [Debug] 無法計算數據範圍: {e}")
        else:
            print(f"    [Debug] CSV 為空或缺少 point_val 欄位")
        
        # 1. 欄位重命名
        raw_data_df.rename(columns={'point_time': 'date', 'point_val': 'value'}, inplace=True)
        
        # 2. 將所有相關參數從 Chart Info 傳遞給 Raw Data DataFrame
        usl = chart_info_row.get('USL')
        lsl = chart_info_row.get('LSL')
        
        raw_data_df['USL'] = usl
        raw_data_df['LSL'] = lsl
        raw_data_df['DetectionLimit'] = chart_info_row.get('DetectionLimit')
        raw_data_df['Target'] = chart_info_row.get('Target')
        raw_data_df['UCL'] = chart_info_row.get('UCL')
        raw_data_df['LCL'] = chart_info_row.get('LCL')
        raw_data_df['tsmc_ucl'] = chart_info_row.get('tsmc_ucl', np.nan)
        raw_data_df['tsmc_lcl'] = chart_info_row.get('tsmc_lcl', np.nan)
        
        # 3. 動態計算 oos_flag - 根據特性類型決定檢查哪個規格限
        characteristic = chart_info_row.get('Characteristics', 'Nominal')
        print(f"    [Debug] USL: {usl}, LSL: {lsl}, 特性: {characteristic}")
        
        # 根據不同特性類型設定 OOS 檢查邏輯
        if characteristic == 'Bigger':
            # Bigger chart 只需要 LSL，不檢查 USL
            if lsl is not None and not np.isnan(lsl):
                raw_data_df['oos_flag'] = (raw_data_df['value'] < lsl)
                oos_count = raw_data_df['oos_flag'].sum()
                print(f"    [Debug] Bigger chart - 只檢查 LSL，OOS 點數: {oos_count}/{len(raw_data_df)}")
            else:
                raw_data_df['oos_flag'] = False
                print(f"    [Debug] Bigger chart - 無有效 LSL，設置所有點為非 OOS")
        elif characteristic == 'Smaller':
            # Smaller chart 只需要 USL，不檢查 LSL  
            if usl is not None and not np.isnan(usl):
                raw_data_df['oos_flag'] = (raw_data_df['value'] > usl)
                oos_count = raw_data_df['oos_flag'].sum()
                print(f"    [Debug] Smaller chart - 只檢查 USL，OOS 點數: {oos_count}/{len(raw_data_df)}")
            else:
                raw_data_df['oos_flag'] = False
                print(f"    [Debug] Smaller chart - 無有效 USL，設置所有點為非 OOS")
        else:
            # Nominal 或其他類型：檢查雙邊規格限
            if (usl is not None and not np.isnan(usl)) and \
               (lsl is not None and not np.isnan(lsl)) and \
               (usl > lsl):
                raw_data_df['oos_flag'] = (raw_data_df['value'] > usl) | (raw_data_df['value'] < lsl)
                oos_count = raw_data_df['oos_flag'].sum()
                print(f"    [Debug] Nominal chart - 檢查雙邊規格限，OOS 點數: {oos_count}/{len(raw_data_df)}")
            else:
                raw_data_df['oos_flag'] = False
                print(f"    [Debug] Nominal chart - 無有效 USL/LSL，設置所有點為非 OOS")
            
        print(f"    [Debug] 準備進入核心計算，特性: {chart_info_row.get('Characteristics', 'Nominal')}")
        
        try:
            # 4. 運行核心計算
            results = self.process_chart(
                df=raw_data_df,
                value_col='value',
                date_col='date',
                oos_col='oos_flag',
                characteristic=chart_info_row.get('Characteristics', 'Nominal')
            )
            
            # 5. 格式化輸出
            final_output = chart_info_row.to_dict()
            final_output.update(results)
            final_output['Status'] = 'Success'
            
            # 讓舊 UCL/LCL 欄位仍為 Suggest 的值
            final_output['Original UCL'] = final_output['Suggest UCL']
            final_output['Original LCL'] = final_output['Suggest LCL']
            
            # � [最終精度檢查] 二次確認：確保所有輸出值都符合 Resolution
            res_val = results.get('Resolution_Estimated')
            pattern_str = str(final_output.get('Pattern', ''))
            is_hard_rule = pattern_str.startswith('Hard Rule')
            
            print(f"\n    [最終精度檢查] === 開始檢查 ===")
            print(f"    [最終精度檢查] Pattern: {pattern_str}")
            print(f"    [最終精度檢查] Resolution: {res_val}")
            print(f"    [最終精度檢查] Is Hard Rule: {is_hard_rule}")
            
            if is_hard_rule:
                print(f"    [最終精度檢查] ✅ Hard Rule 跳過精度修正")
            elif res_val is None or res_val <= 0:
                print(f"    [最終精度檢查] ⚠️ Resolution 無效，跳過精度修正")
            else:
                target_decimals = self.calculate_decimals_from_resolution(res_val)
                print(f"    [最終精度檢查] 目標小數位數: {target_decimals}")
                
                # ✅ 只對計算產生的 CL 值進行二次確認（不包含 Target/USL/LSL）
                cl_cols_to_check = ['Suggest UCL', 'Suggest LCL', 'Static UCL', 'Static LCL', 
                                   'Original UCL', 'Original LCL', 'CL_Center']
                
                precision_issues_found = False
                
                for key in cl_cols_to_check:
                    if key in final_output and pd.notna(final_output[key]):
                        old_val = final_output[key]
                        
                        # 檢查是否已經對齊（容忍浮點誤差）
                        expected_val = round(float(old_val), target_decimals)
                        
                        if abs(float(old_val) - float(expected_val)) > 1e-10:
                            precision_issues_found = True
                            print(f"    [最終精度檢查] ⚠️ {key} 精度不對: {old_val:.12f} → {expected_val}")
                            final_output[key] = expected_val
                            
                            # 如果是 0 位小數，轉為整數
                            if target_decimals == 0:
                                final_output[key] = int(expected_val)
                        else:
                            # 精度已對齊
                            print(f"    [最終精度檢查] ✅ {key}: {old_val} (已對齊)")
                
                if not precision_issues_found:
                    print(f"    [最終精度檢查] ✅ 所有 CL 值精度已正確對齊")
                
                # Sigma 相關欄位（維持 6 位小數）
                # Sigma 相關欄位（維持 6 位小數）
                for sigma_key in ['Sigma_Est', 'Sigma_Est_Upper', 'Sigma_Est_Lower']:
                    if sigma_key in final_output and not np.isnan(final_output.get(sigma_key, np.nan)):
                        final_output[sigma_key] = round(final_output[sigma_key], 6)
                
                # K 倍數欄位（維持 3 位小數）
                for k_set_key in ['Original_UCL_K_Set', 'Original_LCL_K_Set',
                                 'Suggest_UCL_K_Set', 'Suggest_LCL_K_Set',
                                 'Ori_K_Set', 'Sug_K_Set']:
                    if k_set_key in final_output and not np.isnan(final_output.get(k_set_key, np.nan)):
                        final_output[k_set_key] = round(final_output[k_set_key], 3)
            
            print(f"    [最終精度檢查] === 檢查完成 ===\n")
            
            # 6. 生成管制圖
            try:
                # 🔥 修復問題1: 只傳遞經過時間篩選的數據給繪圖函數
                # 根據設定的時間範圍篩選數據用於繪圖
                if 'date' in raw_data_df.columns:
                    # 轉換日期列
                    raw_data_df['date'] = pd.to_datetime(raw_data_df['date'], errors='coerce')
                    
                    # 應用與 data_integrity 相同的時間篩選邏輯
                    if self.start_date is not None and self.end_date is not None:
                        # 使用自訂日期範圍
                        cutoff_start = pd.Timestamp(self.start_date)
                        cutoff_end = pd.Timestamp(self.end_date)
                        filtered_chart_data = raw_data_df[
                            (raw_data_df['date'] >= cutoff_start) & 
                            (raw_data_df['date'] <= cutoff_end)
                        ].copy()
                        print(f"    [繪圖] 使用自訂日期範圍: {cutoff_start.date()} 至 {cutoff_end.date()}，資料筆數: {len(filtered_chart_data)}")
                    else:
                        # 使用預設的最近2年
                        cutoff = pd.Timestamp.today() - pd.DateOffset(years=2)
                        filtered_chart_data = raw_data_df[raw_data_df['date'] >= cutoff].copy()
                        print(f"    [繪圖] 使用預設日期範圍: 最近2年，資料筆數: {len(filtered_chart_data)}")
                    
                    # 排除 OOS 點（與計算邏輯一致）
                    if 'oos_flag' in filtered_chart_data.columns:
                        filtered_chart_data = filtered_chart_data[~filtered_chart_data['oos_flag'].astype(bool)].copy()
                        print(f"    [繪圖] 排除 OOS 點後資料筆數: {len(filtered_chart_data)}")
                else:
                    # 沒有日期欄位，使用全部數據
                    filtered_chart_data = raw_data_df.copy()
                    print(f"    [繪圖] 無日期欄位，使用全部數據: {len(filtered_chart_data)}")
                
                # 檢查篩選後是否還有數據
                if len(filtered_chart_data) == 0:
                    print(f"     [Warning] 經過時間篩選後無數據，跳過繪圖")
                    final_output['PlotFile'] = 'No Data After Filtering'
                else:
                    plot_filename = self.plot_control_chart(
                        chart_data=filtered_chart_data,
                        chart_info=chart_info_row,
                        suggest_ucl=final_output['Suggest UCL'],
                        suggest_lcl=final_output['Suggest LCL'],
                        static_ucl=final_output['Static UCL'],
                        static_lcl=final_output['Static LCL'],
                        cl_center=final_output['CL_Center'],
                        pattern=final_output['Pattern'],
                        total_data_count=final_output.get('TotalDataCount', len(raw_data_df)),
                        used_data_count=final_output.get('DataCountUsed', 0),
                        tsmc_ucl=chart_info_row.get('tsmc_ucl', np.nan),
                        tsmc_lcl=chart_info_row.get('tsmc_lcl', np.nan)
                    )
                    final_output['PlotFile'] = plot_filename if plot_filename else 'Plot Failed'
            except Exception as plot_error:
                import traceback
                print(f"     [Warning] 繪圖失敗: {plot_error}")
                traceback.print_exc()
                final_output['PlotFile'] = 'Plot Failed'
            
            return final_output
            
        except Exception as e:
            import traceback
            print(f"     [Error] 運行核心計算時發生錯誤: {e}")
            print(f"     [Error] 詳細錯誤追踪:")
            traceback.print_exc()
            return {
                'ChartName': chart_info_row['ChartName'], 
                'Status': 'Calculation Error', 
                'ErrorMessage': str(e),
                'PlotFile': 'Calculation Error'
            }

    def run_calculation(self, output_filename='CL_Calculation_Results.xlsx', progress_callback=None):
        """執行完整的 CL 計算流程"""
        
        if not self.chart_info_path or not os.path.exists(self.chart_info_path):
            raise ValueError(f"圖表資訊檔案不存在: {self.chart_info_path}")
            
        if not self.raw_data_dir or not os.path.exists(self.raw_data_dir):
            raise ValueError(f"原始數據目錄不存在: {self.raw_data_dir}")

        print("--- 1. 載入圖表配置 ---")
        all_charts_info = self.load_chart_information(self.chart_info_path)
        if all_charts_info.empty:
            raise ValueError("無法載入有效的圖表配置")

        self.results = []
        total_charts = len(all_charts_info)
        
        print(f"--- 2. 處理 {total_charts} 張圖表的數據 ---")
        
        for i, (index, chart_info) in enumerate(all_charts_info.iterrows()):
            # 更新進度
            if progress_callback:
                progress_callback(i + 1, total_charts)

            group_name = chart_info.get('GroupName', 'N/A')
            chart_name = chart_info.get('ChartName', 'N/A')
            
            print(f"  > 處理 Chart: {group_name}_{chart_name}...")
            
            filepath = self.find_matching_file(self.raw_data_dir, group_name, chart_name)
            
            if filepath is None:
                print(f"    [Warning] 未找到匹配的原始數據文件。跳過。")
                result = chart_info.to_dict()
                result['Status'] = 'No Raw Data'
                result['PlotFile'] = 'No Raw Data'
                self.results.append(result)
                continue

            try:
                raw_df = pd.read_csv(filepath, float_precision='round_trip')
                
                # 強制轉換 point_val 為數字型別（容錯處理）
                if 'point_val' in raw_df.columns:
                    raw_df['point_val'] = pd.to_numeric(raw_df['point_val'], errors='coerce')
                
                results_dict = self.process_single_chart_data(chart_info, raw_df)
                self.results.append(results_dict)
                
            except Exception as e:
                print(f"    [Error] 讀取數據時發生錯誤: {e}")
                error_result = chart_info.to_dict()
                error_result['Status'] = 'File Read Error'
                error_result['ErrorMessage'] = str(e)
                error_result['PlotFile'] = 'File Read Error'
                self.results.append(error_result)
                
        # --- 3. 準備輸出結果 ---
        
        df_output = pd.DataFrame(self.results)
        
        # 調整輸出欄位順序
        output_cols_priority = [
            'Figure',  # ✅ 新增：圖表欄位放在第一位
            'GroupName', 'ChartName', 'ChartID', 'Material_no', 
            'Status', 'ErrorMessage',
            'Suggest UCL', 'Suggest LCL', 
            'Static UCL', 'Static LCL',
            'Raw UCL After Resolution', 'Raw LCL After Resolution',
            'Sug USL', 'Sug LSL',
            'UCL', 'LCL',
            'Original UCL', 'Original LCL',
            'CL_Center', 'Sigma_Est', 'Sigma_Est_Upper', 'Sigma_Est_Lower',
            'Original_UCL_K_Set', 'Original_LCL_K_Set',
            'Suggest_UCL_K_Set', 'Suggest_LCL_K_Set',
            'Ori_K_Set', 'Sug_K_Set',
            'Ori_OOC_Count', 'Static_OOC_Count', 'Final_OOC_Count', 'Total_Adj_Units',
            'Target', 'USL', 'LSL', 
            'DetectionLimit', 
            'Pattern', 'Resolution_Estimated', 'Characteristics', 
            'TightenNeeded', 'Original_Tolerance', 'New_Tolerance', 'Diff_Ratio_%', 'Tighten_Threshold_%',
            'TotalDataCount', 'DataCountUsed', 'HardRule',
            'PlotFile' 
        ]
        
        # 添加空的 Figure 欄位（稍後可用於插入圖片）
        df_output.insert(0, 'Figure', '')
        
        existing_output_cols = [col for col in output_cols_priority if col in df_output.columns]
        df_output = df_output[existing_output_cols]
        
        print("\n--- 4. 計算完成 ---")
        print(f"成功處理 {len(df_output)} 張圖表")
        
        # 注意：不再自動輸出 Excel 報告
        # 如需輸出報告，請使用 export_results() 方法
        
        return df_output

    def export_results(self, results_df, output_filename='CL_Calculation_Results.xlsx'):
        """
        將計算結果匯出為 Excel 報告，包含圖表插入
        
        Args:
            results_df: 計算結果的 DataFrame（來自 run_calculation 的返回值）
            output_filename: 輸出檔案名稱
        
        Returns:
            bool: 匯出是否成功
        """
        if results_df is None or results_df.empty:
            print("沒有可匯出的結果")
            return False
            
        print("\n--- 開始匯出 Excel 報告 ---")
        
        try:
            import xlsxwriter
            import math
            
            # 重新排列欄位順序：圖片欄位在第一位
            columns = ['Figure'] + [c for c in results_df.columns if c != 'Figure']
            
            # 創建工作簿和工作表
            workbook = xlsxwriter.Workbook(output_filename)
            worksheet = workbook.add_worksheet('Results')
            
            # 設定欄寬
            worksheet.set_column(0, 0, 50)  # Figure 欄位設寬 60
            for i in range(1, len(columns)):
                if columns[i] in ['GroupName', 'ChartName']:
                    worksheet.set_column(i, i, 25)
                elif 'UCL' in columns[i] or 'LCL' in columns[i] or 'USL' in columns[i] or 'LSL' in columns[i]:
                    worksheet.set_column(i, i, 15)
                elif columns[i] == 'PlotFile':
                    worksheet.set_column(i, i, 50)
                else:
                    worksheet.set_column(i, i, 18)
            
            # 設定格式
            bold = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter'})
            cell_format = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
            
            # 寫入標題
            for col_idx, col_name in enumerate(columns):
                worksheet.write(0, col_idx, col_name, bold)
            
            # 寫入資料 & 插入圖片
            image_count = 0
            
            for row_idx, (_, row) in enumerate(results_df.iterrows()):
                excel_row = row_idx + 1  # Excel 行號（跳過標題行）
                
                # 處理圖片插入
                plot_file = row.get('PlotFile', '')
                print(f"    Row {row_idx} -> Excel Row {excel_row}: PlotFile = {plot_file}")
                
                if plot_file and isinstance(plot_file, str) and plot_file not in ['No Raw Data', 'Plot Failed', 'Calculation Error', 'File Read Error']:
                    if os.path.exists(plot_file):
                        try:
                            # 設定固定行高115像素以容納圖片
                            worksheet.set_row(excel_row, 115)
                            
                            # 插入圖片（針對900px原始圖使用固定縮放比例0.35）
                            worksheet.insert_image(excel_row, 0, plot_file, {
                                'x_scale': 0.35,
                                'y_scale': 0.35,
                                'object_position': 1,
                                'y_offset': 10
                            })
                            image_count += 1
                            print(f"      ✓ 成功插入圖片到 Excel 行 {excel_row}: {plot_file} (縮放比例: 0.18)")
                        except Exception as e:
                            print(f"      ✗ 無法插入圖片 {plot_file}: {e}")
                    else:
                        print(f"      ✗ 檔案不存在: {plot_file}")
                
                # 寫入其他欄位資料（從第2欄開始）
                for col_idx, col_name in enumerate(columns[1:], 1):
                    val = row.get(col_name, '')
                    
                    # 處理 NaN/Inf/None 問題
                    if val is None:
                        val = 'N/A'
                    elif isinstance(val, float):
                        if math.isnan(val) or math.isinf(val):
                            val = 'N/A'
                    
                    worksheet.write(excel_row, col_idx, val, cell_format)
            
            # 關閉工作簿
            workbook.close()
            
            print(f"\n--- Excel 報告匯出完成 ---")
            print(f"總共插入 {image_count} 張圖片")
            print(f"計算結果已成功輸出至：{output_filename}")
            print(f"  - 已插入圖表到 Figure 欄位")
            print(f"  - 已自動調整欄寬")
            print(f"  - 標題列已加粗")
            
            return True
            
        except Exception as e:
            print(f"\n--- Excel 報告匯出失敗 ---")
            print(f"錯誤訊息：{e}")
            return False

    def get_results(self):
        """取得計算結果"""
        return self.results if hasattr(self, 'results') else []

# === 執行入口 ===
if __name__ == '__main__':
    calculator = CLTightenCalculator(
        chart_info_path='input/All_Chart_Information.xlsx',
        raw_data_dir='input/raw_charts/'
    )
    # 執行計算（不會自動生成報告）
    results_df = calculator.run_calculation()
    
    # 如果需要匯出報告，手動調用
    # calculator.export_results(results_df, 'CL_Calculation_Results.xlsx')