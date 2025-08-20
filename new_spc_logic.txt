import sys
import os
os.environ["QT_API"] = "PyQt6" # 確認使用 PyQt6
import re
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor
import traceback
from scipy import stats  # 添加 scipy import
# Excel 和圖片處理
from openpyxl import Workbook
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.utils.dataframe import dataframe_to_rows
import xlsxwriter  # 如果你有用 xlsxwriter 存檔可以留著
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas # 習慣上會將 FigureCanvasQTAgg 命名為 FigureCanvas
from matplotlib.backends.backend_qtagg import NavigationToolbar2QT as NavigationToolbar
from matplotlib.figure import Figure
from io import BytesIO
import streamlit as st
import streamlit as st
import tempfile
from PIL import Image

def load_execution_time(raw_data_file):
    if not os.path.exists(raw_data_file):
        print(f" - load_execution_time: 檔案不存在: {raw_data_file}. 返回 None.")
        return None
    try:
        try:
            df = pd.read_excel(raw_data_file, sheet_name='Time', engine='openpyxl')
        except Exception as e:
            print(f" - load_execution_time: 無法讀取 'Time' Sheet 或檔案格式錯誤: {e}. 返回 None.")
            return None
        if df.empty:
            print(" - load_execution_time: 'Time' Sheet 是空的. 返回 None.")
            return None

        if 'execTime' not in df.columns:
            print(" - load_execution_time: 'Time' Sheet 中找不到 'execTime' 欄位. 返回 None.")
            return None
        execution_time_str = df.iloc[0]['execTime']

        if pd.isna(execution_time_str):
             print(" - load_execution_time: 'execTime' 儲存格是空的或無效. 返回 None.")
             return None

        try:
            execution_time = pd.to_datetime(execution_time_str, format='%Y-%m-%d %H:%M:%S')
            print(f" - load_execution_time: 成功讀取執行時間: {execution_time}")
            return execution_time
        except ValueError as e:
            print(f" - load_execution_time: 無法將 '{execution_time_str}' 轉換為日期時間: {e}. 返回 None.")
            return None # 轉換失敗也返回 None

    except Exception as e:
        print(f" - load_execution_time: 讀取執行時間時發生未知錯誤: {e}. 返回 None.")
        return None

def load_chart_information(raw_data_file):
    print("載入圖表信息...")
    all_charts_info = pd.read_excel(raw_data_file, sheet_name='Chart', engine='openpyxl')
    
    expected_columns = ['GroupName', 'ChartName', 'Material_no', 'CHART_CREATE_TIME', 'USL', 'LSL', 'UCL', 'LCL', 'Target', 'ChartID', 'Characteristics']
    for col in expected_columns:
        if col not in all_charts_info.columns:
            raise KeyError(f"欄位 '{col}' 不存在於圖表信息中")
    
    return all_charts_info

def preprocess_raw_df(raw_df):
    raw_df.replace([np.inf, -np.inf, 'na', 'NA', 'NaN', 'nan'], np.nan, inplace=True)
    required_columns = ['GroupName', 'ChartName', 'point_val', 'Batch_ID', 'point_time']
    missing_columns = [col for col in required_columns if col not in raw_df.columns]
    if missing_columns:
        raise ValueError(f"原始數據缺少的欄位: {missing_columns}")
    column_types = {
        'GroupName': 'str',
        'ChartName': 'str',
        'point_val': 'float',
        'Batch_ID': 'str',
        'point_time': 'str'
    }
    return raw_df.astype(column_types)

def format_datetime(dt):
    try:
        return pd.to_datetime(dt, format='%Y/%m/%d %H:%M', errors='coerce')
    except Exception as e:
        print(f"日期格式化錯誤: {e}")
        return pd.NaT

def format_and_clean_data(raw_df, chart_info):
    raw_df['point_time'] = raw_df['point_time'].apply(format_datetime)
    create_time = pd.to_datetime(chart_info['CHART_CREATE_TIME'], format="%m/%d/%Y %I:%M:%S %p", errors='coerce')
    raw_df.dropna(subset=['point_val', 'point_time'], inplace=True)
    raw_df = raw_df[raw_df['point_time'] >= create_time]
    return raw_df
def update_chart_limits(raw_df, chart_info):
    raw_df.sort_values(by='point_time', inplace=True)
    raw_df.reset_index(drop=True, inplace=True)
    
    required_columns = ['usl_val', 'lsl_val', 'ucl_val', 'lcl_val', 'target_val']
    for col in required_columns:
        if col not in raw_df.columns:
            raw_df[col] = np.nan  # 初始化欄位為 NaN
    
    raw_df[required_columns] = raw_df[required_columns].fillna({
        'usl_val': chart_info['USL'],
        'lsl_val': chart_info['LSL'],
        'ucl_val': chart_info['UCL'],
        'lcl_val': chart_info['LCL'],
        'target_val': chart_info['Target']
    })
    
    raw_df = raw_df.round(8)
    
    return raw_df, chart_info

def exclude_oos_data(raw_df):
    usl = raw_df['usl_val'].iat[0]
    lsl = raw_df['lsl_val'].iat[0]
    
    if pd.notna(usl) and pd.notna(lsl):
        return raw_df[(raw_df['point_val'] <= usl) & (raw_df['point_val'] >= lsl)]
    elif pd.isna(usl):
        return raw_df[raw_df['point_val'] >= lsl]
    elif pd.isna(lsl):
        return raw_df[raw_df['point_val'] <= usl]
    return raw_df  # 如果都沒有符合條件，則直接回傳原始資料

def preprocess_data(chart_info, raw_df):
    try:
        raw_df = format_and_clean_data(raw_df, chart_info)  # 確保這個函數已經是最佳化的
        
        if raw_df.empty:
            return False, None, None
        
        raw_df, chart_info = update_chart_limits(raw_df, chart_info)  # 確保這個函數已經是最佳化的
        raw_df = exclude_oos_data(raw_df)
        raw_df = raw_df[['point_val', 'point_time']]
        
        
        chart_info = chart_info.rename({
            'Material_no': 'material_no', 
            'GroupName': 'group_name',
            'ChartName': 'chart_name'
        })
        
        return True, raw_df, chart_info
    except ValueError as ve:
        print(f'跳過圖表處理，因爲缺少欄位: {ve}')
        return False, None, None
    except Exception as e:
        print(f'預處理過程中出錯: {e}')
        return False, None, None

def find_matching_file(directory, group_name, chart_name):
    group_name = str(group_name)
    chart_name = str(chart_name)
    
    pattern = re.compile(rf"{re.escape(group_name)}_{re.escape(chart_name)}(?:_\d+_\d+)?\.csv$")
    
    matching_files = [
        os.path.join(directory, filename)
        for filename in os.listdir(directory)
        if pattern.match(filename)
    ]
    
    return matching_files[0] if matching_files else None

def get_percentiles(values):
    values = np.array(values)  # 確保數值是 NumPy 陣列，這樣計算會更快
    return {
        'P05': np.percentile(values, 5),
        'P50': np.percentile(values, 50),
        'P75': np.percentile(values, 75),
        'P25': np.percentile(values, 25),
        'P95': np.percentile(values, 95),
        'P99.865': np.percentile(values, 99.865),
        'P0.135': np.percentile(values, 0.135)
    }

def rolling_calculation(data_values, days_to_roll):
    data_values = np.array(data_values)
    
    return data_values[-days_to_roll:] if len(data_values) >= days_to_roll else data_values

def record_high_low_calculator(current_week_data, historical_data):
    """
    判斷當週數據是否創下歷史新高或新低
    
    Args:
        current_week_data: 當週數據的 point_val 值 (array-like)
        historical_data: 歷史數據的 point_val 值 (array-like)  
    
    Returns:
        dict: 包含 record_high, record_low, highlight_status 的字典
    """
    try:
        # 快速檢查：如果任一數據集為空，直接返回
        if len(current_week_data) == 0 or len(historical_data) == 0:
            return {
                'record_high': False,
                'record_low': False, 
                'highlight_status': 'NO_HIGHLIGHT'
            }
        
        # 性能優化：使用numpy操作，避免Python循環
        current_week_data = np.asarray(current_week_data)
        historical_data = np.asarray(historical_data)
        
        # DEBUG: 輸出數據詳細信息
        print(f"  DEBUG: 當週數據點數={len(current_week_data)}, 基線數據點數={len(historical_data)}")
        print(f"  DEBUG: 當週數據前5個值={current_week_data[:5] if len(current_week_data) >= 5 else current_week_data}")
        print(f"  DEBUG: 基線數據前5個值={historical_data[:5] if len(historical_data) >= 5 else historical_data}")
        print(f"  DEBUG: 基線數據後5個值={historical_data[-5:] if len(historical_data) >= 5 else historical_data}")
        
        # 計算當週最高值和最低值 - 使用numpy的快速操作
        current_max = np.max(current_week_data)
        current_min = np.min(current_week_data)
        
        # 計算歷史最高值和最低值 - 使用numpy的快速操作
        historical_max = np.max(historical_data)
        historical_min = np.min(historical_data)
        
        # DEBUG: 輸出詳細比較信息
        print(f"  DEBUG: 當週最高值={current_max:.8f}, 歷史最高值={historical_max:.8f}")
        print(f"  DEBUG: 當週最低值={current_min:.8f}, 歷史最低值={historical_min:.8f}")
        print(f"  DEBUG: 最高值差異={current_max - historical_max:.8f}")
        print(f"  DEBUG: 最低值差異={current_min - historical_min:.8f}")
        
        # 檢查當週數據是否包含歷史極值
        current_has_hist_max = np.any(current_week_data == historical_max)
        current_has_hist_min = np.any(current_week_data == historical_min)
        print(f"  DEBUG: 當週數據是否包含歷史最高值={current_has_hist_max}")
        print(f"  DEBUG: 當週數據是否包含歷史最低值={current_has_hist_min}")
        
        # 判斷是否創下新高或新低 - 簡單的數值比較，非常快速
        record_high = current_max > historical_max
        record_low = current_min < historical_min
        
        # 如果創下新高或新低，則需要高亮顯示
        highlight_status = 'HIGHLIGHT' if (record_high or record_low) else 'NO_HIGHLIGHT'
        
        print(f"  record_high_low_calculator: 當週最高={current_max:.4f}, 歷史最高={historical_max:.4f}, 創新高={record_high}")
        print(f"  record_high_low_calculator: 當週最低={current_min:.4f}, 歷史最低={historical_min:.4f}, 創新低={record_low}")
        print(f"  record_high_low_calculator: 高亮狀態={highlight_status}")
        
        return {
            'record_high': record_high,
            'record_low': record_low,
            'highlight_status': highlight_status
        }
        
    except Exception as e:
        print(f"  record_high_low_calculator: 計算過程中發生錯誤: {e}")
        return {
            'record_high': False,
            'record_low': False,
            'highlight_status': 'NO_HIGHLIGHT'
        }
def review_kshift_results(results, resolution, characteristic, data_percentiles, base_percentiles):
    highlight_conditions = {key: 'NO_HIGHLIGHT' for key in ['P95_shift', 'P50_shift', 'P05_shift']}

    for percentile in ['P95', 'P50', 'P05']:
        k_value = results.get(f'{percentile}_k', np.nan) # 使用 .get 安全獲取 K 值 (絕對值)

        data_p = data_percentiles.get(percentile, np.nan)
        base_p = base_percentiles.get(percentile, np.nan)
        abs_diff = np.nan # 預設絕對差值為 NaN

        if not pd.isna(data_p) and not pd.isna(base_p):
            abs_diff = abs(data_p - base_p)

        is_significant_diff = not pd.isna(abs_diff) and abs_diff >= resolution

        # 判斷 K 絕對值是否超過 2 (且非NaN)
        is_significant_k = not pd.isna(k_value) and abs(k_value) > 2

        # 新增判斷 k_value 是否為無限值
        is_infinite_k = not pd.isna(k_value) and np.isinf(abs(k_value))

        # 設定初始高亮: 絕對差值 > resolution 且 (K絕對值 > 2 或 K絕對值為無限)
        if is_significant_diff and (is_significant_k or is_infinite_k):  # 使用 AND 和 OR 結合邏輯
            highlight_conditions[f'{percentile}_shift'] = 'HIGHLIGHT'   

    if characteristic == 'Bigger':
        # 檢查 data_percentiles 和 base_percentiles 的鍵是否存在且值非空
        if data_percentiles.get('P95') is not None and base_percentiles.get('P05') is not None and data_percentiles['P95'] >= base_percentiles['P05']:
            highlight_conditions['P95_shift'] = 'NO_HIGHLIGHT'
        if data_percentiles.get('P50') is not None and base_percentiles.get('P25') is not None and data_percentiles['P50'] >= base_percentiles['P25']:
            highlight_conditions['P50_shift'] = 'NO_HIGHLIGHT'
        # 檢查 results 的鍵是否存在且值非空
        if results.get('P95_k_ori') is not None and results['P95_k_ori'] >= 0:
            highlight_conditions['P95_shift'] = 'NO_HIGHLIGHT'
        if results.get('P50_k_ori') is not None and results['P50_k_ori'] >= 0:
            highlight_conditions['P50_shift'] = 'NO_HIGHLIGHT'
        if results.get('P05_k_ori') is not None and results['P05_k_ori'] >= 0:
            highlight_conditions['P05_shift'] = 'NO_HIGHLIGHT'

    elif characteristic == 'Smaller':
        if data_percentiles.get('P05') is not None and base_percentiles.get('P95') is not None and data_percentiles['P05'] <= base_percentiles['P95']:
            highlight_conditions['P05_shift'] = 'NO_HIGHLIGHT'
        if data_percentiles.get('P50') is not None and base_percentiles.get('P75') is not None and data_percentiles['P50'] <= base_percentiles['P75']:
            highlight_conditions['P50_shift'] = 'NO_HIGHLIGHT'
        if results.get('P95_k_ori') is not None and results['P95_k_ori'] <= 0:
            highlight_conditions['P95_shift'] = 'NO_HIGHLIGHT'
        if results.get('P50_k_ori') is not None and results['P50_k_ori'] <= 0:
            highlight_conditions['P50_shift'] = 'NO_HIGHLIGHT'
        if results.get('P05_k_ori') is not None and results['P05_k_ori'] <= 0:
            highlight_conditions['P05_shift'] = 'NO_HIGHLIGHT'

    elif characteristic == 'Nominal':
        if data_percentiles.get('P95') is not None and base_percentiles.get('P95') is not None and data_percentiles['P95'] <= base_percentiles['P95']:
            highlight_conditions['P95_shift'] = 'NO_HIGHLIGHT'
        if data_percentiles.get('P05') is not None and base_percentiles.get('P05') is not None and data_percentiles['P05'] >= base_percentiles['P05']:
            highlight_conditions['P05_shift'] = 'NO_HIGHLIGHT'
        # 檢查 P25, P50, P75 的鍵是否存在且值非空
        if (base_percentiles.get('P25') is not None and
            data_percentiles.get('P50') is not None and
            base_percentiles.get('P75') is not None and
            base_percentiles['P25'] <= data_percentiles['P50'] <= base_percentiles['P75']):
            highlight_conditions['P50_shift'] = 'NO_HIGHLIGHT'
        if results.get('P95_k_ori') is not None and results['P95_k_ori'] <= 0:
            highlight_conditions['P95_shift'] = 'NO_HIGHLIGHT'
        if results.get('P05_k_ori') is not None and results['P05_k_ori'] >= 0:
            highlight_conditions['P05_shift'] = 'NO_HIGHLIGHT'

    return highlight_conditions


def safe_division(numerator, denominator, epsilon=1e-9):
    """
    執行安全除法，避免除以零或極小值。
    如果分母接近零，返回 np.nan。
    """
    if abs(denominator) < epsilon:
        return np.nan # 或者返回 float('inf'), 根據您希望在結果中如何表示這種情況
    return np.round(numerator, 8) / denominator


def kshift_sigma_ratio_calculator(base, data, characteristic, resolution, ucl, lcl):

    results = {
        'P95_k': np.nan,
        'P50_k': np.nan,
        'P05_k': np.nan,
        # 確保所有 results 的鍵都有初始值，包括 review_kshift_results 返回的
        'P95_k_ori': np.nan,
        'P50_k_ori': np.nan,
        'P05_k_ori': np.nan,
        'P95_shift': 'NO_HIGHLIGHT',
        'P50_shift': 'NO_HIGHLIGHT',
        'P05_shift': 'NO_HIGHLIGHT'
    }

    print("--- 進入 kshift_sigma_ratio_calculator 函數 ---")

    if 'values' not in base or 'values' not in data:
         print("  kshift: 錯誤：輸入數據字典缺少 'values' 鍵。")
         return pd.Series(results)

    data_values = data['values']
    base_values = base['values']

    data_cnt = len(data_values)
    base_cnt = len(base_values) # 也獲取基線數據長度

    print(f"  kshift: 接收到的 data_values shape: {data_values.shape}, base_values shape: {base_values.shape}")
    print(f"  kshift: data_cnt: {data_cnt}, base_cnt: {base_cnt}")

    # 如果週數據少於 1 個點，直接返回預設結果
    if data_cnt < 1:
        print("  kshift: data_cnt < 1, 返回預設結果。")
        return pd.Series(results)

    # 計算基線百分位數。請確保 get_percentiles 能處理 base_cnt = 3 的情況
    try:
        base_percentiles = get_percentiles(base_values)
        print(f"  kshift: 計算出的 base_percentiles (部分): P05={base_percentiles.get('P05')}, P50={base_percentiles.get('P50')}, P95={base_percentiles.get('P95')}")
        # 檢查計算分母所需的關鍵百分位數是否存在且不是 NaN
        if np.isnan(base_percentiles.get('P99.865', np.nan)) or np.isnan(base_percentiles.get('P0.135', np.nan)) or np.isnan(base_percentiles.get('P50', np.nan)):
             print("  kshift: 警告：基線百分位數計算結果無效 (包含 NaN)，可能基線數據不足。無法計算 K 值。")
             return pd.Series(results) # 無法計算分母，返回預設結果

    except Exception as e:
         print(f"  kshift: 計算基線百分位數時發生錯誤: {e}")
         traceback.print_exc()
         return pd.Series(results)


    rolled_data = None # 預設沒有滾動數據
    data_percentiles = None # 預設沒有當前週數據的百分位數

    if data_cnt == 1:
        print("  kshift: 處理 data_cnt == 1 分支 (週數據只有 1 點)")
        days_to_roll = 1
        rolled_data_values = np.copy(data_values) # 從單點週數據開始

        # 修正無限迴圈：在 base_values 變空時跳出迴圈
        # 並且確保合併後的數據長度達到 5
        while len(rolled_data_values) < 5:
            print(f"  kshift: While 迴圈開始, days_to_roll: {days_to_roll}, rolled_data_values len: {len(rolled_data_values)}")

            # 如果 base_values 已經是空的，無法再滾動，跳出
            if len(base_values) == 0:
                print("  kshift: base_values 已經是空的，無法再滾動。跳出迴圈。")
                break

            # 呼叫 rolling_calculation 前
            print(f"  kshift: 呼叫 rolling_calculation 前, base_values shape: {base_values.shape}, days_to_roll: {days_to_roll}")
            try:
                rolled_base_values = rolling_calculation(base_values, days_to_roll)
                print(f"  kshift: rolling_calculation 返回 shape: {rolled_base_values.shape}")

                # 合併數據
                print(f"  kshift: 合併 rolled_data_values ({rolled_data_values.shape}) 與 rolled_base_values ({rolled_base_values.shape})")
                rolled_data_values = np.concatenate((rolled_data_values, rolled_base_values))
                print(f"  kshift: 合併後 rolled_data_values shape: {rolled_data_values.shape}")

                # 縮短 base_values，用於下一次迴圈
                base_values = base_values[:-days_to_roll]
                print(f"  kshift: 縮短後 base_values shape: {base_values.shape}")

            except Exception as e:
                 print(f"  kshift: rolling_calculation 或合併數據時發生錯誤: {e}")
                 traceback.print_exc()
                 # 如果發生錯誤，可能無法繼續，返回預設結果
                 return pd.Series(results)


            days_to_roll += 1
            # 保持安全跳出機制，防止意外情況
            if days_to_roll > base_cnt + 10: # 如果 rolling 天數遠超過原始基線數據，可能出錯
                print("  kshift: 警告：rolling 迴圈 days_to_roll 過大，可能邏輯有誤或數據問題。強制跳出。")
                break


        print(f"  kshift: While 迴圈結束。最終 rolled_data_values shape: {rolled_data_values.shape}")

        # 迴圈結束後，檢查是否湊滿了至少 5 個點用於滾動計算
        if len(rolled_data_values) < 5:
             print(f"  kshift: 警告：無法湊滿至少 5 個點用於滾動計算 (實際湊到 {len(rolled_data_values)} 點)。滾動結果將使用現有數據，計算可能不穩定。")
             # 您可以根據需求決定如果點數不足 5 時是否返回預設結果
             return pd.Series(results) # 如果少於 5 點則視為無法計算並返回預設值


        # 計算百分位數 (使用原始單點數據和滾動/填充後的數據)
        try:
            data_percentiles = get_percentiles(data_values) # 這是用原始單點週數據算的
            print(f"  kshift: 原始週數據 percentiles (data_cnt=1): {data_percentiles}")

            rolled_data_percentiles = get_percentiles(rolled_data_values) # 這是用滾動後的數據算的
            print(f"  kshift: 滾動數據 percentiles (shape={rolled_data_values.shape}): {rolled_data_percentiles}")

            rolled_data = {'values': rolled_data_values, 'percentiles': rolled_data_percentiles}

            # 檢查計算K值所需的當前和滾動數據百分位數是否存在且非NaN
            for p in ['P95', 'P50', 'P05']:
                if np.isnan(data_percentiles.get(p, np.nan)):
                     print(f"  kshift: 警告：原始週數據 {p} 百分位數為 NaN。無法計算 K 值。")
                     return pd.Series(results)
                if np.isnan(rolled_data_percentiles.get(p, np.nan)):
                     print(f"  kshift: 警告：滾動數據 {p} 百分位數為 NaN。影響滾動 K 值計算。") # 這裡只是警告，可能可以繼續

        except Exception as e:
            print(f"  kshift: 計算百分位數時發生錯誤 (data_cnt=1 分支): {e}")
            traceback.print_exc()
            return pd.Series(results)


    elif data_cnt >= 2:
        print(f"  kshift: 處理 data_cnt >= 2 分支, data_cnt: {data_cnt}")
        try:
             data_percentiles = get_percentiles(data_values)
             print(f"  kshift: 當前週數據 percentiles (data_cnt>1): {data_percentiles}")
             # 檢查計算K值所需的當前百分位數是否存在且非NaN
             for p in ['P95', 'P50', 'P05']:
                 if np.isnan(data_percentiles.get(p, np.nan)):
                      print(f"  kshift: 警告：當前週數據 {p} 百分位數為 NaN。無法計算 K 值。")
                      return pd.Series(results)

        except Exception as e:
            print(f"  kshift: 計算百分位數時發生錯誤 (data_cnt>=2 分支): {e}")
            traceback.print_exc()
            return pd.Series(results)


        rolled_data = None # data_cnt >= 2 時，沒有滾動數據的概念用於 highlight 判斷
    else: # 這個分支理論上不會走到，因為開頭已經處理 data_cnt < 1
        print(f"  kshift: Warning: 未預期的 data_cnt 情況: {data_cnt}")
        return pd.Series(results)

    # --- 計算分母 ---
    try:
        # 計算分母，加入安全除法和替代邏輯
        p95k_deno = np.round(np.max([safe_division(base_percentiles.get('P99.865', np.nan) - base_percentiles.get('P50', np.nan), 3),
                                     safe_division(ucl - base_percentiles.get('P50', np.nan), 6)]), 8)
        p50k_deno = np.round(np.max([safe_division(base_percentiles.get('P99.865', np.nan) - base_percentiles.get('P0.135', np.nan), 6),
                                     safe_division(ucl - lcl, 12)]), 8)
        p05k_deno = np.round(np.max([safe_division(base_percentiles.get('P50', np.nan) - base_percentiles.get('P0.135', np.nan), 3),
                                     safe_division(base_percentiles.get('P50', np.nan) - lcl, 6)]), 8)

        # YC edit：分母為 0 時的處理邏輯
        if p95k_deno == 0:
            if p05k_deno == 0:
                p95k_deno = p50k_deno
            elif p50k_deno == 0:
                p95k_deno = p05k_deno
            else:
                p95k_deno = min(p50k_deno, p05k_deno)
        if p05k_deno == 0:
            if p95k_deno == 0:
                p05k_deno = p50k_deno
            elif p50k_deno == 0:
                p05k_deno = p95k_deno
            else:
                p05k_deno = min(p50k_deno, p95k_deno)
        if p50k_deno == 0:
            if p95k_deno == 0:
                p50k_deno = p05k_deno
            elif p05k_deno == 0:
                p50k_deno = p95k_deno
            else:
                p50k_deno = min(p05k_deno, p95k_deno)

        denominators = {
            'p95k_deno': p95k_deno,
            'p50k_deno': p50k_deno,
            'p05k_deno': p05k_deno
        }
        print(f"  kshift: 計算出的分母: {denominators}")

        # 檢查分母是否有效 (非 NaN, 非 Inf)
        if np.isnan(p95k_deno) or np.isnan(p50k_deno) or np.isnan(p05k_deno) or np.isinf(p95k_deno) or np.isinf(p50k_deno) or np.isinf(p05k_deno):
            print("  kshift: 警告：計算出的分母無效 (包含 NaN 或 Inf)。無法計算 K 值。")
            return pd.Series(results)

    except Exception as e:
        print(f"  kshift: 計算分母時發生錯誤: {e}")
        traceback.print_exc()
        return pd.Series(results)
    # --- 計算 K 值 ---
    try:
        # 計算 K 值 (原始) - 使用安全除法
        results['P95_k_ori'] = safe_division(np.round(data_percentiles.get('P95', np.nan) - base_percentiles.get('P95', np.nan), 8), p95k_deno)
        results['P50_k_ori'] = safe_division(np.round(data_percentiles.get('P50', np.nan) - base_percentiles.get('P50', np.nan), 8), p50k_deno)
        results['P05_k_ori'] = safe_division(np.round(data_percentiles.get('P05', np.nan) - base_percentiles.get('P05', np.nan), 8), p05k_deno)

        # 計算 K 值 (絕對值) - 使用安全除法
        results['P95_k'] = safe_division(np.round(abs(data_percentiles.get('P95', np.nan) - base_percentiles.get('P95', np.nan)), 8), p95k_deno)
        results['P50_k'] = safe_division(np.round(abs(data_percentiles.get('P50', np.nan) - base_percentiles.get('P50', np.nan)), 8), p50k_deno)
        results['P05_k'] = safe_division(np.round(abs(data_percentiles.get('P05', np.nan) - base_percentiles.get('P05', np.nan)), 8), p05k_deno)

        print(f"  kshift: 計算出的 K 值結果: {results}")

    except Exception as e:
        print(f"  kshift: 計算 K 值時發生錯誤: {e}")
        traceback.print_exc()
        return pd.Series(results)

    # --- 判斷當前高亮條件 ---
    try:
        # 確保傳給 review_kshift_results 的 percentiles 字典是完整的
        current_highlight_conditions = review_kshift_results(results, resolution, characteristic, data_percentiles, base_percentiles)
        print(f"  kshift: current_highlight_conditions: {current_highlight_conditions}")
    except Exception as e:
        print(f"  kshift: 判斷當前高亮條件時發生錯誤: {e}")
        traceback.print_exc()
        # 如果判斷高亮失敗，相關結果可能不準確，但可以返回計算出的 K 值
        current_highlight_conditions = {key: 'ERROR' for key in ['P95_shift', 'P50_shift', 'P05_shift']} # 用 ERROR 標記

    # --- 計算滾動結果高亮條件 (如果存在滾動數據) ---
    rolling_highlight_conditions = {key: 'NO_HIGHLIGHT' for key in ['P95_shift', 'P50_shift', 'P05_shift']}

    if rolled_data is not None:
        print(f"  kshift: 處理 rolled_data != None 分支，rolled_data shape: {rolled_data['values'].shape}")
        print(f"  kshift: 滾動後 base_percentiles: {base_percentiles}")
        try:
            # 計算滾動結果 (K 值) - 使用安全除法
            rolling_results = {
                'P95_k': safe_division(np.round(abs(rolled_data['percentiles'].get('P95', np.nan) - base_percentiles.get('P95', np.nan)), 8), p95k_deno),
                'P50_k': safe_division(np.round(abs(rolled_data['percentiles'].get('P50', np.nan) - base_percentiles.get('P50', np.nan)), 8), p50k_deno),
                'P05_k': safe_division(np.round(abs(rolled_data['percentiles'].get('P05', np.nan) - base_percentiles.get('P05', np.nan)), 8), p05k_deno),
                'P95_k_ori': safe_division(np.round((rolled_data['percentiles'].get('P95', np.nan) - base_percentiles.get('P95', np.nan)), 8), p95k_deno),
                'P50_k_ori': safe_division(np.round((rolled_data['percentiles'].get('P50', np.nan) - base_percentiles.get('P50', np.nan)), 8), p50k_deno),
                'P05_k_ori': safe_division(np.round((rolled_data['percentiles'].get('P05', np.nan) - base_percentiles.get('P05', np.nan)), 8), p05k_deno),
            }
            print(f"  kshift: 計算出的 rolling_results: {rolling_results}")

            # 判斷滾動結果高亮條件
            # 確保傳給 review_kshift_results 的 percentiles 字典是完整的
            rolling_highlight_conditions = review_kshift_results(rolling_results, resolution, characteristic, rolled_data['percentiles'], base_percentiles)
            print(f"  kshift: rolling_highlight_conditions: {rolling_highlight_conditions}")

        except Exception as e:
            print(f"  kshift: 判斷滾動高亮條件時發生錯誤: {e}")
            traceback.print_exc()
            # 如果判斷滾動高亮失敗，用 ERROR 標記
            rolling_highlight_conditions = {key: 'ERROR' for key in ['P95_shift', 'P50_shift', 'P05_shift']}

    results['P95_shift'] = 'HIGHLIGHT' if current_highlight_conditions.get('P95_shift') == 'HIGHLIGHT' and (rolled_data is None or rolling_highlight_conditions.get('P95_shift') == 'HIGHLIGHT') else 'NO_HIGHLIGHT'
    results['P50_shift'] = 'HIGHLIGHT' if current_highlight_conditions.get('P50_shift') == 'HIGHLIGHT' and (rolled_data is None or rolling_highlight_conditions.get('P50_shift') == 'HIGHLIGHT') else 'NO_HIGHLIGHT'
    results['P05_shift'] = 'HIGHLIGHT' if current_highlight_conditions.get('P05_shift') == 'HIGHLIGHT' and (rolled_data is None or rolling_highlight_conditions.get('P05_shift') == 'HIGHLIGHT') else 'NO_HIGHLIGHT'

    print(f"  kshift: 最終 shift 結果: P95={results['P95_shift']}, P50={results['P50_shift']}, P05={results['P05_shift']}")
    print("--- 退出 kshift_sigma_ratio_calculator 函數 ---")

    return pd.Series(results)

# 數據類型判斷
def determine_data_type(data_values):
    """
    判斷數據是離散型還是連續型
    
    判斷標準：
    1. (unique數值種類/總樣本數N < 1/3 且 unique數值種類 < 5) OR
    2. (總樣本數N >= 30 且 unique數值種類 <= 10)
    滿足以上任一條件即認定為離散型
    
    Parameters:
    - data_values: 數據值的 numpy array 或 pandas Series
    
    Returns:
    - 'discrete' 或 'continuous'
    """
    
    # 移除 NaN 值
    clean_values = data_values.dropna() if hasattr(data_values, 'dropna') else data_values[~np.isnan(data_values)]
    
    if len(clean_values) == 0:
        print("  數據類型判斷: 沒有有效數據，預設為連續型")
        return 'continuous'
    
    unique_values = np.unique(clean_values)
    unique_count = len(unique_values)
    total_count = len(clean_values)
    unique_ratio = unique_count / total_count
    
    print(f"  數據類型判斷: 唯一值數量={unique_count}, 總數量={total_count}, 比例={unique_ratio:.3f}")
    
    # 判斷邏輯：
    # 條件1: unique數值種類/總樣本數N < 1/3 且 unique數值種類 < 5
    condition1 = (unique_ratio <= 1/3) and (unique_count <= 5)
    
    # 條件2: 總樣本數N >= 30 且 unique數值種類 <= 10
    condition2 = (total_count >= 30) and (unique_count <= 10)
    
    if condition1 or condition2:
        print("  數據類型判斷: 判定為離散型")
        return 'discrete'
    else:
        print("  數據類型判斷: 判定為連續型")
        return 'continuous'

# OOC計算
def ooc_calculator(data, ucl, lcl):
    data_cnt = len(data)
    ooc_cnt = ((data['point_val'] > ucl) | (data['point_val'] < lcl)).sum()
    ooc_ratio = ooc_cnt / data_cnt if data_cnt != 0 else 0
    return data_cnt, ooc_cnt, ooc_ratio

# OOC結果檢查
def review_ooc_results(ooc_cnt, ooc_ratio, threshold=0.05):
    return 'HIGHLIGHT' if ooc_ratio > threshold and ooc_cnt > 1 else 'NO_HIGHLIGHT'

# 計算Sticking Rate
def sticking_rate_calculator(baseline_data, weekly_data):
    def get_mode(data):
        return data.mode()[0]

    def get_percentage(data, value):
        return (data == value).sum() / len(data)

    # 確保輸入是 pandas Series
    if isinstance(baseline_data, np.ndarray):
        baseline_data = pd.Series(baseline_data)
    if isinstance(weekly_data, np.ndarray):
        weekly_data = pd.Series(weekly_data)

    # 如果週資料少於10筆，與基線資料進行合併
    if len(weekly_data) < 10:
        rolling_window_size = 20 if len(baseline_data) > 1000 else 10
        # 修正: 使用 iloc[-rolling_window_size:] 而不是 .tail()
        baseline_tail = baseline_data.iloc[-rolling_window_size:] if len(baseline_data) >= rolling_window_size else baseline_data
        weekly_data = pd.concat([baseline_tail, weekly_data])

    threshold = 0.7
    baseline_mode = get_mode(baseline_data)
    weekly_mode = get_mode(weekly_data)

    baseline_mode_percentage_in_baseline = get_percentage(baseline_data, baseline_mode)
    baseline_mode_percentage_in_weekly = get_percentage(weekly_data, baseline_mode)
    weekly_mode_percentage_in_baseline = get_percentage(baseline_data, weekly_mode)
    weekly_mode_percentage_in_weekly = get_percentage(weekly_data, weekly_mode)

    baseline_mode_diff = abs(baseline_mode_percentage_in_baseline - baseline_mode_percentage_in_weekly)
    weekly_mode_diff = abs(weekly_mode_percentage_in_baseline - weekly_mode_percentage_in_weekly)

    highlight_needed = (baseline_mode_diff >= threshold) or (weekly_mode_diff >= threshold)
    highlight_status = 'HIGHLIGHT' if highlight_needed else 'NO_HIGHLIGHT'

    return {
        'baseline_mode': baseline_mode,
        'weekly_mode': weekly_mode,
        'baseline_mode_percentage_in_baseline': baseline_mode_percentage_in_baseline,
        'baseline_mode_percentage_in_weekly': baseline_mode_percentage_in_weekly,   
        'weekly_mode_percentage_in_baseline': weekly_mode_percentage_in_baseline,
        'weekly_mode_percentage_in_weekly': weekly_mode_percentage_in_weekly,
        'highlight_status': highlight_status
    }

# 趨勢檢查
def trending(raw_df, weekly_start_date, weekly_end_date, baseline_start_date, baseline_end_date):
    # 時間欄位轉換
    raw_df['point_time'] = pd.to_datetime(raw_df['point_time'])
    weekly_end_date = pd.to_datetime(weekly_end_date)
    baseline_start_date = pd.to_datetime(baseline_start_date)
    baseline_end_date = pd.to_datetime(baseline_end_date)

    # 每週資料的摘要
    weekly_summary = []
    current_end = weekly_end_date
    week_count = 0

    while week_count < 7:
        current_start = current_end - timedelta(days=6)
        week_data = raw_df[
            (raw_df['point_time'] >= current_start) &
            (raw_df['point_time'] <= current_end)
        ]['point_val']

        weekly_summary.append({
            'week_start': current_start,
            'week_end': current_end,
            'median': week_data.median() if not week_data.empty else np.nan,
            'count': len(week_data)
        })

        current_end = current_start - timedelta(days=1)
        week_count += 1

    weekly_data = pd.DataFrame(weekly_summary)

    if weekly_data.empty:
        return 'NO_HIGHLIGHT'

    weekly_medians = weekly_data['median'].tolist()
    weekly_counts = weekly_data['count'].tolist()

    # 檢查最近幾週的資料點數條件
    def check_weeks_condition(weeks_counts):
        if len(weeks_counts) >= 4 and sum(x >= 10 for x in weeks_counts[:4]) >= 3 and weeks_counts[0] >= 10:
            return 4
        elif len(weeks_counts) >= 5 and sum(x >= 6 for x in weeks_counts[:5]) >= 4 and weeks_counts[0] >= 6:
            return 5
        elif len(weeks_counts) >= 6 and sum(x >= 3 for x in weeks_counts[:6]) >= 5 and weeks_counts[0] >= 3:
            return 6
        elif len(weeks_counts) >= 7 and sum(x >= 1 for x in weeks_counts[:7]) >= 6 and weeks_counts[0] >= 1:
            return 7
        return 0

    num_weeks_to_check = check_weeks_condition(weekly_counts)

    if num_weeks_to_check == 0:
        return 'NO_HIGHLIGHT'

    # 趨勢檢查函式
    def is_trending_up(medians):
        return all(earlier > later for earlier, later in zip(medians, medians[1:]))

    def is_trending_down(medians):
        return all(earlier < later for earlier, later in zip(medians, medians[1:]))

    # 基準區間百分位
    baseline_df = raw_df[
        (raw_df['point_time'] >= baseline_start_date) &
        (raw_df['point_time'] <= baseline_end_date)
    ]
    baseline_values = baseline_df['point_val']

    if baseline_values.empty:
        return 'NO_HIGHLIGHT'

    p95 = np.percentile(baseline_values, 95)
    p05 = np.percentile(baseline_values, 5)

    # 檢查是否上升或下降
    check_medians = [m for m in weekly_medians[:num_weeks_to_check] if not np.isnan(m)]

    if len(check_medians) < 2:
        return 'NO_HIGHLIGHT'  # 資料不夠比趨勢

    if is_trending_up(check_medians) and check_medians[0] > p95:
        return 'HIGHLIGHT'
    elif is_trending_down(check_medians) and check_medians[0] < p05:
        return 'HIGHLIGHT'
    return 'NO_HIGHLIGHT'

# 新的 category_LT_Shift 函數
def category_lt_shift_calculator(base_data, weekly_data, threshold=0.7):
    """
    計算 category_LT_Shift
    
    邏輯：
    1. 當周<20則rolling to 20筆
    2. 拿當周data範圍去對應baseline同樣data範圍
    3. 檢查data所佔比例是否超過70%
    
    Parameters:
    - base_data: 基線數據字典
    - weekly_data: 週數據字典  
    - threshold: 佔比差異閾值，預設0.7 (70%)
    
    Returns:
    - dict: 包含highlight_status的結果字典
    """
    
    print("  category_LT_shift: 開始計算")
    
    result = {
        'highlight_status': 'NO_HIGHLIGHT',
        'weekly_range': None,
        'baseline_ratio_in_range': None,
        'weekly_ratio_in_range': None, 
        'ratio_diff': None
    }
    
    try:
        weekly_values = weekly_data['values'].copy()
        base_values = base_data['values'].copy()
        
        print(f"  category_LT_shift: 原始當周點數 = {len(weekly_values)}")
        
        # 1. 如果當周 < 20 則 rolling to 20筆
        if len(weekly_values) < 20:
            print(f"  category_LT_shift: 當周點數 < 20，rolling 到 20 筆")
            
            # 需要從基線數據中補充
            needed_points = 20 - len(weekly_values)
            if len(base_values) >= needed_points:
                # 取基線最後的點來補充
                additional_points = base_values[-needed_points:]
                weekly_values = np.concatenate([additional_points, weekly_values])
                print(f"  category_LT_shift: 補充後當周點數 = {len(weekly_values)}")
            else:
                print(f"  category_LT_shift: 基線數據不足以補充到20筆，使用現有數據")
                weekly_values = np.concatenate([base_values, weekly_values])
        
        # 2. 計算當周數據範圍
        weekly_min = np.min(weekly_values)
        weekly_max = np.max(weekly_values)
        result['weekly_range'] = (weekly_min, weekly_max)
        
        print(f"  category_LT_shift: 當周數據範圍 = [{weekly_min:.3f}, {weekly_max:.3f}]")
        
        # 3. 計算基線數據在此範圍內的比例
        baseline_in_range = base_values[(base_values >= weekly_min) & (base_values <= weekly_max)]
        baseline_ratio = len(baseline_in_range) / len(base_values) if len(base_values) > 0 else 0
        result['baseline_ratio_in_range'] = baseline_ratio
        
        # 4. 計算當周數據在此範圍內的比例（應該是100%，因為就是用當周數據定義的範圍）
        weekly_in_range = weekly_values[(weekly_values >= weekly_min) & (weekly_values <= weekly_max)]  
        weekly_ratio = len(weekly_in_range) / len(weekly_values) if len(weekly_values) > 0 else 0
        result['weekly_ratio_in_range'] = weekly_ratio
        
        # 5. 計算比例差異
        ratio_diff = abs(weekly_ratio - baseline_ratio)
        result['ratio_diff'] = ratio_diff
        
        print(f"  category_LT_shift: 基線在範圍內比例 = {baseline_ratio:.3f}")
        print(f"  category_LT_shift: 當周在範圍內比例 = {weekly_ratio:.3f}")
        print(f"  category_LT_shift: 比例差異 = {ratio_diff:.3f}")
        
        # 6. 判斷是否需要高亮
        if ratio_diff > threshold:
            result['highlight_status'] = 'HIGHLIGHT'
            print(f"  category_LT_shift: 比例差異 {ratio_diff:.3f} > {threshold}，需要 HIGHLIGHT")
        else:
            result['highlight_status'] = 'NO_HIGHLIGHT' 
            print(f"  category_LT_shift: 比例差異 {ratio_diff:.3f} <= {threshold}，NO_HIGHLIGHT")
            
    except Exception as e:
        print(f"  category_LT_shift: 計算時發生錯誤: {e}")
        traceback.print_exc()
        result['highlight_status'] = 'NO_HIGHLIGHT'
    
    return result

# 離散型 OOB 處理函數
def discrete_oob_calculator(base_data, weekly_data, chart_info):
    """
    離散型數據的 OOB 計算方法
    包含修改後的 k-shift 和新增的 category_LT_Shift
    
    Parameters:
    - base_data: 基線數據字典 (包含 'values', 'cnt', 'mean', 'sigma')
    - weekly_data: 週數據字典 (包含 'values', 'cnt', 'mean', 'sigma')
    - chart_info: 圖表信息

    
    Returns:
    - dict: 包含 OOB 結果的字典
    """
    from scipy import stats
    
    print(f"  離散型 OOB 計算: 基線數據點數={base_data['cnt']}, 週數據點數={weekly_data['cnt']}")
    
    results = {
        'HL_P95_shift': 'NO_HIGHLIGHT',
        'HL_P50_shift': 'NO_HIGHLIGHT', 
        'HL_P05_shift': 'NO_HIGHLIGHT',
        'HL_sticking_shift': 'NO_HIGHLIGHT',
        'HL_trending': 'NO_HIGHLIGHT',
        'HL_high_OOC': 'NO_HIGHLIGHT',
        'HL_category_LT_shift': 'NO_HIGHLIGHT',
        'discrete_method': True
    }
    
    try:
        # 1. Sticking Rate 計算
        print("  離散型 OOB: 計算 Sticking Rate...")
        sticking_rate_results = sticking_rate_calculator(base_data['values'], weekly_data['values'])
        results['HL_sticking_shift'] = sticking_rate_results.get('highlight_status', 'NO_HIGHLIGHT')
        
        # 2. Trending 目前暫時設為 NO_HIGHLIGHT (離散型可能需要特別處理)
        print("  離散型 OOB: Trending 暫時設為 NO_HIGHLIGHT")
        results['HL_trending'] = 'NO_HIGHLIGHT'  # 可能需要額外處理
        
        # 3. OOC 計算
        print("  離散型 OOB: 計算 OOC...")
        weekly_df = pd.DataFrame({'point_val': weekly_data['values']})
        ucl = chart_info.get('UCL', chart_info.get('ucl_val', np.nan))
        lcl = chart_info.get('LCL', chart_info.get('lcl_val', np.nan))
        
        data_cnt, ooc_cnt, ooc_ratio = ooc_calculator(weekly_df, ucl, lcl)
        ooc_highlight = review_ooc_results(ooc_cnt, ooc_ratio)
        results['HL_high_OOC'] = ooc_highlight
        
        # 4. K-shift 計算 (離散型專用版本，加入 capping rule)
        print("  離散型 OOB: 計算離散型 K-shift...")
        characteristic = chart_info.get('Characteristics', 'Nominal')
        resolution = 0.1  # 預設解析度
        
        kshift_results = discrete_kshift_calculator(base_data, weekly_data, characteristic, resolution, ucl, lcl)
        results['HL_P95_shift'] = kshift_results.get('P95_shift', 'NO_HIGHLIGHT')
        results['HL_P50_shift'] = kshift_results.get('P50_shift', 'NO_HIGHLIGHT')
        results['HL_P05_shift'] = kshift_results.get('P05_shift', 'NO_HIGHLIGHT')
        
        # 5. 新增的 category_LT_Shift 計算
        print("  離散型 OOB: 計算 category_LT_Shift...")
        category_lt_results = category_lt_shift_calculator(base_data, weekly_data)
        results['HL_category_LT_shift'] = category_lt_results.get('highlight_status', 'NO_HIGHLIGHT')
        
        print(f"  離散型 OOB 計算完成: {results}")
        
    except Exception as e:
        print(f"  離散型 OOB 計算錯誤: {e}")
        traceback.print_exc()
    
    return results

# 修改後的 K-shift 函數（加入 capping rule）
def discrete_kshift_calculator(base_data, weekly_data, characteristic, resolution, ucl, lcl):
    """
    離散型數據的 K-shift 計算，加入 capping rule
    
    Capping rule: 如果當周點數<=10 且 當周P95/P50/P05沒有超過 baseline的P05和P95範圍外，就不HL
    """
    
    print("  discrete_kshift: 開始計算離散型 K-shift")
    
    # 先使用原本的 kshift_sigma_ratio_calculator 獲取結果
    kshift_results = kshift_sigma_ratio_calculator(
        base_data, weekly_data, characteristic, resolution, ucl, lcl
    )
    
    weekly_cnt = weekly_data['cnt']
    print(f"  discrete_kshift: 當周點數 = {weekly_cnt}")
    
    # 應用 capping rule
    if weekly_cnt <= 10:
        print("  discrete_kshift: 應用 capping rule (當周點數 <= 10)")
        
        # 計算基線和當周的百分位數
        try:
            base_percentiles = get_percentiles(base_data['values'])
            weekly_percentiles = get_percentiles(weekly_data['values'])
            
            # 檢查當周P95/P50/P05是否超過baseline的P05和P95範圍外
            base_p05 = base_percentiles.get('P05', np.nan)
            base_p95 = base_percentiles.get('P95', np.nan)
            
            weekly_p05 = weekly_percentiles.get('P05', np.nan)
            weekly_p50 = weekly_percentiles.get('P50', np.nan) 
            weekly_p95 = weekly_percentiles.get('P95', np.nan)
            
            # 如果當周的任何百分位數都沒有超出基線的P05-P95範圍，則設為NO_HIGHLIGHT
            if (not pd.isna(base_p05) and not pd.isna(base_p95) and 
                not pd.isna(weekly_p05) and not pd.isna(weekly_p50) and not pd.isna(weekly_p95)):
                
                if (base_p05 <= weekly_p05 <= base_p95 and 
                    base_p05 <= weekly_p50 <= base_p95 and 
                    base_p05 <= weekly_p95 <= base_p95):
                    
                    print("  discrete_kshift: Capping rule 觸發 - 當周百分位數未超出基線範圍，設為 NO_HIGHLIGHT")
                    kshift_results['P95_shift'] = 'NO_HIGHLIGHT'
                    kshift_results['P50_shift'] = 'NO_HIGHLIGHT'
                    kshift_results['P05_shift'] = 'NO_HIGHLIGHT'
                    
        except Exception as e:
            print(f"  discrete_kshift: Capping rule 檢查時發生錯誤: {e}")
    else:
        print("  discrete_kshift: 當周點數 > 10，不應用 capping rule")
    
    print(f"  discrete_kshift: 最終結果 - P95:{kshift_results.get('P95_shift')}, P50:{kshift_results.get('P50_shift')}, P05:{kshift_results.get('P05_shift')}")
    
    return kshift_results

# 離散型數據專用處理函數
def process_discrete_chart(raw_df, chart_info, weekly_start_date, weekly_end_date, initial_baseline_start_date, baseline_end_date):
    """
    離散型數據的專用處理流程，包含 record high low 判斷
    """
    group_name = chart_info.get('group_name', chart_info.get('GroupName', 'Unknown'))
    chart_name = chart_info.get('chart_name', chart_info.get('ChartName', 'Unknown'))
    
    print(f"  process_discrete_chart: 開始離散型專用處理 {group_name}/{chart_name}")
    
    try:
        # 篩選週數據和基線數據
        weekly_data = raw_df[(raw_df['point_time'] >= weekly_start_date) & (raw_df['point_time'] <= weekly_end_date)].copy()
        
        # 同樣的基線選擇邏輯
        baseline_data_one_year = raw_df[(raw_df['point_time'] >= initial_baseline_start_date) & (raw_df['point_time'] <= baseline_end_date)].copy()
        baseline_count_one_year = len(baseline_data_one_year)
        
        if baseline_count_one_year < 10:
            actual_baseline_start_date = baseline_end_date - pd.Timedelta(days=365 * 2)
            print(f"  process_discrete_chart: 基線數據不足，擴展至兩年: {actual_baseline_start_date} 至 {baseline_end_date}")
        else:
            actual_baseline_start_date = initial_baseline_start_date
            print(f"  process_discrete_chart: 使用一年基線期: {actual_baseline_start_date} 至 {baseline_end_date}")
            
        baseline_data = raw_df[(raw_df['point_time'] >= actual_baseline_start_date) & (raw_df['point_time'] <= baseline_end_date)].copy()
        
        print(f"  process_discrete_chart: 週數據點數={len(weekly_data)}, 基線數據點數={len(baseline_data)}")
        
        # 初始化結果
        result = {
            'group_name': group_name,
            'chart_name': chart_name,
            'weekly_start_date': weekly_start_date,
            'weekly_end_date': weekly_end_date,
            'baseline_start_date': actual_baseline_start_date,
            'baseline_end_date': baseline_end_date,
            'data_type': 'discrete',
        }
        
        if len(baseline_data) >= 5:  # 基線數據足夠
            print(f"  process_discrete_chart: 基線數據充足 ({len(baseline_data)} >= 5)，進行離散型 OOB 計算")
            
            # 準備數據字典
            base_data_dict = {
                'values': baseline_data['point_val'].values,
                'cnt': len(baseline_data),
                'mean': baseline_data['point_val'].mean(),
                'sigma': baseline_data['point_val'].std()
            }
            
            weekly_data_dict = {
                'values': weekly_data['point_val'].values,
                'cnt': len(weekly_data),
                'mean': weekly_data['point_val'].mean(),
                'sigma': weekly_data['point_val'].std()
            }
            
            # 離散型 OOB 計算
            print("  process_discrete_chart: 計算離散型 OOB...")
            discrete_oob_result = discrete_oob_calculator(base_data_dict, weekly_data_dict, chart_info)
            
            # Record High Low 計算
            print("  process_discrete_chart: 計算 record high low...")
            record_results = record_high_low_calculator(
                weekly_data['point_val'].values, 
                baseline_data['point_val'].values
            )
            
            # 更新結果
            result.update({
                'HL_P95_shift': discrete_oob_result.get('HL_P95_shift', 'NO_HIGHLIGHT'),
                'HL_P50_shift': discrete_oob_result.get('HL_P50_shift', 'NO_HIGHLIGHT'),
                'HL_P05_shift': discrete_oob_result.get('HL_P05_shift', 'NO_HIGHLIGHT'),
                'HL_sticking_shift': discrete_oob_result.get('HL_sticking_shift', 'NO_HIGHLIGHT'),
                'HL_trending': discrete_oob_result.get('HL_trending', 'NO_HIGHLIGHT'),
                'HL_high_OOC': discrete_oob_result.get('HL_high_OOC', 'NO_HIGHLIGHT'),
                'HL_category_LT_shift': discrete_oob_result.get('HL_category_LT_shift', 'NO_HIGHLIGHT'),
                'HL_record_high_low': record_results.get('highlight_status', 'NO_HIGHLIGHT'),
                'record_high': record_results.get('record_high', False),
                'record_low': record_results.get('record_low', False)
            })
            
            print(f" - process_discrete_chart: 離散型 OOB 計算完成")
            
        else:
            # 基線不足時設置所有 OOB 為 NO_HIGHLIGHT
            result.update({
                'HL_P95_shift': 'NO_HIGHLIGHT',
                'HL_P50_shift': 'NO_HIGHLIGHT',
                'HL_P05_shift': 'NO_HIGHLIGHT',
                'HL_sticking_shift': 'NO_HIGHLIGHT',
                'HL_trending': 'NO_HIGHLIGHT',
                'HL_high_OOC': 'NO_HIGHLIGHT',
                'HL_category_LT_shift': 'NO_HIGHLIGHT',
                'HL_record_high_low': 'NO_HIGHLIGHT',
                'record_high': False,
                'record_low': False
            })
            print(f" - process_discrete_chart: 基線數據不足，所有 OOB 設為 NO_HIGHLIGHT")

        print(f" - process_discrete_chart: 離散型處理完成 {group_name}/{chart_name}")
        return result

    except Exception as e:
        print(f" - process_discrete_chart: 處理錯誤 {group_name}/{chart_name}: {e}")
        traceback.print_exc()
        return None


def process_single_chart(chart_info, raw_df, initial_baseline_start_date, baseline_end_date, weekly_start_date, weekly_end_date):
    print("--- 進入外部 process_single_chart 函數 ---")
    print(f"  接收到的 raw_df shape: {raw_df.shape}")
    print(f"  週數據範圍: {weekly_start_date} 至 {weekly_end_date}")
    # 注意：這裡接收的是 initial_baseline_start_date (通常是往前一年)
    print(f"  初始基線數據範圍 (往前一年): {initial_baseline_start_date} 至 {baseline_end_date}")

    if raw_df is None or raw_df.empty:
        print("  raw_df 是空的或 None, 返回 None")
        return None

    try:
        print("  正在篩選週數據...")
        weekly_data = raw_df[(raw_df['point_time'] >= weekly_start_date) & (raw_df['point_time'] <= weekly_end_date)].copy() # Use copy()
        print(f"  篩選後 weekly_data shape: {weekly_data.shape}")

        if weekly_data.empty:
             print(f'未找到週數據, GroupName: {chart_info.get("group_name", "N/A")}, ChartName: {chart_info.get("chart_name", "N/A")}, 返回 None')
             return None

        # === 數據類型判斷和分流處理 ===
        print("  正在判斷數據類型...")
        data_type = determine_data_type(weekly_data['point_val'])
        
        if data_type == 'discrete':
            print("  數據類型為離散型，調用離散型專用處理流程")
            return process_discrete_chart(raw_df, chart_info, weekly_start_date, weekly_end_date, initial_baseline_start_date, baseline_end_date)
        else:
            print("  數據類型為連續型，繼續原有處理流程")

        # --- 基線數據範圍選擇邏輯開始 ---

        # 步驟 1: 使用初始的一年基線範圍過濾數據並計數
        print("  正在篩選初始一年基線數據...")
        baseline_data_one_year = raw_df[(raw_df['point_time'] >= initial_baseline_start_date) & (raw_df['point_time'] <= baseline_end_date)].copy() # Use copy()
        baseline_count_one_year = len(baseline_data_one_year)
        print(f"  初始一年基線數據點數量: {baseline_count_one_year}")

        # 步驟 2: 根據計數決定最終使用的基線開始日期
        if baseline_count_one_year < 10:
            # 如果少於 10 點，將基線期擴展到兩年
            actual_baseline_start_date = baseline_end_date - pd.Timedelta(days=365 * 2)
            print(f"  基線數據點數量 ({baseline_count_one_year}) < 10，將基線期擴展至兩年: {actual_baseline_start_date} 至 {baseline_end_date}")
        else:
            # 如果大於等於 10 點，使用一年的基線期
            actual_baseline_start_date = initial_baseline_start_date
            print(f"  基線數據點數量 ({baseline_count_one_year}) >= 10，使用一年基線期: {actual_baseline_start_date} 至 {baseline_end_date}")

        # 步驟 3: 使用最終確定的基線範圍過濾數據
        print("  正在篩選最終基線數據...")
        baseline_data = raw_df[(raw_df['point_time'] >= actual_baseline_start_date) & (raw_df['point_time'] <= baseline_end_date)].copy() # Use copy()
        print(f"  篩選後 baseline_data shape (使用 {len(baseline_data)} 點從 {actual_baseline_start_date} 至 {baseline_end_date}): {baseline_data.shape}")


        if baseline_data.empty:
             print(f'未找到基線數據 (在確定的範圍內), GroupName: {chart_info.get("group_name", "N/A")}, ChartName: {chart_info.get("chart_name", "N/A")}, 返回 None')
             return None

        # --- 基線數據範圍選擇邏輯結束 ---


        # 計算統計數據（週數據與基線數據）
        def calculate_statistics(data):
             # 新增檢查，避免對只有一個點的數據計算標準差產生 NaN (ddof=1 時)
             if data.shape[0] <= 1:
                  sigma = 0.0 if data.shape[0] == 1 else 0.0 # 單點或零點標準差視為 0
             else:
                  sigma = data['point_val'].std() # ddof=1 是 pandas 預設，計算樣本標準差

             # 如果 sigma 是 NaN (例如，所有值都相同，但數據點多於 1 且少於某個閾值，或計算出問題)
             if np.isnan(sigma):
                 print(f"  calculate_statistics 警告: 計算 sigma 得到 NaN. Data shape: {data.shape}")
                 sigma = 0.0 # 將無效的標準差視為 0

             return {
                 'values': data['point_val'].values,
                 'cnt': data.shape[0],
                 'mean': data['point_val'].mean(),
                 'sigma': sigma # 使用處理過的 sigma
                 }

        print("  正在計算週數據統計...")
        weekly_data_dict = calculate_statistics(weekly_data)
        print(f"  週數據統計結果 (部分): cnt={weekly_data_dict['cnt']}, mean={weekly_data_dict['mean']}, sigma={weekly_data_dict['sigma']}")


        # IMPORTANT: 這裡的 baseline_data_dict 現在是使用 *實際確定* 的基線範圍數據計算的
        print("  正在計算基線數據統計...")
        baseline_data_dict = calculate_statistics(baseline_data)
        print(f"  基線數據統計結果 (部分): cnt={baseline_data_dict['cnt']}, mean={baseline_data_dict['mean']}, sigma={baseline_data_dict['sigma']}")

        # 確保基線統計數據的標準差不會導致後續計算問題
        if baseline_data_dict['sigma'] == 0 or np.isnan(baseline_data_dict['sigma']):
             print("  警告: 基線標準差為零或無效，可能影響 K 值計算和需要標準差的其他指標。")
             # 您可以選擇在這裡返回 None，或讓後續函數自行處理 NaN/inf

        print("  正在呼叫 kshift_sigma_ratio_calculator...")
        # 傳入使用實際基線範圍計算出的 baseline_data_dict
        # kshift_sigma_ratio_calculator 需要處理 sigma=0 或其他分母為 0 的情況 (已在 safe_division 中處理)
        kshift_results = kshift_sigma_ratio_calculator(baseline_data_dict, weekly_data_dict, chart_info.get('Characteristics'), chart_info.get('Resolution'), chart_info.get('UCL'), chart_info.get('LCL')) # 使用 .get 防止 key 錯誤

        print(f"  kshift_sigma_ratio_calculator 返回: {kshift_results}")

        print("  正在呼叫 ooc_calculator...")
        # ooc_calculator 使用週數據計算 OOC 點數
        ooc_results = ooc_calculator(weekly_data, chart_info.get('UCL'), chart_info.get('LCL')) # 使用 .get 防止 key 錯誤
        print(f"  ooc_calculator 返回: {ooc_results}")

        print("  正在呼叫 review_ooc_results...")
        ooc_highlight = review_ooc_results(ooc_results[1], ooc_results[2]) # 注意 ooc_results[1] 是 ooc_cnt, ooc_results[2] 是 ooc_points
        print(f"  review_ooc_results 返回: {ooc_highlight}")

        print("  正在呼叫 sticking_rate_calculator...")

        sticking_rate_results = sticking_rate_calculator(baseline_data['point_val'], weekly_data['point_val'])
        print(f"  sticking_rate_calculator 返回: {sticking_rate_results}")

        print("  正在呼叫 trending...")
        # trending 也需要使用實際確定後的基線範圍
        trending_results = trending(raw_df, weekly_start_date, weekly_end_date, actual_baseline_start_date, baseline_end_date)
        print(f"  trending 返回: {trending_results}")

        # 判斷是否需要 highlight (任何一個子指標需要高亮，則總體高亮)
        highlight_status = 'HIGHLIGHT' if (
             kshift_results.get('P95_shift') == 'HIGHLIGHT' or
             kshift_results.get('P50_shift') == 'HIGHLIGHT' or
             kshift_results.get('P05_shift') == 'HIGHLIGHT' or
             sticking_rate_results.get('highlight_status') == 'HIGHLIGHT' or
             trending_results == 'HIGHLIGHT' or
             ooc_highlight == 'HIGHLIGHT' # 應該也要考慮 ooc_highlight
        ) else 'NO_HIGHLIGHT'
        print(f"  計算出的 highlight_status: {highlight_status}")

        result = {
            'data_cnt': ooc_results[0], # 週數據點數
            'ooc_cnt': ooc_results[1], # 週數據 OOC 點數
            'WE_Rule': '', # 這個欄位在 GUI 類的 build_result 中填充
            'OOB_Rule': '', # 這個欄位在 GUI 類的 build_result 中填充
            'data_type': 'continuous',  # 新增：標記為連續型數據
            'HL_P95_shift': kshift_results.get('P95_shift', 'N/A'), # 使用 get 並提供預設值，避免 key 錯誤
            'HL_P50_shift': kshift_results.get('P50_shift', 'N/A'),
            'HL_P05_shift': kshift_results.get('P05_shift', 'N/A'),
            'HL_sticking_shift': sticking_rate_results.get('highlight_status', 'N/A'),
            'HL_trending': trending_results, # trending_results 本身就是 HIGHLIGHT/NO_HIGHLIGHT
            'HL_high_OOC': ooc_highlight, # ooc_highlight 本身就是 HIGHLIGHT/NO_HIGHLIGHT
            'HL_record_high_low': 'NO_HIGHLIGHT',  # 連續型暫時設為 NO_HIGHLIGHT
            'HL_category_LT_shift': 'NO_HIGHLIGHT',  # 連續型不使用此功能
            'record_high': False,  # 新增 record_high 欄位
            'record_low': False,   # 新增 record_low 欄位
            'Material_no': chart_info.get('material_no', 'N/A'),
            'group_name': chart_info.get('group_name', 'N/A'),
            'chart_name': chart_info.get('chart_name', 'N/A'),
            'chart_ID': chart_info.get('ChartID', 'N/A'),
            'Characteristics': chart_info.get('Characteristics', 'N/A'),
            'USL': chart_info.get('USL', 'N/A'),
            'LSL': chart_info.get('LSL', 'N/A'),
            'UCL': chart_info.get('UCL', 'N/A'),
            'LCL': chart_info.get('LCL', 'N/A'),
            'Target': chart_info.get('Target', 'N/A'),
            'Resolution': chart_info.get('Resolution', 'N/A'),
            # 可以考慮添加 actual_baseline_start_date 到結果中，用於記錄實際使用的基線範圍
            # 'Actual_Baseline_Start': actual_baseline_start_date
        }
        print("--- 外部 process_single_chart 函數成功退出 ---")
        return result

    except Exception as e:
        # 在外部函數中捕獲異常並印出 traceback
        print(f'處理圖表時出錯 (外部函數) {chart_info.get("group_name", "N/A")} - {chart_info.get("chart_name", "N/A")}: {e}')
        traceback.print_exc()
        return None

def calculate_sigma(UCL, LCL, mean):
    sigma_upper = (UCL - mean) / 3
    sigma_lower = (mean - LCL) / 3
    return sigma_upper, sigma_lower    
def check_rules(raw_df, chart_info):
    mean = chart_info['Target']
    sigma_upper, sigma_lower = calculate_sigma(chart_info['UCL'], chart_info['LCL'], mean)
    UWL = mean + 2 * sigma_upper
    LWL = mean - 2 * sigma_lower
    characteristics = chart_info['Characteristics']

    rules = {
        "WE2": False,
        "WE3": False,
        "WE4": False,
        "WE6": False,
        "WE7": False,
        "WE8": False,
        "WE9": False,
        "WE10": False
    }

    rules["WE1"] = raw_df['point_val'].iloc[-1] > chart_info['UCL']
    rules["WE5"] = raw_df['point_val'].iloc[-1] < chart_info['LCL']

    if chart_info.get('WE2', 'N') == 'Y':
        rules["WE2"] = (raw_df['point_val'].tail(3) > UWL).sum() >= 2 if characteristics not in ['Bigger', 'Smaller'] else False
    if chart_info.get('WE3', 'N') == 'Y':
        rules["WE3"] = (raw_df['point_val'].tail(5) > 0.5 * (mean + UWL)).sum() >= 4
    if chart_info.get('WE4', 'N') == 'Y':
        rules["WE4"] = (raw_df['point_val'].tail(8) > mean).all()
    if chart_info.get('WE6', 'N') == 'Y':
        rules["WE6"] = (raw_df['point_val'].tail(3) < LWL).sum() >= 2 if characteristics not in ['Bigger', 'Smaller'] else False
    if chart_info.get('WE7', 'N') == 'Y':
        rules["WE7"] = (raw_df['point_val'].tail(5) < 0.5 * (mean + LWL)).sum() >= 4
    if chart_info.get('WE8', 'N') == 'Y':
        rules["WE8"] = (raw_df['point_val'].tail(8) < mean).all()
    if chart_info.get('WE9', 'N') == 'Y':
        # 取得最後 15 筆資料
        tail_points = raw_df['point_val'].tail(15)
        
        # 如果所有資料點報定值（唯一值數量為 1），則直接返回 False
        if tail_points.nunique() == 1:  # 檢查唯一值數量是否為 1
            rules["WE9"] = False
        else:
            # 正常執行條件判斷
            condition_result = (tail_points > (mean - sigma_lower)) & \
                            (tail_points < (mean + sigma_upper))
            rules["WE9"] = condition_result.all()
    if chart_info.get('WE10', 'N') == 'Y':                   
        rules["WE10"] = ((raw_df['point_val'].tail(8) < (mean - sigma_lower) ) | 
                        (raw_df['point_val'].tail(8) > (mean + sigma_upper) )).all() if characteristics not in ['Bigger', 'Smaller'] else False
    return rules
def calculate_cpk(raw_df, chart_info):
    mean = raw_df['point_val'].mean()
    std = raw_df['point_val'].std()
    characteristic = chart_info['Characteristics']
    usl = chart_info.get('USL', None)
    lsl = chart_info.get('LSL', None)

    cpk = None

    if std > 0:
        if characteristic == 'Nominal':
            if usl is not None and lsl is not None:
                cpu = (usl - mean) / (3 * std)
                cpl = (mean - lsl) / (3 * std)
                cpk = min(cpu, cpl)
        elif characteristic == 'Smaller':
            if usl is not None:
                cpk = (usl - mean) / (3 * std)
        elif characteristic == 'Bigger':
            if lsl is not None:
                cpk = (mean - lsl) / (3 * std)

    if cpk is not None:
        cpk = round(cpk, 3)  # 統一四捨五入到小數第三位

    return {'Cpk': cpk}
def plot_spc_chart(raw_df, chart_info, weekly_start_date, weekly_end_date):
    plt.figure(figsize=(14, 6))

    group_name = chart_info['group_name']
    display_group_name = "" if group_name == "Default" else f"Group: [{group_name}]"
    title = (f"{display_group_name}[{chart_info['chart_name']}][{chart_info['Characteristics']}]\n"
             f"UCL: [{chart_info['UCL']}] | Target: [{chart_info['Target']}] | LCL: [{chart_info['LCL']}]")
    plt.title(title, loc='left', fontsize=12)

    if len(raw_df) > 300:
        raw_df = raw_df.tail(len(raw_df))

    points_num = len(raw_df)
    x_values = np.arange(points_num)

    plt.hlines(chart_info['UCL'], -0.8, points_num + 2, colors='#E83F6F', linestyles='--', linewidth=1)
    plt.hlines(chart_info['Target'], -0.8, points_num + 2, colors='#087E8B', linestyles='--', linewidth=1)
    plt.hlines(chart_info['LCL'], -0.8, points_num + 2, colors='#E83F6F', linestyles='--', linewidth=1)

    plt.text(x=points_num + 2, y=chart_info['UCL'], s='UCL', va='center', ha='left', fontsize=10, color='#E83F6F')
    plt.text(x=points_num + 2, y=chart_info['Target'], s='Target', va='center', ha='left', fontsize=10, color='#087E8B')
    plt.text(x=points_num + 2, y=chart_info['LCL'], s='LCL', va='center', ha='left', fontsize=10, color='#E83F6F')

    raw_df['point_time'] = pd.to_datetime(raw_df['point_time'])
    weekly_start_date = pd.to_datetime(weekly_start_date)
    weekly_end_date = pd.to_datetime(weekly_end_date)

    start_index = raw_df[raw_df['point_time'] >= weekly_start_date].index.min() - raw_df.index.min()
    end_index = raw_df[raw_df['point_time'] <= weekly_end_date].index.max() - raw_df.index.min()

    plt.plot(x_values, raw_df['point_val'], color='#5863F8', marker='o', linestyle='-')

    violated_rules = {rule: False for rule in chart_info.get('rule_list', [])}

    for i in range(start_index, end_index + 1):
        weekly_data_subset = raw_df.iloc[:i+1].tail(15)
        if not weekly_data_subset.empty:
            rules = check_rules(weekly_data_subset.copy(), chart_info)
            for rule, violated in rules.items():
                if violated:
                    violated_rules[rule] = True
                    plt.plot(x_values[i], raw_df['point_val'].iloc[i], 'ro', markersize=10)

    interval = max(1, len(raw_df) // 30)
    plt.xticks(x_values[::interval], raw_df['point_time'].dt.strftime("%Y-%m-%d")[::interval], rotation=90)

    plt.axvspan(start_index, end_index, color='#E83F6F', alpha=0.1, label='Weekly Data')
    baseline_end_index = start_index - 1
    plt.axvspan(-1, baseline_end_index + 1, color='#3772FF', alpha=0.1, label='Baseline Data')
    plt.xlim([x_values[0] - 1, None])
    plt.legend()

    ax = plt.gca()
    ax.spines['right'].set_visible(False)
    ax.spines['top'].set_visible(False)

    plt.tight_layout()

    output_path = 'output'
    if not os.path.exists(output_path):
        os.makedirs(output_path)

    safe_group_name = "" if group_name == "Default" else group_name
    image_path = f'{output_path}/SPC_{safe_group_name}_{chart_info["chart_name"]}.png'
    plt.savefig(image_path, bbox_inches='tight')
    plt.close()

    return image_path, violated_rules

    

def plot_weekly_spc_chart(raw_df, chart_info, weekly_start_date, weekly_end_date):
    raw_df_weekly = raw_df[(raw_df['point_time'] >= pd.to_datetime(weekly_start_date)) &
                           (raw_df['point_time'] <= pd.to_datetime(weekly_end_date))].copy()

    plt.figure(figsize=(14, 6))

    group_name = chart_info['group_name']
    display_group_name = "" if group_name == "Default" else f"Group: [{group_name}]"
    title = (f"{display_group_name}[{chart_info['chart_name']}][{chart_info['Characteristics']}]\n"
             f"UCL: [{chart_info['UCL']}] | Target: [{chart_info['Target']}] | LCL: [{chart_info['LCL']}]")
    plt.title(title, loc='left', fontsize=12)

    if len(raw_df_weekly) > 300:
        raw_df_weekly = raw_df_weekly.tail(len(raw_df_weekly))

    points_num = len(raw_df_weekly)
    x_values = np.arange(points_num)

    plt.hlines(chart_info['UCL'], -0.8, points_num + 2, colors='#E83F6F', linestyles='--', linewidth=1)
    plt.hlines(chart_info['Target'], -0.8, points_num + 2, colors='#087E8B', linestyles='--', linewidth=1)
    plt.hlines(chart_info['LCL'], -0.8, points_num + 2, colors='#E83F6F', linestyles='--', linewidth=1)

    plt.text(x=points_num + 2, y=chart_info['UCL'], s='UCL', va='center', ha='left', fontsize=10, color='#E83F6F')
    plt.text(x=points_num + 2, y=chart_info['Target'], s='Target', va='center', ha='left', fontsize=10, color='#087E8B')
    plt.text(x=points_num + 2, y=chart_info['LCL'], s='LCL', va='center', ha='left', fontsize=10, color='#E83F6F')

    plt.plot(x_values, raw_df_weekly['point_val'], color='#5863F8', marker='o', linestyle='-')

    for i in range(len(raw_df_weekly)):
        weekly_data_subset = raw_df_weekly.iloc[:i+1].tail(15)
        if not weekly_data_subset.empty:
            rules = check_rules(weekly_data_subset.copy(), chart_info)
            if any(rules.values()):
                plt.plot(x_values[i], raw_df_weekly['point_val'].iloc[i], 'ro', markersize=10)

    interval = max(1, len(raw_df_weekly) // 30)
    plt.xticks(x_values[::interval], raw_df_weekly['point_time'].dt.strftime("%Y-%m-%d")[::interval], rotation=90)

    plt.axvspan(0, points_num - 1, color='#E83F6F', alpha=0.1, label='Weekly Data')
    plt.xlim([x_values[0] - 1, None])
    plt.legend()

    ax = plt.gca()
    ax.spines['right'].set_visible(False)
    ax.spines['top'].set_visible(False)

    plt.tight_layout()

    output_path = 'output'
    if not os.path.exists(output_path):
        os.makedirs(output_path)

    safe_group_name = "" if group_name == "Default" else group_name
    image_path = f'{output_path}/Weekly_SPC_{safe_group_name}_{chart_info["chart_name"]}.png'
    plt.savefig(image_path, bbox_inches='tight')
    plt.close()

    return image_path

def save_oob_results_to_excel(results_df):
    """
    專門為 OOB 結果匯出的 Excel 函數
    包含主要結果、高亮異常摘要和統計分析
    """
    output = BytesIO()
    
    try:
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # 1. 主要結果工作表
            results_df.to_excel(writer, sheet_name='OOB_Results', index=False)
            
            # 2. 高亮異常摘要
            oob_columns = [col for col in results_df.columns if col.startswith('HL_')]
            if oob_columns:
                # 篩選有異常的記錄
                highlight_mask = pd.Series(False, index=results_df.index)
                for col in oob_columns:
                    if col in results_df.columns:
                        highlight_mask |= (results_df[col] == 'HIGHLIGHT')
                
                highlighted_results = results_df[highlight_mask]
                if not highlighted_results.empty:
                    highlighted_results.to_excel(writer, sheet_name='異常警示', index=False)
                
                # 3. 統計摘要
                summary_data = []
                for col in oob_columns:
                    if col in results_df.columns:
                        highlight_count = (results_df[col] == 'HIGHLIGHT').sum()
                        total_count = results_df[col].notna().sum()
                        oob_type = col.replace('HL_', '').replace('_', ' ')
                        summary_data.append({
                            'OOB類型': oob_type,
                            '異常數量': highlight_count,
                            '總數量': total_count,
                            '異常比例': f"{highlight_count/total_count*100:.1f}%" if total_count > 0 else "0%"
                        })
                
                if summary_data:
                    summary_df = pd.DataFrame(summary_data)
                    summary_df.to_excel(writer, sheet_name='統計摘要', index=False)
            
            # 4. 數據類型摘要
            if 'data_type' in results_df.columns:
                type_summary = results_df['data_type'].value_counts().reset_index()
                type_summary.columns = ['數據類型', '數量']
                type_summary.to_excel(writer, sheet_name='數據類型統計', index=False)
        
        output.seek(0)
        return output
        
    except Exception as e:
        st.error(f"生成 OOB Excel 報告時發生錯誤: {str(e)}")
        return None

def save_results_to_excel(results_df, scale_factor=0.3):
    # 先清理數據，替換 NaN 和 Inf 值
    results_df = results_df.copy()
    results_df['group_name'] = results_df['group_name'].replace("Default", "")  # 替換 Default 為空白
    
    # 處理 NaN 和 Inf 值
    results_df = results_df.replace([np.nan, np.inf, -np.inf], 'N/A')

    output = BytesIO()
    # 添加 nan_inf_to_errors 選項來處理 NaN/Inf 值
    workbook = xlsxwriter.Workbook(output, {
        'in_memory': True,
        'nan_inf_to_errors': True
    })
    worksheet = workbook.add_worksheet()

    cell_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'font_name': 'Arial', 'font_size': 10})
    header_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'font_name': 'Arial', 'font_size': 12, 'bold': True})

    col_widths = {}
    max_image_height = 0
    image_column_width = 0

    # 寫入標題列
    for col_idx, header in enumerate(results_df.columns):
        worksheet.write(0, col_idx + 2, header, header_format)
        col_widths[col_idx + 2] = max(len(header), col_widths.get(col_idx + 2, 0))

    for row_idx, row in enumerate(results_df.itertuples(index=False), start=1):
        img_path = row.chart_path if hasattr(row, 'chart_path') else None
        weekly_path = row.weekly_chart_path if hasattr(row, 'weekly_chart_path') else None

        x_offset = 0
        y_offset = 10
        options = {
            'x_scale': scale_factor,
            'y_scale': scale_factor,
            'x_offset': x_offset,
            'y_offset': y_offset,
            'object_position': 1
        }

        # 插入圖片
        if img_path and isinstance(img_path, str) and os.path.exists(img_path):
            try:
                worksheet.insert_image(row_idx, 0, img_path, options)
                img = Image.open(img_path)
                w, h = img.size
                scaled_w, scaled_h = w * scale_factor, h * scale_factor
                max_image_height = max(max_image_height, scaled_h)
                image_column_width = max(image_column_width, scaled_w)
            except Exception as e:
                print(f"插入圖片失敗 {img_path}: {e}")

        if weekly_path and isinstance(weekly_path, str) and os.path.exists(weekly_path):
            try:
                worksheet.insert_image(row_idx, 1, weekly_path, options)
            except Exception as e:
                print(f"插入週圖片失敗 {weekly_path}: {e}")

        # 插入資料欄位
        for col_idx, value in enumerate(row, start=1):
            try:
                # 確保值不是 NaN 或 Inf
                if pd.isna(value) or value in [np.inf, -np.inf]:
                    value = 'N/A'
                worksheet.write(row_idx, col_idx + 1, value, cell_format)
                col_widths[col_idx + 1] = max(col_widths.get(col_idx + 1, 0), len(str(value)))
            except Exception as e:
                print(f"寫入資料失敗 (行{row_idx}, 列{col_idx+1}, 值{value}): {e}")
                worksheet.write(row_idx, col_idx + 1, 'Error', cell_format)

    # 設定欄寬與列高
    try:
        worksheet.set_column(0, 1, image_column_width / 7)
        for col_idx, width in col_widths.items():
            worksheet.set_column(col_idx, col_idx, width + 5)
        for row_idx in range(1, len(results_df) + 1):
            worksheet.set_row(row_idx, max_image_height)
    except Exception as e:
        print(f"設定欄寬列高失敗: {e}")

    workbook.close()
    output.seek(0)
    return output

    # 常數定義
HEADERS = ["Total Chart", "Weekly Chart", "Chart Info."]
OOB_KEYS = ['HL_P95_shift', 'HL_P50_shift', 'HL_P05_shift', 'HL_sticking_shift', 'HL_trending', 'HL_high_OOC', 'HL_record_high_low', 'HL_category_LT_shift']
def analyze_chart(execution_time, raw_df, chart_info):
    group_name = str(chart_info.get('group_name', chart_info.get('GroupName', 'Unknown')))
    chart_name = str(chart_info.get('chart_name', chart_info.get('ChartName', 'Unknown')))

    if 'point_time' not in raw_df.columns or not pd.api.types.is_datetime64_any_dtype(raw_df['point_time']):
        # 無法分析
        return None

    latest_raw_data_time = raw_df['point_time'].max()

    if execution_time is None or pd.isna(execution_time):
        weekly_end_date = latest_raw_data_time
    else:
        weekly_end_date = execution_time

    if pd.isna(weekly_end_date):
        return None

    weekly_start_date = weekly_end_date - pd.Timedelta(days=6)
    baseline_end_date = weekly_start_date - pd.Timedelta(seconds=1)
    baseline_start_date = baseline_end_date - pd.Timedelta(days=365)

    try:
        result = process_single_chart(chart_info.copy(), raw_df, baseline_start_date, baseline_end_date, weekly_start_date, weekly_end_date)
        if result is None or not isinstance(result, dict):
            return None

        image_path, violated_rules = plot_spc_chart(raw_df, chart_info, weekly_start_date, weekly_end_date)

        weekly_data = raw_df[(raw_df['point_time'] >= weekly_start_date) & (raw_df['point_time'] <= weekly_end_date)].copy()

        cpk_result = calculate_cpk(weekly_data, chart_info)
        if cpk_result and 'Cpk' in cpk_result:
            result['Cpk'] = cpk_result['Cpk']
        else:
            result['Cpk'] = np.nan

        weekly_image_path = plot_weekly_spc_chart(raw_df, chart_info, weekly_start_date, weekly_end_date)

        result['violated_rules'] = violated_rules if violated_rules is not None else {}

        build_result(result, image_path, weekly_image_path)

        return result

    except Exception:
        traceback.print_exc()
        return None


def build_result(result, image_path, weekly_image_path):
    violated_rules = result.get('violated_rules', {})
    we_true_keys = [k for k, v in violated_rules.items() if v]
    result['WE_Rule'] = ', '.join(we_true_keys) if we_true_keys else 'N/A'

    oob_true_keys = [k for k in OOB_KEYS if result.get(k) == 'HIGHLIGHT']
    result['OOB_Rule'] = ', '.join(oob_true_keys) if oob_true_keys else 'N/A'

    for key in OOB_KEYS:
        result.pop(key, None)
    result.pop('violated_rules', None)  # 移除原始字典

    result['chart_path'] = image_path
    result['weekly_chart_path'] = weekly_image_path

    if 'Cpk' not in result:
        result['Cpk'] = np.nan


def process_all_charts(raw_data_dir, all_charts_info_df, execution_time=None):
    results = []
    skipped_charts_count = 0

    for idx, chart_info in all_charts_info_df.iterrows():
        group_name = str(chart_info.get('GroupName', chart_info.get('group_name', 'Unknown')))
        chart_name = str(chart_info.get('ChartName', chart_info.get('chart_name', 'Unknown')))

        try:
            filepath = find_matching_file(raw_data_dir, group_name, chart_name)
            if not filepath or not os.path.exists(filepath):
                skipped_charts_count += 1
                continue

            raw_df = pd.read_csv(filepath)
            if 'point_time' in raw_df.columns:
                raw_df['point_time'] = pd.to_datetime(raw_df['point_time'], errors='coerce')
                raw_df.dropna(subset=['point_time'], inplace=True)

            is_successful, processed_df, updated_chart_info = preprocess_data(chart_info.copy(), raw_df.copy())
            if not is_successful or processed_df is None or processed_df.empty:
                skipped_charts_count += 1
                continue

            result = analyze_chart(execution_time, processed_df, updated_chart_info)
            if result is None:
                skipped_charts_count += 1
                continue

            results.append(result)

        except Exception:
            skipped_charts_count += 1
            traceback.print_exc()

    return results, skipped_charts_count


def save_results(results):
    results_df = pd.DataFrame(results)

    expected_cols = ['data_cnt', 'ooc_cnt', 'WE_Rule', 'OOB_Rule', 'data_type', 'Material_no',
                     'group_name', 'chart_name', 'chart_ID', 'Characteristics',
                     'USL', 'LSL', 'UCL', 'LCL', 'Target', 'Cpk', 'Resolution',
                     'chart_path', 'weekly_chart_path']

    for col in expected_cols:
        if col not in results_df.columns:
            results_df[col] = np.nan

    cols_to_order = [col for col in expected_cols if col in results_df.columns]
    results_df = results_df[cols_to_order]

    results_df = results_df.replace([np.nan, np.inf, -np.inf], 'N/A')

    try:
        return save_results_to_excel(results_df)
    except Exception:
        traceback.print_exc()
        raise
st.set_page_config(page_title="SPC 圖表處理系統", layout="wide")

# 初始化 session_state
if 'results' not in st.session_state:
    st.session_state['results'] = None
if 'skipped_count' not in st.session_state:
    st.session_state['skipped_count'] = 0

# ==================== Tool Matching 模組函數 ====================

def get_k_value(n):
    """根據樣本數量 n 返回 K 值"""
    if n <= 4:  # 樣本數量太少，不進行比較
        return "不比較"  # 返回特殊標記，表示不進行比較
    elif 5 <= n <= 10:
        return 1.73
    elif 11 <= n <= 120:
        return 1.414
    else:
        return 1.15

def calculate_mean_index(mean1, mean2, min_sigma, characteristic):
    """計算 mean matching index，考慮方向性"""
    if min_sigma <= 0:
        return float('inf')
    
    if characteristic == 'up':  # Bigger is better
        return (mean2 - mean1) / min_sigma
    elif characteristic == 'down':  # Smaller is better
        return (mean1 - mean2) / min_sigma
    else:  # Nominal
        return abs(mean1 - mean2) / min_sigma

def analyze_two_groups(group_stats, gname, cname, characteristic, results):
    """分析兩台設備的匹配情況"""
    row1 = group_stats.iloc[0]
    row2 = group_stats.iloc[1]

    group1 = row1["matching_group"]
    group2 = row2["matching_group"]
    mean1, std1, n1 = row1["mean"], row1["std"], row1["count"]
    mean2, std2, n2 = row2["mean"], row2["std"], row2["count"]

    min_sigma = min(std1, std2)

    if n1 < 5 or n2 < 5:
        results.append([
            gname, cname, group1, 'group_all',
            '資料不足', '資料不足',
            get_k_value(n1), mean1, std1,
            mean2, min_sigma, n1
        ])
        results.append([
            gname, cname, group2, 'group_all',
            '資料不足', '資料不足',
            get_k_value(n2), mean2, std2,
            mean1, min_sigma, n2
        ])
        return

    k1 = get_k_value(n1)
    k2 = get_k_value(n2)

    # 分母為零時，判斷 mean 是否全相等
    if min_sigma > 0:
        mean_index_1 = abs(mean1 - mean2) / min_sigma
        sigma_index_1 = std1 / min_sigma
    else:
        all_means = [mean1, mean2]
        if len(set([round(m, 8) for m in all_means])) == 1:
            mean_index_1 = 0
            sigma_index_1 = 0
        else:
            mean_index_1 = float('inf')
            sigma_index_1 = float('inf')

    if k1 == "不比較":
        results.append([
            gname, cname, group1, 'group_all',
            '資料不足', '資料不足',
            '不比較', round(mean1, 2), round(std1, 2),
            round(mean2, 2), round(min_sigma, 2), n1
        ])
    else:
        results.append([
            gname, cname, group1, 'group_all',
            round(mean_index_1, 2), round(sigma_index_1, 2),
            round(k1, 2), round(mean1, 2), round(std1, 2),
            round(mean2, 2), round(min_sigma, 2), n1
        ])

    # 第二組
    if min_sigma > 0:
        mean_index_2 = abs(mean2 - mean1) / min_sigma
        sigma_index_2 = std2 / min_sigma
    else:
        all_means = [mean1, mean2]
        if len(set([round(m, 8) for m in all_means])) == 1:
            mean_index_2 = 0
            sigma_index_2 = 0
        else:
            mean_index_2 = float('inf')
            sigma_index_2 = float('inf')

    if k2 == "不比較":
        results.append([
            gname, cname, group2, 'group_all',
            '資料不足', '資料不足',
            '不比較', round(mean2, 2), round(std2, 2),
            round(mean1, 2), round(min_sigma, 2), n2
        ])
    else:
        results.append([
            gname, cname, group2, 'group_all',
            round(mean_index_2, 2), round(sigma_index_2, 2),
            round(k2, 2), round(mean2, 2), round(std2, 2),
            round(mean1, 2), round(min_sigma, 2), n2
        ])

def analyze_multiple_groups(subdf, group_stats, gname, cname, characteristic, results):
    """分析多台設備的匹配情況"""
    # 只納入樣本數 >= 5 的 group 計算 median
    valid_stats = group_stats[group_stats['count'] >= 5]
    if valid_stats.shape[0] <= 1:
        # 只有一個有效群組，全部標記資料不足
        for i, row in group_stats.iterrows():
            group = row["matching_group"]
            mean = row["mean"]
            std = row["std"]
            n = row["count"]
            results.append([
                gname, cname, group, "group_all",
                '資料不足', '資料不足', 
                get_k_value(n), mean, std, 
                '-', '-', n
            ])
        return

    mean_median = valid_stats['mean'].median() if not valid_stats.empty else 0
    median_sigma = valid_stats['std'].median() if not valid_stats.empty else 0

    for i, row in group_stats.iterrows():
        group = row["matching_group"]
        mean = row["mean"]
        std = row["std"]
        n = row["count"]

        # 計算 mean matching index（考慮方向性）
        if n < 5:  # 樣本數不足5個，不進行比較
            results.append([
                gname, cname, group, "group_all",
                '資料不足', '資料不足', 
                get_k_value(n), mean, std, 
                mean_median, median_sigma, n
            ])
            continue

        if median_sigma > 0:
            if characteristic == 'up':
                mean_index = (mean_median - mean) / median_sigma
            elif characteristic == 'down':
                mean_index = (mean - mean_median) / median_sigma
            else:
                mean_index = abs(mean - mean_median) / median_sigma
            sigma_index = std / median_sigma
        else:
            # 分母為零時，判斷所有 mean 是否相等
            all_means = group_stats['mean'].tolist() if not group_stats.empty else [mean]
            if len(set([round(m, 8) for m in all_means])) == 1:
                mean_index = 0
                sigma_index = 0
            else:
                mean_index = float('inf')
                sigma_index = float('inf')

        K = get_k_value(n)

        # 檢查 K 值是否為字串 "不比較"
        if K == "不比較":
            # 樣本數不足，使用 "資料不足" 標記
            results.append([
                gname, cname, group, "group_all",
                '資料不足', '資料不足', 
                '不比較', round(mean, 2), round(std, 2), 
                round(mean_median, 2), round(median_sigma, 2), n
            ])
        else:
            # 正常比較情況
            results.append([
                gname, cname, group, "group_all",
                round(mean_index, 2), round(sigma_index, 2), 
                round(K, 2), round(mean, 2), round(std, 2), 
                round(mean_median, 2), round(median_sigma, 2), n
            ])

def analyze_multiple_groups_time(mean_df, sigma_df, group_stats, gname, cname, characteristic, results):
    """
    多組分析（mean/std/count 來自一個月 window，median(sigma) 來自半年 window）
    """
    # 只納入樣本數 >= 5 的 group 計算 median
    valid_mean_df = mean_df.groupby("matching_group").filter(lambda x: len(x) >= 5)
    sigma_by_group = sigma_df.groupby("matching_group")["point_val"].std()
    valid_groups = group_stats[group_stats['count'] >= 5]['matching_group']
    valid_sigma = sigma_by_group[valid_groups] if not valid_groups.empty else pd.Series(dtype=float)
    
    # 防呆：如果有效 group 只有一個，全部標記資料不足
    if len(valid_groups) <= 1:
        for i, row in group_stats.iterrows():
            group = row["matching_group"]
            mean = row["mean"]
            std = row["std"]
            n = row["count"]
            results.append([
                gname, cname, group, "group_all",
                '資料不足', '資料不足', 
                get_k_value(n), mean, std, 
                '-', '-', n
            ])
        return
    
    mean_median = valid_mean_df["point_val"].median() if not valid_mean_df.empty else 0
    median_sigma = valid_sigma.median() if not valid_sigma.empty else 0
    
    for i, row in group_stats.iterrows():
        group = row["matching_group"]
        mean = row["mean"]
        std = row["std"]  # 這是來自 mean_df（一個月 window）
        n = row["count"]
        
        if n < 5:
            results.append([
                gname, cname, group, "group_all",
                '資料不足', '資料不足', 
                get_k_value(n), mean, std, 
                mean_median, median_sigma, n
            ])
            continue
            
        if median_sigma > 0:
            if characteristic == 'up':
                mean_index = (mean_median - mean) / median_sigma
            elif characteristic == 'down':
                mean_index = (mean - mean_median) / median_sigma
            else:
                mean_index = abs(mean - mean_median) / median_sigma
            sigma_index = std / median_sigma
        else:
            # 分母為零時，判斷所有 mean 是否相等
            all_means = group_stats['mean'].tolist() if not group_stats.empty else [mean]
            if len(set([round(m, 8) for m in all_means])) == 1:
                mean_index = 0
                sigma_index = 0
            else:
                mean_index = float('inf')
                sigma_index = float('inf')
                
        K = get_k_value(n)
        if K == "不比較":
            results.append([
                gname, cname, group, "group_all",
                '資料不足', '資料不足', 
                '不比較', round(mean, 2), round(std, 2), 
                round(mean_median, 2), round(median_sigma, 2), n
            ])
        else:
            results.append([
                gname, cname, group, "group_all",
                round(mean_index, 2), round(sigma_index, 2), 
                round(K, 2), round(mean, 2), round(std, 2), 
                round(mean_median, 2), round(median_sigma, 2), n
            ])

def generate_temp_csv():
    """生成範例 CSV 檔案"""
    data = {
        "GroupName": ["GroupA", "GroupA", "GroupB", "GroupB", "GroupA", "GroupA"],
        "ChartName": ["X", "X", "Y", "Y", "X", "X"],
        "point_time": ["2023/5/15 14:39", "2023/5/16 10:20", "2023/5/15 09:30", "2023/5/16 15:45", "2023/5/17 11:10", "2023/5/18 08:55"],
        "matching_group": ["A", "B", "A", "B", "A", "B"],
        "point_val": [99.88135943, 100.12345678, 98.76543210, 99.45678901, 100.02345678, 99.98765432],
        "characteristic": ["Nominal", "Nominal", "up", "up", "Nominal", "Nominal"]
    }
    df = pd.DataFrame(data)
    return df

def create_tool_matching_charts(grouped_data):
    """創建 SPC 圖和盒鬚圖，將 figure 物件保存在 chart_figures 中，不在 UI 上顯示。"""
    try:
        # 這些導入是必要的，因為 Matplotlib 在子線程或不同上下文中可能需要重新導入
        import matplotlib.pyplot as plt
        from matplotlib import cm
        import numpy as np
    except ImportError:
        print("[ERROR] Matplotlib is not installed.")
        return {}

    # 保存圖表與分組鍵的對應關係，用於後續的彈出視窗和 Excel 匯出
    chart_figures = {}
    
    # 為每個 (GroupName, ChartName) 組合創建圖表
    for (gname, cname), subdf in grouped_data:
        # 依 matching_group 字母順序排序
        unique_groups = sorted(subdf["matching_group"].unique(), key=lambda x: str(x))
        labels = [str(mg) for mg in unique_groups]

        # 檢查是否有數據可供繪圖
        if subdf.empty or not any(len(grp["point_val"]) > 0 for _, grp in subdf.groupby("matching_group")):
            print(f"[WARNING] Skipping chart creation for {gname} - {cname} due to empty data.")
            chart_figures[(gname, cname)] = {'scatter': None, 'box': None}
            continue

        # 依排序後 unique_groups 組裝 box_data，確保顏色/label/資料一致
        box_data = [subdf[subdf["matching_group"] == mg]["point_val"].values for mg in unique_groups]
        group_stats = subdf.groupby("matching_group")["point_val"].agg(['mean', 'std', 'count'])

        # 為不同的組設置顏色
        colors = cm.tab10(np.linspace(0, 1, len(unique_groups)))

        # 1. 創建 SPC 風格的圖表
        scatter_fig, scatter_ax = plt.subplots(figsize=(7, 4.5)) # 調整尺寸為較小的長方形
        
        # 計算整體統計量用於控制線
        all_values = subdf["point_val"].values
        # overall_mean = np.mean(all_values)
        # overall_std = np.std(all_values)
        

        # 為每個群組繪製數據點，按時間順序連線
        x_position = 0
        for i, mg in enumerate(unique_groups):
            group_data = subdf[subdf["matching_group"] == mg].sort_values("point_time")
            if not group_data.empty:
                # 為每個群組創建連續的x位置
                x_vals = np.arange(x_position, x_position + len(group_data))
                y_vals = group_data["point_val"].values
                
                # 繪製數據點
                scatter_ax.scatter(x_vals, y_vals, color=colors[i], alpha=0.8, s=40, label=f'{mg}', zorder=3)
                
                # 連接同組內的點
                scatter_ax.plot(x_vals, y_vals, color=colors[i], alpha=0.5, linewidth=1, zorder=2)
                
                # 在群組間添加分隔線
                if i < len(unique_groups) - 1:  # 不在最後一組後面加線
                    separator_x = x_position + len(group_data) - 0.5
                    scatter_ax.axvline(x=separator_x, color='gray', linestyle='-', alpha=0.3, zorder=1)
                
                x_position += len(group_data)
        
        # 設置圖表樣式
        scatter_ax.set_title(f"SPC Chart: {gname} - {cname}", fontsize=10)
        scatter_ax.set_xlabel("Sample Sequence (Grouped by Matching Group)")
        scatter_ax.set_ylabel("Point Value")
        scatter_ax.grid(True, linestyle='--', alpha=0.3, zorder=0)
        
        # 添加群組標籤在x軸上
        if unique_groups:
            group_positions = []
            x_pos = 0
            for mg in unique_groups:
                group_size = len(subdf[subdf["matching_group"] == mg])
                group_positions.append(x_pos + group_size/2 - 0.5)
                x_pos += group_size
            
            # 設置x軸刻度和標籤
            scatter_ax.set_xticks(group_positions)
            scatter_ax.set_xticklabels(labels, rotation=0, ha='center')
            
            # 添加次要刻度顯示樣本序號
            scatter_ax.tick_params(axis='x', which='minor', bottom=True, top=False)
        
        # 調整圖例位置
        scatter_ax.legend(loc='upper left', bbox_to_anchor=(1.02, 1), fontsize='small')
        scatter_fig.tight_layout()

        # 2. 創建盒鬚圖
        box_fig, box_ax = plt.subplots(figsize=(7, 4.5)) # 調整尺寸為較小的長方形
        if box_data:
            bp = box_ax.boxplot(box_data, labels=labels, patch_artist=True, widths=0.6)
            for patch, color in zip(bp['boxes'], colors):
                patch.set_facecolor(color)

            # legend 也照 unique_groups 順序
            legend_labels = [
                f"{label}: μ={group_stats.loc[mg, 'mean']:.2f}, σ={group_stats.loc[mg, 'std']:.2f}, n={int(group_stats.loc[mg, 'count'])}"
                for label, mg in zip(labels, unique_groups)
            ]
            box_ax.legend([bp["boxes"][i] for i in range(len(labels))], legend_labels, loc='upper left', bbox_to_anchor=(1.02, 1), fontsize='small')

        box_ax.set_title(f"Boxplot: {gname} - {cname}", fontsize=10)
        box_ax.set_xlabel("Matching Group")
        box_ax.set_ylabel("Point Value")
        box_ax.grid(True, linestyle='--', alpha=0.6)
        box_fig.subplots_adjust(right=0.7)
        box_fig.tight_layout()

        # 保存圖表與分組鍵的映射
        key = (gname, cname)
        chart_figures[key] = {'scatter': scatter_fig, 'box': box_fig}  # scatter實際上是SPC圖

        # 注意：不在這裡關閉 figure，因為 Streamlit 需要在顯示時使用它們
        # plt.close(scatter_fig)
        # plt.close(box_fig)

    return chart_figures

def display_tool_matching_charts(chart_figures, results_df):
    """在 Streamlit 中顯示 Tool Matching 圖表"""
    if not chart_figures:
        st.info("📈 沒有可顯示的圖表")
        return
    
    st.subheader("📊 Tool Matching 圖表分析")
    
    # 獲取可用的圖表選項
    chart_options = [f"{gname} - {cname}" for gname, cname in chart_figures.keys()]
    
    if len(chart_options) == 1:
        # 只有一個圖表，直接顯示
        selected_option = chart_options[0]
        gname, cname = list(chart_figures.keys())[0]
    else:
        # 多個圖表，提供選擇器
        selected_option = st.selectbox(
            "選擇要顯示的圖表組合",
            options=chart_options,
            index=0
        )
        
        # 找到對應的 key
        for (gname, cname) in chart_figures.keys():
            if f"{gname} - {cname}" == selected_option:
                break
    
    # 顯示選中的圖表
    if (gname, cname) in chart_figures:
        chart_data = chart_figures[(gname, cname)]
        
        # 檢查圖表是否成功生成
        if chart_data.get('scatter') is None and chart_data.get('box') is None:
            st.warning(f"⚠️ {selected_option} 的圖表資料不足，無法生成圖表")
            return
        
        # 使用兩欄布局並排顯示圖表
        chart_col1, chart_col2 = st.columns(2)
        
        # 顯示 SPC 圖
        with chart_col1:
            if chart_data.get('scatter') is not None:
                st.subheader("📈 SPC 圖表")
                st.pyplot(chart_data['scatter'])
            else:
                st.info("SPC 圖表資料不足")
        
        # 顯示盒鬚圖
        with chart_col2:
            if chart_data.get('box') is not None:
                st.subheader("📊 盒鬚圖")
                st.pyplot(chart_data['box'])
            else:
                st.info("盒鬚圖資料不足")
        
        # 顯示對應的統計資訊
        st.subheader(f"📋 {selected_option} 統計摘要")
        
        # 篩選相關結果
        chart_results = results_df[
            (results_df['GroupName'] == gname) & 
            (results_df['ChartName'] == cname)
        ]
        
        if len(chart_results) > 0:
            # 顯示統計表格
            display_cols = [
                "matching_group", "abnormal_type", "mean_matching_index", 
                "sigma_matching_index", "K", "mean", "sigma", "sample_size"
            ]
            
            # 只顯示存在的欄位
            available_cols = [col for col in display_cols if col in chart_results.columns]
            
            st.dataframe(
                chart_results[available_cols],
                use_container_width=True
            )
        else:
            st.info("沒有找到對應的統計結果")
    else:
        st.error("找不到對應的圖表資料")

def save_tool_matching_results_to_excel(results_df, chart_figures=None, source_filename="tool_matching_analysis"):
    """將分析結果匯出為 Excel 檔案，並在第一欄嵌入完整的盒鬚圖和散點圖。包含異常類型欄。"""
    try:
        # 檢查是否已安裝 openpyxl
        try:
            import openpyxl
            import openpyxl.styles
            import openpyxl.utils
            from openpyxl.drawing.image import Image as XLImage
        except ImportError:
            st.error("請安裝 openpyxl 以匯出 Excel 檔案。\n可在終端執行: pip install openpyxl")
            return None

        # 嘗試導入所需的模組
        try:
            import matplotlib.pyplot as plt
            import numpy as np
            import io
            from PIL import Image
            import matplotlib.cm as cm
        except ImportError as e:
            st.error(f"嵌入圖表需要額外套件: {str(e)}\n請安裝所需套件。")
            return None

        # 創建 BytesIO 用於返回
        output = BytesIO()

        # 調整 DataFrame 格式以匹配原始 Qt 版本
        # 需要將 results_df 轉換為 Qt 版本的格式
        df = results_df.copy()
        
        # 重新排列欄位順序以匹配原始格式
        column_mapping = {
            'is_abnormal': 'Need_matching',
            'abnormal_type': 'AbnormalType',
            'GroupName': 'GroupName',
            'ChartName': 'ChartName',
            'matching_group': 'matching_group',
            'mean_matching_index': 'mean_matching_index',
            'sigma_matching_index': 'sigma_matching_index',
            'K': 'K',
            'mean': 'mean',
            'sigma': 'sigma',
            'mean_median': 'mean_median',
            'sigma_median': 'sigma_median',
            'sample_size': 'samplesize'
        }
        
        # 重新命名欄位
        for old_name, new_name in column_mapping.items():
            if old_name in df.columns:
                df = df.rename(columns={old_name: new_name})
        
        # 確保所有必要欄位存在
        required_columns = [
            "Need_matching", "AbnormalType", "GroupName", "ChartName", "matching_group", 
            "mean_matching_index", "sigma_matching_index", "K", "mean", "sigma", 
            "mean_median", "sigma_median", "samplesize"
        ]
        
        for col in required_columns:
            if col not in df.columns:
                df[col] = ""

        # 打印資料框資訊以確認結構
        print(f"DataFrame info: {df.shape}")
        print(f"DataFrame columns: {df.columns.tolist()}")
        print(f"First row: {df.iloc[0].tolist() if len(df) > 0 else 'No data'}")

        # 創建臨時目錄用於保存圖片
        import tempfile
        temp_dir = tempfile.mkdtemp()
        print(f"[INFO] 創建臨時目錄: {temp_dir}")

        # 先在 DataFrame 前添加兩個空白欄位，分別用於SPC圖和盒鬚圖
        df.insert(0, "SPC_Chart", "")    # 第一欄：SPC圖
        df.insert(1, "BoxPlot", "")      # 第二欄：盒鬚圖

        # 創建臨時檔案路徑
        temp_excel_path = os.path.join(temp_dir, f"{source_filename}_matching_results.xlsx")
        
        # 創建 Excel 文件
        writer = pd.ExcelWriter(temp_excel_path, engine='openpyxl')
        df.to_excel(writer, sheet_name='Tool Matching Results', index=False)

        # 獲取工作表
        workbook = writer.book
        worksheet = writer.sheets['Tool Matching Results']

        # 設定標題列格式
        header_font = openpyxl.styles.Font(bold=True, color="FFFFFF")
        header_fill = openpyxl.styles.PatternFill(start_color="344CB7", end_color="344CB7", fill_type="solid")
        header_alignment = openpyxl.styles.Alignment(horizontal="center", vertical="center")

        # 設置標題列格式
        for cell in worksheet[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_alignment

        # 增加圖表欄寬度以容納圖片
        worksheet.column_dimensions['A'].width = 70  # 第一欄：SPC圖
        worksheet.column_dimensions['B'].width = 70  # 第二欄：盒鬚圖

        # 設定異常行的格式
        abnormal_fill = openpyxl.styles.PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")

        # 定義圖表在 Excel 中顯示的尺寸 (單位：像素)
        img_display_width, img_display_height = 450, 250

        # 檢查是否有可用的圖表數據
        has_chart_figures = chart_figures is not None and chart_figures
        if not has_chart_figures:
            print("[WARNING] 沒有可用的圖表數據，將使用簡單的狀態指示圖")

        # 從第二行開始遍歷（跳過標題行）
        for row_idx, row in enumerate(df.iterrows(), start=2):
            _, row_data = row

            # 檢查Need_matching欄位是否為True
            is_abnormal = row_data["Need_matching"]

            if is_abnormal:
                # 將整行設為淺紅色
                for cell in worksheet[row_idx]:
                    cell.fill = abnormal_fill

            # 創建並嵌入圖表到第一欄
            try:
                # 獲取關鍵數據
                group_name = str(row_data["GroupName"])
                chart_name = str(row_data["ChartName"])
                group_id = str(row_data["matching_group"])
                mean_index = row_data["mean_matching_index"]
                sigma_index = row_data["sigma_matching_index"]
                k_value = row_data["K"]

                # 檢查是否資料不足
                is_data_insufficient = (mean_index == '資料不足' or sigma_index == '資料不足' or k_value == '不比較')

                # 嘗試使用完整的SPC圖和盒鬚圖
                chart_key = (group_name, chart_name)
                if has_chart_figures and chart_key in chart_figures:
                    # 存在完整的分析圖表，使用實際的SPC圖和盒鬚圖
                    chart_data = chart_figures[chart_key]

                    # 1. 處理SPC圖 (放在第一欄)
                    try:
                        scatter_fig = chart_data['scatter']
                        temp_scatter_path = os.path.join(temp_dir, f"spc_{group_name}_{chart_name}_{row_idx}.png")
                        scatter_fig.savefig(temp_scatter_path, format='png', bbox_inches='tight', transparent=True, dpi=100)
                        try:
                            scatter_img = XLImage(temp_scatter_path)
                            scatter_img.width = img_display_width
                            scatter_img.height = img_display_height
                            scatter_position = f"A{row_idx}"
                            worksheet.add_image(scatter_img, scatter_position)
                            print(f"[INFO] 已添加SPC圖到單元格: {scatter_position}")
                        except Exception as img_e:
                            print(f"[ERROR] 添加SPC圖到 Excel 失敗: {img_e}")
                            worksheet.cell(row=row_idx, column=1).value = "SPC圖載入失敗"
                    except Exception as scatter_e:
                        print(f"[ERROR] 處理SPC圖時發生錯誤: {scatter_e}")
                        import traceback
                        traceback.print_exc()
                        worksheet.cell(row=row_idx, column=1).value = "SPC圖生成失敗"

                    # 2. 處理盒鬚圖 (放在第二欄)
                    try:
                        box_fig = chart_data['box']
                        temp_box_path = os.path.join(temp_dir, f"box_{group_name}_{chart_name}_{row_idx}.png")
                        box_fig.savefig(temp_box_path, format='png', bbox_inches='tight', transparent=True, dpi=100)
                        try:
                            box_img = XLImage(temp_box_path)
                            box_img.width = img_display_width
                            box_img.height = img_display_height
                            box_position = f"B{row_idx}"
                            worksheet.add_image(box_img, box_position)
                            print(f"[INFO] 已添加盒鬚圖到單元格: {box_position}")
                        except Exception as img_e:
                            print(f"[ERROR] 添加盒鬚圖到 Excel 失敗: {img_e}")
                            worksheet.cell(row=row_idx, column=2).value = "盒鬚圖載入失敗"
                    except Exception as box_e:
                        print(f"[ERROR] 處理盒鬚圖時發生錯誤: {box_e}")
                        import traceback
                        traceback.print_exc()
                        worksheet.cell(row=row_idx, column=2).value = "盒鬚圖生成失敗"

                else:
                    # 沒有找到匹配的圖表，使用狀態指示器
                    print(f"[INFO] 未找到 {group_name}/{chart_name} 的分析圖表，使用狀態指示器")
                    fig, ax = plt.subplots(figsize=(6, 4), dpi=100)
                    title = f"{group_name}\n{chart_name}\n組別: {group_id}"
                    ax.set_title(title, fontsize=12)
                    if is_data_insufficient:
                        circle = plt.Circle((0.5, 0.5), 0.3, color='yellow', alpha=0.6, edgecolor='goldenrod', linewidth=2)
                        ax.add_patch(circle)
                        ax.text(0.5, 0.5, "資料不足", ha='center', va='center', fontsize=14, color='black')
                        status_text = "資料不足，無法進行分析"
                    elif is_abnormal:
                        circle = plt.Circle((0.5, 0.5), 0.3, color='red', alpha=0.6, edgecolor='darkred', linewidth=2)
                        ax.add_patch(circle)
                        ax.text(0.5, 0.5, "需要對齊", ha='center', va='center', fontsize=14, color='white', fontweight='bold')
                        status_text = f"均值差異指數: {mean_index}, 標準差差異指數: {sigma_index}, K值: {k_value}"
                    else:
                        circle = plt.Circle((0.5, 0.5), 0.3, color='green', alpha=0.6, edgecolor='darkgreen', linewidth=2)
                        ax.add_patch(circle)
                        ax.text(0.5, 0.5, "正常", ha='center', va='center', fontsize=14, color='white', fontweight='bold')
                        status_text = f"均值差異指數: {mean_index}, 標準差差異指數: {sigma_index}, K值: {k_value}"
                    ax.text(0.5, 0.2, status_text, ha='center', va='center', fontsize=10, 
                           bbox=dict(boxstyle='round,pad=0.5', facecolor='white', alpha=0.8))
                    ax.set_xticks([])
                    ax.set_yticks([])
                    ax.set_xlim(0, 1)
                    ax.set_ylim(0, 1)
                    ax.set_aspect('equal')
                    temp_img_path = os.path.join(temp_dir, f"status_chart_{row_idx}.png")
                    plt.savefig(temp_img_path, format='png', bbox_inches='tight', transparent=True, dpi=300)
                    plt.close(fig)
                    try:
                        img1 = XLImage(temp_img_path)
                        img1.width = img_display_width
                        img1.height = img_display_height
                        cell_position_1 = f"A{row_idx}"
                        worksheet.add_image(img1, cell_position_1)
                        img2 = XLImage(temp_img_path)
                        img2.width = img_display_width
                        img2.height = img_display_height
                        cell_position_2 = f"B{row_idx}"
                        worksheet.add_image(img2, cell_position_2)
                        print(f"[INFO] 已添加狀態圖到單元格: {cell_position_1} 和 {cell_position_2}")
                    except Exception as img_e:
                        print(f"[ERROR] 添加圖片到 Excel 失敗: {img_e}")
                        worksheet.cell(row=row_idx, column=1).value = "圖片載入失敗"
                        worksheet.cell(row=row_idx, column=2).value = "圖片載入失敗"

            except Exception as img_e:
                print(f"[ERROR] 在第 {row_idx} 行添加圖表時發生錯誤: {img_e}")
                import traceback
                traceback.print_exc()
                worksheet.cell(row=row_idx, column=1).value = "圖片生成失敗"

        # 調整行高以適應圖表
        for i in range(2, worksheet.max_row + 1):
            worksheet.row_dimensions[i].height = 190

        # 調整其他列寬
        for col_idx, column in enumerate(worksheet.columns, start=1):
            if col_idx <= 2:  # 跳過圖表列 A 和 B，已手動設置寬度
                continue
            max_length = 0
            column_letter = openpyxl.utils.get_column_letter(col_idx)
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 4)
            worksheet.column_dimensions[column_letter].width = adjusted_width

        # 儲存 Excel 檔案
        try:
            writer.close()
            print(f"[INFO] Excel 檔案已儲存到: {temp_excel_path}")
            
            # 讀取文件內容到 BytesIO
            with open(temp_excel_path, 'rb') as f:
                output.write(f.read())
            output.seek(0)
            
        except Exception as save_e:
            print(f"[ERROR] 儲存 Excel 檔案失敗: {save_e}")
            import traceback
            traceback.print_exc()
            return None
        finally:
            try:
                import shutil
                shutil.rmtree(temp_dir)
                print(f"[INFO] 已清理臨時目錄: {temp_dir}")
            except Exception as e:
                print(f"[WARNING] 無法清理臨時目錄: {temp_dir}, 錯誤: {e}")

        return output
        
    except Exception as e:
        print(f"匯出 Excel 失敗: {e}")
        import traceback
        traceback.print_exc()
        return None

def render_tool_matching_module():
    """渲染 Tool Matching 模組"""
    st.header("🔧 Tool Matching 分析")
    
    # 初始化 session_state
    if 'tool_matching_results' not in st.session_state:
        st.session_state['tool_matching_results'] = None
    
    # 檔案上傳區
    st.subheader("📂 檔案輸入")
    
    col1, col2 = st.columns([3, 1])
    
    with col1:
        uploaded_file = st.file_uploader(
            "上傳 Tool Matching CSV 檔案",
            type=["csv"],
            help="請上傳包含 GroupName, ChartName, matching_group, point_val, characteristic, point_time 欄位的 CSV 檔案"
        )
    
    with col2:
        if st.button("📄 下載範例檔案"):
            example_df = generate_temp_csv()
            csv_buffer = BytesIO()
            example_df.to_csv(csv_buffer, index=False, encoding='utf-8-sig')
            csv_buffer.seek(0)
            
            st.download_button(
                label="📥 下載範例 CSV",
                data=csv_buffer.getvalue(),
                file_name="tool_matching_example.csv",
                mime="text/csv",
                help="下載 Tool Matching 輸入格式範例"
            )
    
    # 分析參數設定
    st.subheader("⚙️ 分析參數")
    
    # 門檻設定
    col1, col2 = st.columns(2)
    
    with col1:
        use_mean_threshold = st.checkbox("啟用 Mean Index 門檻", value=False)
        mean_threshold = st.number_input(
            "Mean Index 門檻值",
            min_value=0.0,
            max_value=10.0,
            value=1.0,
            step=0.1,
            disabled=not use_mean_threshold
        )
    
    with col2:
        use_sigma_threshold = st.checkbox("啟用 Sigma Index 門檻", value=False)
        sigma_threshold = st.number_input(
            "Sigma Index 門檻值",
            min_value=0.0,
            max_value=10.0,
            value=2.0,
            step=0.1,
            disabled=not use_sigma_threshold
        )
    
    # 補滿樣本數設定
    col1, col2 = st.columns(2)
    
    with col1:
        use_fill_num = st.checkbox("啟用補滿樣本數", value=False)
        fill_num = st.number_input( 
            "補滿樣本數",
            min_value=1,
            max_value=100,
            value=5,
            disabled=not use_fill_num
        )
    
    with col2:
        filter_mode = st.selectbox(
            "資料篩選模式",
            options=["全算", "指定日期(一個月mean/半年sigma)", "最新進點(一個月mean/半年sigma)"],
            index=0
        )
        
        # 為了保持界面對齊，總是顯示日期輸入框，但根據模式決定是否啟用
        is_date_mode = filter_mode == "指定日期(一個月mean/半年sigma)"
        base_date = st.date_input(
            "指定基準日", 
            value=pd.Timestamp.now().date(),
            disabled=not is_date_mode,
            help="僅在選擇'指定日期'模式時生效"
        ) if is_date_mode else None
    
    # 顯示計算公式說明
    with st.expander("📘 計算公式說明", expanded=False):
        st.markdown("""
        **計算公式：**
        
        **Mean Matching Index：**
        - 兩組比較：`|μ₁ − μ₂| / min(σ₁, σ₂)`
        - 多組比較：`|μ − median(μ)| / median(σ)`
        
        **Sigma Matching Index：**
        - 兩組比較：`σ / min(σ₁, σ₂)`
        - 多組比較：`σ / median(σ)`
        
        **K 值：**
        - n ≤ 4：不比較
        - 5 ≤ n ≤ 10：K = 1.73
        - 11 ≤ n ≤ 120：K = 1.414
        - n > 120：K = 1.15
        
        **篩選模式說明：**
        - **全算**：使用所有資料進行分析
        - **指定日期**：從指定日期往前一個月計算 mean，往前半年計算 sigma median
        - **最新進點**：從最新資料點往前一個月計算 mean，往前半年計算 sigma median
        """)
    
    # 執行分析按鈕
    if st.button("🚀 執行 Tool Matching 分析", type="primary"):
        if uploaded_file is None:
            st.warning("請先上傳 CSV 檔案！")
        else:
            with st.spinner("分析中，請稍候..."):
                try:
                    # 讀取檔案
                    df = pd.read_csv(uploaded_file)
                    
                    # 檢查必要欄位
                    required_cols = ["GroupName", "ChartName", "matching_group", "point_val", "characteristic", "point_time"]
                    missing_cols = [col for col in required_cols if col not in df.columns]
                    
                    if missing_cols:
                        st.error(f"缺少必要欄位: {', '.join(missing_cols)}")
                        st.stop()
                    
                    # 轉換 point_time 為 datetime
                    df["point_time"] = pd.to_datetime(df["point_time"])
                    
                    # 執行分析
                    results = []
                    chart_figures = {}
                    filter_mode_idx = ["全算", "指定日期(一個月mean/半年sigma)", "最新進點(一個月mean/半年sigma)"].index(filter_mode)
                    
                    if filter_mode_idx == 0:
                        # 全算模式
                        grouped = df.groupby(["GroupName", "ChartName"])
                        for (gname, cname), subdf in grouped:
                            characteristic = subdf["characteristic"].dropna().unique()
                            if len(characteristic) != 1:
                                st.warning(f"Group: {gname}-{cname} 的 characteristic 不唯一或缺失")
                                continue
                            
                            group_stats = subdf.groupby("matching_group")["point_val"].agg(['mean', 'std', 'count']).reset_index()
                            n_groups = len(group_stats)
                            
                            if n_groups == 2:
                                analyze_two_groups(group_stats, gname, cname, characteristic[0], results)
                            else:
                                analyze_multiple_groups(subdf, group_stats, gname, cname, characteristic[0], results)
                        
                        # 生成圖表
                        chart_figures = create_tool_matching_charts(grouped)
                    
                    elif filter_mode_idx == 1:
                        # 指定日期模式
                        grouped = df.groupby(["GroupName", "ChartName"])
                        for (gname, cname), subdf in grouped:
                            characteristic = subdf["characteristic"].dropna().unique()
                            if len(characteristic) != 1:
                                continue
                                
                            mean_end = pd.Timestamp(base_date)
                            sigma_end = pd.Timestamp(base_date)
                            mean_start = mean_end - pd.DateOffset(months=1)
                            sigma_start = sigma_end - pd.DateOffset(months=6)
                            
                            mean_df = subdf[(subdf["point_time"] > mean_start) & (subdf["point_time"] <= mean_end)].copy()
                            sigma_df = subdf[(subdf["point_time"] > sigma_start) & (subdf["point_time"] <= sigma_end)].copy()
                            
                            # 補足邏輯
                            if use_fill_num:
                                min_time = subdf["point_time"].min()
                                for mg in subdf["matching_group"].unique():
                                    mg_mean = mean_df[mean_df["matching_group"] == mg]
                                    if len(mg_mean) < fill_num:
                                        all_mg = subdf[subdf["matching_group"] == mg].sort_values("point_time")
                                        cur_start = mean_start
                                        while len(mg_mean) < fill_num and cur_start > min_time:
                                            cur_start = cur_start - pd.Timedelta(days=7)
                                            mg_mean = all_mg[(all_mg["point_time"] > cur_start) & (all_mg["point_time"] <= mean_end)]
                                        mean_df = pd.concat([mean_df, mg_mean]).drop_duplicates()
                                
                                for mg in subdf["matching_group"].unique():
                                    mg_sigma = sigma_df[sigma_df["matching_group"] == mg]
                                    if len(mg_sigma) < fill_num:
                                        all_mg = subdf[subdf["matching_group"] == mg].sort_values("point_time")
                                        cur_start = sigma_start
                                        while len(mg_sigma) < fill_num and cur_start > min_time:
                                            cur_start = cur_start - pd.Timedelta(days=14)
                                            mg_sigma = all_mg[(all_mg["point_time"] > cur_start) & (all_mg["point_time"] <= sigma_end)]
                                        sigma_df = pd.concat([sigma_df, mg_sigma]).drop_duplicates()
                            
                            mean_stats = mean_df.groupby("matching_group")["point_val"].agg(['mean', 'count']).reset_index()
                            sigma_stats = sigma_df.groupby("matching_group")["point_val"].agg(['std']).reset_index()
                            group_stats = pd.merge(mean_stats, sigma_stats, on="matching_group", how="outer")
                            group_stats = group_stats.fillna({"mean": 0, "std": 0, "count": 0})
                            
                            n_groups = len(group_stats)
                            if n_groups == 2:
                                analyze_two_groups(group_stats, gname, cname, characteristic[0], results)
                            else:
                                analyze_multiple_groups_time(mean_df, sigma_df, group_stats, gname, cname, characteristic[0], results)
                        
                        # 生成圖表 (使用一個月資料)
                        grouped_for_chart = df.groupby(["GroupName", "ChartName"])
                        filtered_grouped = {}
                        for (gname, cname), subdf in grouped_for_chart:
                            mean_end = pd.Timestamp(base_date)
                            mean_start = mean_end - pd.DateOffset(months=1)
                            chart_df = subdf[(subdf["point_time"] > mean_start) & (subdf["point_time"] <= mean_end)]
                            if len(chart_df) > 0:
                                filtered_grouped[(gname, cname)] = chart_df
                        chart_figures = create_tool_matching_charts(filtered_grouped.items())
                    
                    else:
                        # 最新進點模式
                        grouped = df.groupby(["GroupName", "ChartName"])
                        for (gname, cname), subdf in grouped:
                            characteristic = subdf["characteristic"].dropna().unique()
                            if len(characteristic) != 1:
                                continue
                                
                            latest_time = subdf["point_time"].max()
                            mean_end = latest_time
                            sigma_end = latest_time
                            mean_start = mean_end - pd.DateOffset(months=1)
                            sigma_start = sigma_end - pd.DateOffset(months=6)
                            
                            mean_df = subdf[(subdf["point_time"] > mean_start) & (subdf["point_time"] <= mean_end)].copy()
                            sigma_df = subdf[(subdf["point_time"] > sigma_start) & (subdf["point_time"] <= sigma_end)].copy()
                            
                            # 補足邏輯（同指定日期模式）
                            if use_fill_num:
                                min_time = subdf["point_time"].min()
                                for mg in subdf["matching_group"].unique():
                                    mg_mean = mean_df[mean_df["matching_group"] == mg]
                                    if len(mg_mean) < fill_num:
                                        all_mg = subdf[subdf["matching_group"] == mg].sort_values("point_time")
                                        cur_start = mean_start
                                        while len(mg_mean) < fill_num and cur_start > min_time:
                                            cur_start = cur_start - pd.Timedelta(days=7)
                                            mg_mean = all_mg[(all_mg["point_time"] > cur_start) & (all_mg["point_time"] <= mean_end)]
                                        mean_df = pd.concat([mean_df, mg_mean]).drop_duplicates()
                                
                                for mg in subdf["matching_group"].unique():
                                    mg_sigma = sigma_df[sigma_df["matching_group"] == mg]
                                    if len(mg_sigma) < fill_num:
                                        all_mg = subdf[subdf["matching_group"] == mg].sort_values("point_time")
                                        cur_start = sigma_start
                                        while len(mg_sigma) < fill_num and cur_start > min_time:
                                            cur_start = cur_start - pd.Timedelta(days=14)
                                            mg_sigma = all_mg[(all_mg["point_time"] > cur_start) & (all_mg["point_time"] <= sigma_end)]
                                        sigma_df = pd.concat([sigma_df, mg_sigma]).drop_duplicates()
                            
                            mean_stats = mean_df.groupby("matching_group")["point_val"].agg(['mean', 'count']).reset_index()
                            sigma_stats = sigma_df.groupby("matching_group")["point_val"].agg(['std']).reset_index()
                            group_stats = pd.merge(mean_stats, sigma_stats, on="matching_group", how="outer")
                            group_stats = group_stats.fillna({"mean": 0, "std": 0, "count": 0})
                            
                            n_groups = len(group_stats)
                            if n_groups == 2:
                                analyze_two_groups(group_stats, gname, cname, characteristic[0], results)
                            else:
                                analyze_multiple_groups_time(mean_df, sigma_df, group_stats, gname, cname, characteristic[0], results)
                        
                        # 生成圖表 (使用一個月資料)
                        grouped_for_chart = df.groupby(["GroupName", "ChartName"])
                        filtered_grouped = {}
                        for (gname, cname), subdf in grouped_for_chart:
                            latest_time = subdf["point_time"].max()
                            mean_start = latest_time - pd.DateOffset(months=1)
                            chart_df = subdf[(subdf["point_time"] > mean_start) & (subdf["point_time"] <= latest_time)]
                            if len(chart_df) > 0:
                                filtered_grouped[(gname, cname)] = chart_df
                        chart_figures = create_tool_matching_charts(filtered_grouped.items())
                    
                    # 處理結果
                    if results:
                        # 轉換為 DataFrame
                        results_df = pd.DataFrame(results, columns=[
                            "GroupName", "ChartName", "matching_group", "group_type",
                            "mean_matching_index", "sigma_matching_index", "K",
                            "mean", "sigma", "mean_median", "sigma_median", "sample_size"
                        ])
                        
                        # 篩選異常項目
                        abnormal_results = []
                        all_results = []
                        
                        for _, row in results_df.iterrows():
                            is_abnormal = False
                            abnormal_type = ""
                            
                            mean_index = row["mean_matching_index"]
                            sigma_index = row["sigma_matching_index"]
                            k_value = row["K"]
                            
                            is_data_insufficient = (mean_index == '資料不足' or sigma_index == '資料不足' or k_value == '不比較')
                            
                            if not is_data_insufficient:
                                try:
                                    # 使用設定的門檻值
                                    mean_thresh = mean_threshold if use_mean_threshold else 1.0
                                    sigma_thresh = sigma_threshold if use_sigma_threshold else float(k_value) if k_value not in [None, '', '不比較'] else 2.0
                                    
                                    mean_abn = float(mean_index) >= mean_thresh
                                    sigma_abn = float(sigma_index) >= sigma_thresh
                                    
                                    if mean_abn or sigma_abn:
                                        is_abnormal = True
                                        if mean_abn and sigma_abn:
                                            abnormal_type = "Mean, Sigma"
                                        elif mean_abn:
                                            abnormal_type = "Mean"
                                        elif sigma_abn:
                                            abnormal_type = "Sigma"
                                except (ValueError, TypeError):
                                    pass
                            else:
                                abnormal_type = "資料不足"
                            
                            row_with_abnormal = row.copy()
                            row_with_abnormal["abnormal_type"] = abnormal_type
                            row_with_abnormal["is_abnormal"] = is_abnormal or is_data_insufficient
                            
                            all_results.append(row_with_abnormal)
                            
                            if is_abnormal or is_data_insufficient:
                                abnormal_results.append(row_with_abnormal)
                        
                        # 儲存到 session_state
                        st.session_state['tool_matching_results'] = {
                            'all_results': pd.DataFrame(all_results),
                            'abnormal_results': pd.DataFrame(abnormal_results),
                            'chart_figures': chart_figures
                        }
                        
                        st.success(f"✅ 分析完成！發現 {len(abnormal_results)} 個需注意項目（總共 {len(all_results)} 項）")
                        
                    else:
                        st.error("❌ 分析失敗，請檢查資料格式")
                        
                except Exception as e:
                    st.error(f"❌ 分析過程發生錯誤：{str(e)}")
                    st.exception(e)
    
    # 顯示結果
    if st.session_state.get('tool_matching_results') is not None:
        st.subheader("📊 分析結果")
        
        results_data = st.session_state['tool_matching_results']
        all_df = results_data['all_results']
        abnormal_df = results_data['abnormal_results']
        chart_figures = results_data.get('chart_figures', {})
        
        # 選擇顯示模式
        view_mode = st.radio(
            "選擇顯示模式",
            options=["僅顯示異常項目", "顯示所有結果"],
            index=0,
            horizontal=True
        )
        
        display_df = abnormal_df if view_mode == "僅顯示異常項目" else all_df
        
        if len(display_df) > 0:
            # 顯示結果表格
            st.dataframe(
                display_df[[
                    "GroupName", "ChartName", "matching_group", "abnormal_type",
                    "mean_matching_index", "sigma_matching_index", "K",
                    "mean", "sigma", "mean_median", "sigma_median", "sample_size"
                ]],
                use_container_width=True
            )
            
            # 顯示圖表
            if chart_figures:
                st.markdown("---")
                display_tool_matching_charts(chart_figures, all_df)
            
            # 下載報告
            st.subheader("📥 下載報告")
            
            try:
                # 從上傳的檔案取得檔名
                uploaded_filename = "tool_matching_analysis"
                if uploaded_file is not None:
                    uploaded_filename = os.path.splitext(uploaded_file.name)[0]
                
                excel_buffer = save_tool_matching_results_to_excel(
                    all_df, 
                    chart_figures=chart_figures,
                    source_filename=uploaded_filename
                )
                
                if excel_buffer is not None:
                    st.download_button(
                        label="📥 下載完整分析報告 (含圖表)",
                        data=excel_buffer.getvalue(),
                        file_name=f"Tool_Matching_Results_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        help="下載包含所有分析結果與嵌入圖表的 Excel 報告"
                    )
                else:
                    st.error("❌ 生成報告失敗，請檢查是否已安裝必要套件")
                    
            except Exception as e:
                st.error(f"❌ 生成報告時發生錯誤: {str(e)}")
                import traceback
                traceback.print_exc()
        else:
            st.info("📋 沒有找到需要注意的項目")
    else:
        st.info("請先執行 Tool Matching 分析")

# 主要頁面布局：左側控制面板 + 右側內容區域
left_panel, right_content = st.columns([1, 4])

with left_panel:
    st.header("🎛️ 功能模組")
    
    # SPC 圖表處理模組的勾選框
    show_spc_module = st.checkbox(
        "📊 SPC 圖表分析模組", 
        value=True,  # 預設勾選
        help="包含圖表處理、圖片展示和儀表板總覽功能"
    )
    
    # Tool Matching 模組的勾選框
    show_tool_matching_module = st.checkbox(
        "🔧 Tool Matching 分析模組",
        value=False,
        help="設備間的均值和標準差匹配分析"
    )
    
    st.markdown("---")
    st.write("🔧 其他模組")
    st.write("(模組開發中...)")
    
    # 預留其他模組的勾選框
    # module_3 = st.checkbox("📋 模組三", value=False)

with right_content:
    # 只有勾選 SPC 模組時才顯示三標籤頁
    if show_spc_module:
        tab1, tab2, tab3 = st.tabs(["圖表處理", "圖片展示", "儀表板總覽"])

        with tab1:
            st.header("📈 圖表處理")
            st.subheader("📂 檔案輸入")

            uploaded_chart_info_file = st.file_uploader(
                "上傳 'All_Chart_Information.xlsx' 檔案",
                type=["xlsx"],
                help="請上傳包含所有圖表資訊的 Excel 檔案。"
            )

            uploaded_raw_data_files = st.file_uploader(
                "上傳多份原始資料 CSV 檔案 (例如: GroupName_ChartName.csv)",
                type=["csv"],
                accept_multiple_files=True,
                help="請上傳所有 SPC 圖表的原始資料 CSV 檔案。檔名格式建議為 '組別名稱_圖表名稱.csv'。"
            )

            if st.button("🚀 開始處理圖表"):
                if not uploaded_chart_info_file or not uploaded_raw_data_files:
                    st.warning("請同時上傳圖表資訊檔案和原始資料 CSV 檔案。")
                else:
                    with st.spinner("處理中，請稍候..."):
                        try:
                            all_charts_info_df = pd.read_excel(uploaded_chart_info_file)
                            with tempfile.TemporaryDirectory() as tmpdirname:
                                for f in uploaded_raw_data_files:
                                    dst = os.path.join(tmpdirname, f.name)
                                    with open(dst, 'wb') as out:
                                        out.write(f.getbuffer())

                                results, skipped_count = process_all_charts(tmpdirname, all_charts_info_df)
                                st.session_state['results'] = results
                                st.session_state['skipped_count'] = skipped_count

                                if not results:
                                    st.error("⚠️ 所有圖表處理都失敗，請檢查資料格式。")
                                else:
                                    st.success(f"✅ 圖表處理完成：{len(results)} 筆成功，{skipped_count} 筆略過。")
                        except Exception:
                            st.error("🚫 圖表處理發生錯誤，請查看下方訊息。")
                            st.text(traceback.format_exc())

            # 新增下載圖表報告按鈕
            st.subheader("📥 下載報告")
            if st.session_state.get('results') is not None:
                df_results = pd.DataFrame(st.session_state['results'])
                
                try:
                    excel_buffer = save_results_to_excel(df_results)
                    st.download_button(
                        label="📥 下載圖表報告",
                        data=excel_buffer.getvalue(),
                        file_name=f"Chart_Analysis_Results_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        help="點擊直接下載包含圖表的 Excel 報告"
                    )
                except Exception as e:
                    st.error(f"❌ 生成圖表報告時發生錯誤: {str(e)}")
            else:
                st.info("請先完成圖表處理後才能下載報告。")

        with tab2:
            st.header("📸 圖片展示區")

            if st.session_state.get('results') is None:
                st.info("請先在「圖表處理」頁面執行處理流程。")
            else:
                df = pd.DataFrame(st.session_state['results'])
                
                # 檢查必要的圖片欄位是否存在
                if 'chart_path' not in df.columns or 'weekly_chart_path' not in df.columns:
                    st.error("❌ 結果資料中找不到圖片路徑欄位，請重新處理圖表。")
                    st.info("需要的欄位：chart_path, weekly_chart_path")
                    if len(df.columns) > 0:
                        st.write("目前可用的欄位：", list(df.columns))
                else:
                    for idx, row in df.iterrows():
                        # 顯示圖表資訊
                        group_name = row.get('group_name', 'Unknown')
                        chart_name = row.get('chart_name', 'Unknown')
                        characteristics = row.get('Characteristics', 'Unknown')
                        data_type = row.get('data_type', 'Unknown')
                        
                        st.markdown(f"### 📊 第 {idx+1} 筆：{group_name} - {chart_name}")
                        st.markdown(f"**特性**: {characteristics} | **數據類型**: {data_type}")

                        # 改為三欄布局
                        col1, col2, col3 = st.columns([2, 2, 1])

                        with col1:
                            st.subheader("🔄 完整 SPC 圖表")
                            chart_path = row['chart_path']
                            if isinstance(chart_path, str) and os.path.exists(chart_path):
                                st.image(chart_path, caption=f"SPC Chart - {chart_name}", use_column_width=True)
                                st.caption(f"路徑: {chart_path}")
                            else:
                                st.warning(f"⚠️ 無法讀取 SPC 圖表")
                                st.write(f"圖片路徑: {chart_path}")
                                st.write(f"路徑類型: {type(chart_path)}")
                                if isinstance(chart_path, str):
                                    st.write(f"檔案是否存在: {os.path.exists(chart_path)}")

                        with col2:
                            st.subheader("📅 週數據圖表")
                            weekly_path = row['weekly_chart_path']
                            if isinstance(weekly_path, str) and os.path.exists(weekly_path):
                                st.image(weekly_path, caption=f"Weekly Chart - {chart_name}", use_column_width=True)
                                st.caption(f"路徑: {weekly_path}")
                            else:
                                st.warning(f"⚠️ 無法讀取週數據圖表")
                                st.write(f"圖片路徑: {weekly_path}")
                                st.write(f"路徑類型: {type(weekly_path)}")
                                if isinstance(weekly_path, str):
                                    st.write(f"檔案是否存在: {os.path.exists(weekly_path)}")

                        with col3:
                            st.subheader("📋 分析摘要")
                            
                            # 建立垂直表格資料
                            summary_data = {
                                "項目": [
                                    "組別名稱",
                                    "圖表名稱", 
                                    "特性",
                                    "數據點數",
                                    "OOC點數",
                                    "WE規則",
                                    "OOB規則",
                                    "Cpk值"
                                ],
                                "值": [
                                    str(row.get('group_name', 'N/A')),
                                    str(row.get('chart_name', 'N/A')),
                                    str(row.get('Characteristics', 'N/A')),
                                    str(row.get('data_cnt', 'N/A')),
                                    str(row.get('ooc_cnt', 'N/A')),
                                    str(row.get('WE_Rule', 'N/A')),
                                    str(row.get('OOB_Rule', 'N/A')),
                                    str(row.get('Cpk', 'N/A'))
                                ]
                            }
                            
                            summary_df = pd.DataFrame(summary_data)
                            
                            # 顯示垂直表格，不顯示索引
                            st.dataframe(
                                summary_df, 
                                use_container_width=True,
                                hide_index=True
                            )

        st.markdown("---")

        with tab3:
            st.header("📊 儀表板總覽")

            if st.session_state['results'] is None:
                st.info("請先到【圖表處理】頁面執行處理流程。")
            else:
                df_dashboard = pd.DataFrame(st.session_state['results'])
                
                # 顯示資料表
                st.dataframe(df_dashboard)
                
    elif show_tool_matching_module:
        # Tool Matching 模組內容
        render_tool_matching_module()
        
    else:
        st.info("🔘 請在左側勾選模組來使用相應功能。")
        st.markdown("### 📋 可用模組說明")
        st.markdown("""
        - **📊 SPC 圖表分析模組**：包含完整的 SPC 圖表處理、分析和展示功能
          - 圖表處理：上傳資料並進行 OOB 分析
          - 圖片展示：查看 SPC 圖表和分析結果
          - 儀表板總覽：完整的分析結果總覽
        
        - **🔧 Tool Matching 分析模組**：設備間的均值和標準差匹配分析
          - 根據 GroupName + ChartName 分組
          - 計算 Mean/Sigma Matching Index
          - 支持多種篩選模式（全算、指定日期、最新進點）
          - 提供詳細的異常項目報告
        
        - **其他模組**：開發中...
        """)
    
    # 預留其他模組的顯示區域
    # if module_3:
    #     st.write("模組三的內容")  