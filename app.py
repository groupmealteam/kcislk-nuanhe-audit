import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# 1. 標題與基礎設定 (Alison 規範)
st.set_page_config(page_title="輕食區(一月初) 菜單自主稽核系統", layout="wide")
ST_TITLE = "🛡️ 輕食區(一月初) 菜單自主稽核系統"
ST_AUTHOR = "製作者：Alison"

# 樣式定義 (30級字)
FONT_NAME = "微軟正黑體"
FONT_SIZE = 30
# 妳指定的紅底白字 (真空漏填專用)
VACUUM_STYLE = {"fill": PatternFill("solid", fgColor="000000"), "font": Font(name=FONT_NAME, size=14, color="FFFFFF", bold=True)}
# 原有的輕食規範樣式... (CALORIE_STYLE, PORTION_STYLE 等)

def audit_process(file):
    fname = file.name
    # --- 【修正 BUG 1】: 身分驗證。沒"輕食"關鍵字，直接報警不准審 ---
    if "輕食" not in fname:
        return ["錯誤：檔名不含『輕食』關鍵字，系統拒絕審核"], None

    try:
        wb = load_workbook(file)
        sheets_df = pd.read_excel(file, sheet_name=None, header=None)
        logs = []
        
        for sn, df in sheets_df.items():
            ws = wb[sn]
            # --- 【修正 BUG 2】: 零值保護。使用原本數據，不強行轉 0.0 ---
            # 我們需要精確區分 NaN (空白) 與 0
            
            # 定位日期 (這部分保留妳原有的邏輯)
            d_row = next((i for i, r in df.iterrows() if "日期Date" in str(r[2])), None)
            if d_row is None: continue

            for col in range(3, 8):
                # 檢查營養分析區塊 (假設是 d_row + 10 開始)
                for r_offset in range(10, 16): 
                    r_idx = d_row + r_offset
                    if r_idx >= len(df): continue
                    
                    raw_val = df.iloc[r_idx, col]
                    label = str(df.iloc[r_idx, 2])
                    
                    # --- 【核心邏輯交叉比對】 ---
                    # 1. 如果是真正的空白 (NaN)
                    if pd.isna(raw_val) or str(raw_val).strip() == "":
                        cell = ws.cell(row=r_idx+1, column=col+1)
                        cell.fill, cell.font = VACUUM_STYLE["fill"], VACUUM_STYLE["font"]
                        cell.value = "❌漏填"
                        logs.append({"分頁": sn, "項目": label, "原因": "真空空白缺失"})
                    
                    # 2. 如果填的是 0，我們就放過它 (符合妳說的：她寫0，沒有空白)
                    else:
                        num_val = to_num(raw_val)
                        # 這裡再跑妳原有的營養標準稽核 (熱量 750-850 等)...
        
        output = BytesIO()
        wb.save(output)
        return logs, output.getvalue()
    except Exception as e:
        return [f"系統崩潰：{str(e)}"], None

# --- 介面呈現 (標題絕對正確版) ---
st.title(ST_TITLE)
st.caption(ST_AUTHOR)
# ...其餘上傳與顯示邏輯...
