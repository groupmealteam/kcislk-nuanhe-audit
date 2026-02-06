import streamlit as st
import pandas as pd
import re
import numpy as np
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# 1. 網頁基本設定
st.set_page_config(page_title="輕食區(一月初) 菜單自主稽核系統", layout="wide")

# 顏色標註設定
RED_FILL = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")    # 違規標色
YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid") # 提醒標色

# 教育部國高中營養基準 (法規標準)
STANDARD = {
    "熱量": (750, 850),
    "全榖": 4.0,
    "蛋白質": 4.0,
    "蔬菜": 2.0
}

def to_float(val):
    """將包含文字的數據轉為純數字，例如 '4.2份' -> 4.2"""
    try:
        nums = re.findall(r"\d+\.?\d*", str(val))
        return float(nums[0]) if nums else 0.0
    except:
        return 0.0

def clean_dish_name(text):
    if pd.isna(text): return ""
    # 僅提取第一行中文字，移除英文描述
    return "".join(re.findall(r'[\u4e00-\u9fa5]+', str(text).split('\n')[0]))

def audit_process(file):
    try:
        # 讀取 Excel
        wb = load_workbook(file, data_only=True)
        # 用 pandas 讀取原始數據來定位座標
        all_sheets_df = pd.read_excel(file, sheet_name=None, header=None)
    except Exception as e:
        return [f"❌ 檔案開啟失敗：{str(e)}"], None

    results = []
    output = BytesIO()
    found_any_valid_sheet = False

    for sheet_name, df in all_sheets_df.items():
        df = df.fillna("")
        ws = wb[sheet_name]
        
        # 尋找「日期Date」所在行
        date_row = None
        for i, row in df.iterrows():
            if "日期Date" in str(row[2]):
                date_row = i
                break
        
        if date_row is None: continue
        found_any_valid_sheet = True

        # 遍歷週一到週五 (D欄到H欄, Index 3 到 7)
        for col in range(3, 8):
            if col >= len(df.columns): break
            
            date_val = str(df.iloc[date_row, col]).split(" ")[0]
            day_val = str(df.iloc[date_row + 1, col])
            if not date_val or "202" not in date_val: continue

            # --- A. 營養成分稽核 (教育部法規) ---
            for r_idx in range(date_row + 5, len(df)):
                label = str(df.iloc[r_idx, 2]) # B欄標籤
                val = to_float(df.iloc[r_idx, col])
                
                # 判斷標籤與數值
                if "熱量" in label:
                    if val < STANDARD["熱量"][0] or val > STANDARD["熱量"][1]:
                        ws.cell(row=r_idx+1, column=col+1).fill = RED_FILL
                        results.append({"分頁": sheet_name, "日期": date_val, "項目": "熱量", "原因": f"❌ 不合法規：{val} Kcal (標示需在750-850間)"})
                elif "全榖" in label and val < STANDARD["全榖"]:
                    ws.cell(row=r_idx+1, column=col+1).fill = RED_FILL
                    results.append({"分頁": sheet_name, "日期": date_val, "項目": "全榖雜糧", "原因": f"❌ 份數不足：{val} 份 (低於法規 4 份)"})
                elif "豆魚蛋肉" in label and val < STANDARD["蛋白質"]:
                    ws.cell(row=r_idx+1, column=col+1).fill = RED_FILL
                    results.append({"分頁": sheet_name, "日期": date_val, "項目": "蛋白質", "原因": f"❌ 份數不足：{val} 份 (低於法規 4 份)"})
                elif "蔬菜" in label and val < STANDARD["蔬菜"]:
                    ws.cell(row=r_idx+1, column=col+1).fill = RED_FILL
                    results.append({"分頁": sheet_name, "日期": date_val, "項目": "蔬菜類", "原因": f"❌ 份數不足：{val} 份 (低於法規 2 份)"})

            # --- B. 禁辣日審核 (原則四) ---
            if any(d in day_val for d in ["週一", "週二", "週四"]):
                for r_idx in range(date_row + 2, date_row + 15):
                    cell_val = str(df.iloc[r_idx, col])
                    if "●" in cell_val or "🌶️" in cell_val:
                        ws.cell(row=r_idx+1, column=col+1).fill = RED_FILL
                        results.append({"分頁": sheet_name, "日期": date_val, "項目": "禁辣日違規", "原因": f"🚫 週一二四不應有辣標：{cell_val}"})

            # --- C. 餐道重複比對 (A vs B) ---
            # 已排除水果與湯品，只抓主食名稱
            main_A = clean_dish_name(df.iloc[date_row + 3, col])
            main_B = ""
            for r_idx in range(date_row + 10, len(df)):
                if "輕食B餐" in str(df.iloc[r_idx, 2]):
                    main_B = clean_dish_name(df.iloc[r_idx + 1, col])
                    break
            
            if main_A and main_B:
                core_A, core_B = main_A[:2], main_B[:2]
                if core_A == core_B:
                    results.append({"分頁": sheet_name, "日期": date_val, "項目": "多樣性建議", "原因": f"⚠️ A/B主食雷同({core_A})"})

    if not found_any_valid_sheet:
        return ["❌ 偵測失敗：找不到日期座標，請確認 Excel 是否為 V5 格式。"], None

    wb.save(output)
    return results, output.getvalue()

# --- 介面呈現 ---
st.title("🛡️ 輕食區(一月初) 菜單自主稽核系統")
st.caption("製作者：Alison / 版本：V5")
st.markdown("---")

up_file = st.file_uploader("👉 請上傳一月初週菜單 Excel (.xlsx)", type=["xlsx"])

if up_file:
    with st.spinner("系統檢查中..."):
        logs, final_data = audit_process(up_file)
        
        if logs and isinstance(logs[0], str) and "❌" in logs[0]:
            st.error(logs[0])
        elif logs:
            st.error(f"🚩 發現 {len(logs)} 項需修正項目")
            st.download_button("📥 下載退件標註檔案", final_data, f"稽核報告_{up_file.name}")
            st.table(pd.DataFrame(logs))
        else:
            st.success("🎉 檢查完成！符合法規標準與校內原則。")
