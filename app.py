import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# 1. 網頁基本設定
st.set_page_config(page_title="輕食區(一月初) 菜單自主稽核系統", layout="wide")

# 顏色標註設定
RED_FILL = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")    # 違規標色
YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid") # 提醒標色

# 教育部國高中營養基準
STANDARD = {
    "熱量": (750, 850),
    "全榖": 4.0,
    "蛋白質": 4.0,
    "蔬菜": 2.0
}

def clean_dish_name(text):
    if pd.isna(text): return ""
    # 僅提取第一行中文字，移除英文描述
    return "".join(re.findall(r'[\u4e00-\u9fa5]+', str(text).split('\n')[0]))

def audit_process(file):
    try:
        wb = load_workbook(file)
        all_sheets = pd.read_excel(file, sheet_name=None, header=None)
    except Exception as e:
        return [f"❌ 檔案讀取失敗：{str(e)}"], None

    results = []
    output = BytesIO()
    found_any_valid_sheet = False

    for sheet_name, df in all_sheets.items():
        df = df.fillna("")
        ws = wb[sheet_name]
        
        # 尋找「日期Date」所在行 (V5版通常在 index 2)
        date_row = next((i for i, row in df.iterrows() if "日期Date" in str(row[2])), None)
        if date_row is None: continue
        found_any_valid_sheet = True

        # 遍歷週一到週五 (D欄到H欄, Index 3 到 7)
        for col in range(3, 8):
            if col >= len(df.columns): break
            
            date_val = str(df.iloc[date_row, col]).split(" ")[0]
            day_val = str(df.iloc[date_row + 1, col])
            if not date_val or "202" not in date_val: continue

            # --- A. 營養成分稽核 (教育部法規) ---
            # 根據 V5 結構搜尋下方標籤
            today_protein = 0
            for r_idx in range(date_row + 10, len(df)):
                label = str(df.iloc[r_idx, 2])
                try:
                    raw_val = str(df.iloc[r_idx, col])
                    val = float(re.findall(r"\d+\.?\d*", raw_val)[0])
                    
                    if "熱量" in label:
                        if val < STANDARD["熱量"][0] or val > STANDARD["熱量"][1]:
                            ws.cell(row=r_idx+1, column=col+1).fill = RED_FILL
                            results.append({"分頁": sheet_name, "日期": date_val, "項目": "熱量", "原因": f"❌ 不合法規：{val} Kcal (應在750-850)"})
                    elif "全榖" in label and val < STANDARD["全榖"]:
                        ws.cell(row=r_idx+1, column=col+1).fill = RED_FILL
                        results.append({"分頁": sheet_name, "日期": date_val, "項目": "全榖雜糧", "原因": f"❌ 份數不足：{val} 份 (要求 {STANDARD['全榖']})"})
                    elif "豆魚蛋肉" in label and val < STANDARD["蛋白質"]:
                        ws.cell(row=r_idx+1, column=col+1).fill = RED_FILL
                        results.append({"分頁": sheet_name, "日期": date_val, "項目": "蛋白質", "原因": f"❌ 份數不足：{val} 份 (要求 {STANDARD['蛋白質']})"})
                    elif "蔬菜" in label and val < STANDARD["蔬菜"]:
                        ws.cell(row=r_idx+1, column=col+1).fill = RED_FILL
                        results.append({"分頁": sheet_name, "日期": date_val, "項目": "蔬菜類", "原因": f"❌ 份數不足：{val} 份 (要求 {STANDARD['蔬菜']})"})
                except: continue

            # --- B. 禁辣日審核 (原則四：週一、二、四) ---
            if any(d in day_val for d in ["週一", "週二", "週四"]):
                # 掃描當天主食與副菜區 (不含湯品與水果)
                for r_idx in range(date_row + 2, date_row + 10):
                    cell_val = str(df.iloc[r_idx, col])
                    if "●" in cell_val or "🌶️" in cell_val:
                        ws.cell(row=r_idx+1, column=col+1).fill = RED_FILL
                        results.append({"分頁": sheet_name, "日期": date_val, "項目": "禁辣日違規", "原因": f"🚫 禁止標註：{cell_val}"})

            # --- C. 餐道間重複性比對 (原則九：A餐 vs B餐) ---
            # 找到輕食A與輕食B的主食位置進行比對
            main_A = clean_dish_name(df.iloc[date_row + 3, col])
            main_B = ""
            for r_idx in range(date_row + 15, len(df)):
                if "輕食B餐" in str(df.iloc[r_idx, 2]):
                    main_B = clean_dish_name(df.iloc[r_idx + 1, col])
                    break
            
            if main_A and main_B:
                core_A, core_B = main_A[:2], main_B[:2]
                if core_A == core_B:
                    results.append({"分頁": sheet_name, "日期
