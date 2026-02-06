import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# 1. 網頁基本設定
st.set_page_config(page_title="康橋校園菜單 - 法規稽核系統", layout="wide")
RED_FILL = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid") # 違規
YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid") # 提醒

# 教育部法規基準 (國高中標準)
STANDARD = {
    "熱量": (750, 850),
    "全榖": 4.0,
    "蛋白質": 4.0,
    "蔬菜": 2.0
}

def clean_text(val):
    if pd.isna(val): return ""
    return str(val).split('\n')[0].strip()

def audit_process(file):
    try:
        wb = load_workbook(file)
        # 一月初通常有多個分頁，全部掃描
        all_sheets = pd.read_excel(file, sheet_name=None, header=None)
    except:
        return ["❌ 檔案讀取失敗，請確認為 Excel 檔"], None

    results = []
    output = BytesIO()

    for sheet_name, df in all_sheets.items():
        df = df.fillna("")
        ws = wb[sheet_name]
        
        # 1. 座標定位：尋找「日期Date」所在的行
        date_row_idx = None
        for i, row in df.iterrows():
            if "日期Date" in str(row[2]):
                date_row_idx = i
                break
        
        if date_row_idx is None: continue

        # 2. 橫向掃描週一到週五 (D 欄到 H 欄，Index 3 到 7)
        for col_idx in range(3, 8):
            if col_idx >= len(df.columns): break
            
            date_val = str(df.iloc[date_row_idx, col_idx])
            day_val = str(df.iloc[date_row_idx + 1, col_idx]) # 週幾
            if not date_val or "202" not in date_val: continue

            # --- A. 營養標示審核 (對標教育部法規) ---
            # 根據 V5 結構，數據在日期下方約 15-20 行
            nutrition_mapping = {
                "熱量": "熱量",
                "全榖": "全榖雜糧類",
                "蛋白質": "豆魚蛋肉類",
                "蔬菜": "蔬菜類"
            }
            
            for row_idx in range(date_row_idx + 10, len(df)):
                row_label = str(df.iloc[row_idx, 2]) # B 欄的標籤
                
                for key, label in nutrition_mapping.items():
                    if label in row_label:
                        try:
                            raw_val = str(df.iloc[row_idx, col_idx])
                            val = float(re.findall(r"\d+\.?\d*", raw_val)[0])
                            
                            # 熱量稽核 (750-850)
                            if key == "熱量":
                                if val < STANDARD["熱量"][0] or val > STANDARD["熱量"][1]:
                                    ws.cell(row=row_idx+1, column=col_idx+1).fill = RED_FILL
                                    results.append({"日期": date_val, "類別": "熱量", "原因": f"❌ 不合法規：{val} Kcal (應在750-850)"})
                            # 份數稽核 (低於 4 份)
                            elif key in STANDARD and val < STANDARD[key]:
                                ws.cell(row=row_idx+1, column=col_idx+1).fill = RED_FILL
                                results.append({"日期": date_val, "類別": key, "原因": f"❌ 份數不足：{val} 份 (要求 {STANDARD[key]})"})
                        except:
                            pass # 避免非數字導致崩潰

            # --- B. 禁辣日審核 (週一、二、四) ---
            if any(d in day_val for d in ["週一", "週二", "週四"]):
                for row_idx in range(date_row_idx + 2, date_row_idx + 15):
                    cell_content = str(df.iloc[row_idx, col_idx])
                    if "●" in cell_content or "🌶️" in cell_content:
                        ws.cell(row=row_idx+1, column=col_idx+1).fill = RED_FILL
                        results.append({"日期": date_val, "類別": "禁辣日", "原因": f"❌ 違規標註：{cell_content}"})

            # --- C. A/B 餐道重複性比對 (原則九) ---
            # 這裡我們只抓「主食」這兩行來比對肉種
            main_dish_A = clean_text(df.iloc[date_row_idx + 3, col_idx])
            # B餐主食在下方約 15 行處
            main_dish_B = ""
            for r_idx in range(date_row_idx + 10, len(df)):
                if "輕食B餐" in str(df.iloc[r_idx, 2]):
                    main_dish_B = clean_text(df.iloc[r_idx + 1, col_idx])
                    break
            
            if main_dish_A and main_dish_B:
                core_A, core_B = main_dish_A[:2], main_dish_B[:2]
                if core_A == core_B:
                    results.append({"日期": date_val, "類別": "食材重複", "原因": f"⚠️ A/B餐主食重複({core_A})，建議修正"})

    wb.save(output)
    return results, output.getvalue()

# --- 介面呈現 ---
st.title("🛡️ 康橋行政專用：菜單法規稽核系統 (V5 修正版)")
st.info("本系統已對標『教育部營養基準』：熱量 750-850 / 蛋白質 4 份 / 全榖 4 份。")

up = st.file_uploader("👉 請上傳『一月初』Excel 菜單", type=["xlsx"])
if up:
    with st.spinner("正在進行法規深度稽核..."):
        logs, final_file = audit_process(up)
        if logs:
            st.error(f"🚩 發現 {len(logs)} 項不符法規或原則之項目！")
            st.download_button("📥 下載退件標註檔", final_file, f"退件報告_{up.name}")
            st.table(pd.DataFrame(logs))
        else:
            st.success("🎉 完美！該菜單符合法規營養標準與校內原則。")
