import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# 1. 網頁基本設定
st.set_page_config(page_title="輕食區(一月初) 菜單自主稽核系統", layout="wide")

# 顏色標註
RED_FILL = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")    # 嚴重違規(辣、營養不足)
YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid") # 提醒(重複性)

# 教育部法規基準
STANDARD = {"熱量": (750, 850), "全榖": 4.0, "蛋白質": 4.0, "蔬菜": 2.0}

def get_core_ingredient(text):
    """提取核心食材，例如 '沙茶豬肉燴飯' -> '豬肉'"""
    if pd.isna(text): return ""
    text = str(text).split('\n')[0] # 只看第一行
    # 優先找肉類關鍵字
    for meat in ["豬", "雞", "牛", "魚", "蝦", "蛋", "肉"]:
        if meat in text: return meat
    return "".join(re.findall(r'[\u4e00-\u9fa5]+', text))[:2] # 否則取前兩個中文字

def audit_process(file):
    wb = load_workbook(file)
    all_sheets = pd.read_excel(file, sheet_name=None, header=None)
    results = []
    output = BytesIO()

    for sheet_name, df in all_sheets.items():
        df = df.fillna("")
        ws = wb[sheet_name]
        date_row = next((i for i, row in df.iterrows() if "日期Date" in str(row[2])), None)
        if date_row is None: continue

        # 建立當週食材統計 (用於檢查同週跨日重複)
        weekly_ingredients = {} 

        for col in range(3, 8): # 週一到週五
            date_val = str(df.iloc[date_row, col]).split(" ")[0]
            day_val = str(df.iloc[date_row + 1, col])
            if "202" not in date_val: continue

            # --- 1. 營養標示稽核 (法規) ---
            for r_idx in range(date_row + 10, len(df)):
                label = str(df.iloc[r_idx, 2])
                try:
                    val_str = str(df.iloc[r_idx, col])
                    val = float(re.findall(r"\d+\.?\d*", val_str)[0])
                    if "熱量" in label and (val < STANDARD["熱量"][0] or val > STANDARD["熱量"][1]):
                        ws.cell(row=r_idx+1, column=col+1).fill = RED_FILL
                        results.append({"分頁": sheet_name, "日期": date_val, "項目": "熱量", "原因": f"❌ 不合法規：{val}"})
                    elif "豆魚蛋肉" in label and val < STANDARD["蛋白質"]:
                        ws.cell(row=r_idx+1, column=col+1).fill = RED_FILL
                        results.append({"分頁": sheet_name, "日期": date_val, "項目": "蛋白質", "原因": f"❌ 份數不足：{val}"})
                except: continue

            # --- 2. 禁辣日稽核 (週一二四) ---
            if any(d in day_val for d in ["週一", "週二", "週四"]):
                for r_idx in range(date_row + 2, date_row + 15):
                    if any(x in str(df.iloc[r_idx, 2]) for x in ["水果", "湯品", "熱量"]): continue
                    cell_content = str(df.iloc[r_idx, col])
                    if "●" in cell_content or "🌶️" in cell_content:
                        ws.cell(row=r_idx+1, column=col+1).fill = RED_FILL
                        results.append({"分頁": sheet_name, "日期": date_val, "項目": "禁辣日", "原因": f"🚫 禁止出現辣標：{cell_content}"})

            # --- 3. 重複性深度稽核 (原則九) ---
            # A. 餐道間 (A vs B)
            main_A = str(df.iloc[date_row + 3, col])
            main_B = ""
            # 尋找 B 餐主食位置
            for r_idx in range(date_row + 10, len(df)):
                if "輕食B餐" in str(df.iloc[r_idx, 2]):
                    main_B = str(df.iloc[r_idx + 1, col])
                    break
            
            core_A = get_core_ingredient(main_A)
            core_B = get_core_ingredient(main_B)

            if core_A and core_B and core_A == core_B:
                results.append({"分頁": sheet_name, "日期": date_val, "項目": "重複性", "原因": f"⚠️ A/B餐主食肉種重複({core_A})"})

            # B. 同一餐道內的食材重複 (排除湯品水果)
            today_all_items = []
            for r_idx in range(date_row + 2, date_row + 15):
                label = str(df.iloc[r_idx, 2])
                if any(x in label for x in ["水果", "湯品", "熱量", "食材內容"]): continue
                item_name = str(df.iloc[r_idx, col])
                if len(item_name) > 2:
                    core = get_core_ingredient(item_name)
                    if core in today_all_items:
                        results.append({"分頁": sheet_name, "日期": date_val, "項目": "重複性", "原因": f"⚠️ 單一餐道內食材重複({core})"})
                    today_all_items.append(core)

    wb.save(output)
    return results, output.getvalue()

# --- 介面呈現 ---
st.title("🛡️ 輕食區(一月初) 菜單自主稽核系統")
st.caption("製作者：Alison / 版本：V5 (全面稽核強化版)")
st.markdown("---")

up = st.file_uploader("👉 上傳菜單 Excel", type=["xlsx"])
if up:
    logs, file_data = audit_process(up)
    if logs:
        st.error(f"🚩 發現 {len(logs)} 項違反原則或法規之處")
        st.download_button("📥 下載退件標註檔", file_data, f"稽核報告_{up.name}")
        st.table(pd.DataFrame(logs))
    else:
        st.success("🎉 完美！通過所有法規與重複性審核。")
