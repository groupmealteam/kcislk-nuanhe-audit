import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# 1. 網頁基本設定
st.set_page_config(page_title="輕食區(一月初) 菜單自主稽核系統", layout="wide")

# --- 重新定義【高對比】視覺規範 ---
# 1. 食材重複：黃底紅字 (維持您的要求)
REPEAT_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
REPEAT_FONT = Font(color="FF0000", bold=True)

# 2. 熱量異常：淺粉色底 + 深紫色字 (柔和且清晰)
CALORIE_FILL = PatternFill(start_color="FDE9D9", end_color="FDE9D9", fill_type="solid") # 改用淺橘粉
CALORIE_FONT = Font(color="963634", bold=True)

# 3. 份數異常：淺紅色底 + 深紅色粗體 (取代深紅底，保證數字清晰)
PORTION_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid") 
PORTION_FONT = Font(color="9C0006", bold=True)

# 4. 其他異常：淺黃色底 + 黑色字
OTHER_FILL = PatternFill(start_color="FFFFE0", end_color="FFFFE0", fill_type="solid")
OTHER_FONT = Font(color="000000", bold=False)

STANDARD = {"熱量": (750, 850), "全榖": 4.0, "蛋白質": 4.0, "蔬菜": 2.0}

def clean_name(text):
    if pd.isna(text): return ""
    return "".join(re.findall(r'[\u4e00-\u9fa5]+', str(text).split('\n')[0]))

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

        for col in range(3, 8): 
            if col >= len(df.columns): break
            date_val = str(df.iloc[date_row, col]).split(" ")[0]
            day_val = str(df.iloc[date_row+1, col])
            if "202" not in date_val: continue

            # --- 1. 食材重複性 (黃底紅字) ---
            idx_A = date_row + 3
            main_A = clean_name(df.iloc[idx_A, col])
            idx_B = None
            for r_idx in range(date_row + 5, len(df)):
                if "輕食B餐" in str(df.iloc[r_idx, 2]):
                    idx_B = r_idx + 1
                    break
            
            if idx_B and idx_B < len(df):
                main_B = clean_name(df.iloc[idx_B, col])
                if main_A[:2] == main_B[:2] and len(main_A) >= 2:
                    for r in [idx_A, idx_B]:
                        cell = ws.cell(row=r+1, column=col+1)
                        cell.fill = REPEAT_FILL
                        cell.font = REPEAT_FONT
                    results.append({"日期": date_val, "項目": "食材重複", "原因": f"黃底紅字：A/B餐主食重複({main_A[:2]})"})

            # --- 2. 營養標示 (視覺強化版) ---
            for r_idx in range(date_row + 10, len(df)):
                label = str(df.iloc[r_idx, 2])
                try:
                    val = float(re.findall(r"\d+\.?\d*", str(df.iloc[r_idx, col]))[0])
                    cell = ws.cell(row=r_idx+1, column=col+1)
                    if "熱量" in label and (val < STANDARD["熱量"][0] or val > STANDARD["熱量"][1]):
                        cell.fill = CALORIE_FILL
                        cell.font = CALORIE_FONT
                        results.append({"日期": date_val, "項目": "熱量", "原因": "淺橘底：熱量異常"})
                    elif any(x in label for x in ["全榖", "豆魚蛋肉", "蔬菜"]):
                        key = "全榖" if "全榖" in label else "蛋白質" if "豆魚" in label else "蔬菜"
                        if val < STANDARD.get(key, 0):
                            cell.fill = PORTION_FILL
                            cell.font = PORTION_FONT
                            results.append({"日期": date_val, "項目": key, "原因": "淺紅底：份數不足"})
                except: continue

            # --- 3. 禁辣日 (淺黃底) ---
            if any(d in day_val for d in ["週一", "週二", "週四"]):
                for r_idx in range(date_row + 2, date_row + 15):
                    if "水果" in str(df.iloc[r_idx, 2]): continue
                    cell_val = str(df.iloc[r_idx, col])
                    if "●" in cell_val or "🌶️" in cell_val:
                        cell = ws.cell(row=r_idx+1, column=col+1)
                        cell.fill = OTHER_FILL
                        cell.font = OTHER_FONT
                        results.append({"日期": date_val, "項目": "禁辣日", "原因": "淺黃底：標示違規"})

    wb.save(output)
    return results, output.getvalue()

# --- 介面 ---
st.title("🛡️ 輕食區(一月初) 菜單自主稽核系統")
st.caption("製作者：Alison / 版本：V5.1")

st.markdown("""
### 🎨 標註說明 (高對比清晰版)
* **黃底紅字**：食材重複 (A/B 餐道主食雷同)
* **淺橘底深紅字**：熱量異常 (低於 750 或高於 850)
* **淺紅底深紅字**：份數不足 (教育部基準不達標)
* **淺黃底黑字**：其他異常 (如禁辣日標示錯誤)
""")

up = st.file_uploader("👉 上傳菜單 Excel", type=["xlsx"])
if up:
    logs, data = audit_process(up)
    if logs:
        st.error(f"🚩 偵測到 {len(logs)} 項異常項目")
        st.download_button("📥 下載退件標註檔", data, f"稽核報告_{up.name}")
        st.table(pd.DataFrame(logs))
    else:
        st.success("🎉 完美！通過所有稽核。")
