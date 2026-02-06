import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# 1. 網頁基本設定
st.set_page_config(page_title="輕食區(一月初) 菜單自主稽核系統", layout="wide")

# --- 定義顏色規範 ---
# 1. 食材重複：黃底紅字
REPEAT_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
REPEAT_FONT = Font(color="FF0000", bold=True)

# 2. 熱量異常：粉色底
CALORIE_FILL = PatternFill(start_color="FFCCFF", end_color="FFCCFF", fill_type="solid")

# 3. 份數異常：紅色底
PORTION_FILL = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
WHITE_FONT = Font(color="FFFFFF") # 紅底配白字較清晰

# 4. 其他異常 (如禁辣日)：淺黃底
OTHER_FILL = PatternFill(start_color="FFFFE0", end_color="FFFFE0", fill_type="solid")

# 教育部國高中營養基準
STANDARD = {"熱量": (750, 850), "全榖": 4.0, "蛋白質": 4.0, "蔬菜": 2.0}

def clean_name(text):
    if pd.isna(text): return ""
    return "".join(re.findall(r'[\u4e00-\u9fa5]+', str(text).split('\n')[0]))

def to_float(val):
    try:
        nums = re.findall(r"\d+\.?\d*", str(val))
        return float(nums[0]) if nums else 0.0
    except: return 0.0

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
            day_val = str(df.iloc[date_row + 1, col])
            if "202" not in date_val: continue

            # --- A. 營養標示稽核 (標色規範) ---
            for r_idx in range(date_row + 10, len(df)):
                label = str(df.iloc[r_idx, 2])
                val = to_float(df.iloc[r_idx, col])
                
                if "熱量" in label:
                    if val < STANDARD["熱量"][0] or val > STANDARD["熱量"][1]:
                        ws.cell(row=r_idx+1, column=col+1).fill = CALORIE_FILL # 粉色底
                        results.append({"日期": date_val, "項目": "熱量", "原因": f"粉底：熱量異常({val})"})
                elif any(x in label for x in ["全榖", "豆魚蛋肉", "蔬菜"]):
                    key = "全榖" if "全榖" in label else "蛋白質" if "豆魚" in label else "蔬菜"
                    if val < STANDARD.get(key, 0):
                        cell = ws.cell(row=r_idx+1, column=col+1)
                        cell.fill = PORTION_FILL # 紅色底
                        cell.font = WHITE_FONT
                        results.append({"日期": date_val, "項目": key, "原因": f"紅底：份數不足({val})"})

            # --- B. 禁辣日審核 (淺黃底) ---
            if any(d in day_val for d in ["週一", "週二", "週四"]):
                for r_idx in range(date_row + 2, date_row + 12):
                    if "水果" in str(df.iloc[r_idx, 2]): continue # 忽略水果
                    cell_val = str(df.iloc[r_idx, col])
                    if "●" in cell_val or "🌶️" in cell_val:
                        ws.cell(row=r_idx+1, column=col+1).fill = OTHER_FILL # 淺黃底
                        results.append({"日期": date_val, "項目": "禁辣日", "原因": "淺黃底：禁辣日違規"})

            # --- C. 食材重複性 (黃底紅字) ---
            # 1. 餐道間 A vs B
            main_A = clean_name(df.iloc[date_row + 3, col])
            main_B = ""
            for r_idx in range(date_row + 10, len(df)):
                if "輕食B餐" in str(df.iloc[r_idx, 2]):
                    main_B = clean_name(df.iloc[r_idx + 1, col])
                    break
            
            if main_A and main_B and main_A[:2] == main_B[:2]:
                cell_a = ws.cell(row=date_row + 4, column=col+1)
                cell_a.fill = REPEAT_FILL
                cell_a.font = REPEAT_FONT
                results.append({"日期": date_val, "項目": "重複性", "原因": f"黃底紅字：A/B餐主食重複({main_A[:2]})"})

    wb.save(output)
    return results, output.getvalue()

# --- 介面呈現 ---
st.title("🛡️ 輕食區(一月初) 菜單自主稽核系統")
st.caption("製作者：Alison / 版本：V5 (顏色規範版)")
st.markdown("""
| 異常類型 | 標註樣式 | 說明 |
| :--- | :--- | :--- |
| **食材重複** | <span style='color:red; background:yellow'>黃底紅字</span> | A/B 餐道主食雷同 |
| **熱量異常** | <span style='background:#FFCCFF'>粉色底</span> | 低於 750 或高於 850 Kcal |
| **份數異常** | <span style='color:white; background:red'>紅色底</span> | 低於教育部份數基準 |
| **其他異常** | <span style='background:#FFFFE0'>淺黃底</span> | 禁辣日違規等 |
""", unsafe_allow_html=True)

up = st.file_uploader("👉 上傳一月初 Excel (.xlsx)", type=["xlsx"])
if up:
    logs, data = audit_process(up)
    if logs:
        st.error(f"🚩 偵測到 {len(logs)} 項異常項目")
        st.download_button("📥 下載標註檔案", data, f"稽核報告_{up.name}")
        st.table(pd.DataFrame(logs))
    else:
        st.success("🎉 通過所有稽核！")
