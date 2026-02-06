import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# 1. 網頁基本設定
st.set_page_config(page_title="輕食區(一月初) 菜單自主稽核系統", layout="wide")

# 顏色標註設定
RED_FILL = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")    # 違規(辣、營養不足)
YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid") # 提醒(重複性)

# 教育部國高中營養基準 (法規標準)
STANDARD = {
    "熱量": (750, 850),
    "全榖": 4.0,
    "蛋白質": 4.0,
    "蔬菜": 2.0
}

def clean_name(text):
    if pd.isna(text): return ""
    # 僅提取第一行中文字，移除英文描述與雜訊
    return "".join(re.findall(r'[\u4e00-\u9fa5]+', str(text).split('\n')[0]))

def to_float(val):
    try:
        nums = re.findall(r"\d+\.?\d*", str(val))
        return float(nums[0]) if nums else 0.0
    except:
        return 0.0

def audit_process(file):
    try:
        wb = load_workbook(file)
        all_sheets = pd.read_excel(file, sheet_name=None, header=None)
    except Exception as e:
        return [f"❌ 檔案開啟失敗：{str(e)}"], None

    results = []
    output = BytesIO()
    found_any_valid_sheet = False

    for sheet_name, df in all_sheets.items():
        df = df.fillna("")
        ws = wb[sheet_name]
        
        # 定位「日期Date」所在行
        date_row = next((i for i, row in df.iterrows() if "日期Date" in str(row[2])), None)
        if date_row is None: continue
        found_any_valid_sheet = True

        # 遍歷週一到週五 (D欄到H欄)
        for col in range(3, 8):
            if col >= len(df.columns): break
            
            date_val = str(df.iloc[date_row, col]).split(" ")[0]
            day_val = str(df.iloc[date_row + 1, col])
            if "202" not in date_val: continue

            # --- 1. 營養成分稽核 (法規標準) ---
            for r_idx in range(date_row + 10, len(df)):
                label = str(df.iloc[r_idx, 2])
                val = to_float(df.iloc[r_idx, col])
                
                if "熱量" in label:
                    if val < STANDARD["熱量"][0] or val > STANDARD["熱量"][1]:
                        ws.cell(row=r_idx+1, column=col+1).fill = RED_FILL
                        results.append({"分頁": sheet_name, "日期": date_val, "項目": "熱量", "原因": f"❌ 不合法規：{val} (需在750-850間)"})
                elif "全榖" in label and val < STANDARD["全榖"]:
                    ws.cell(row=r_idx+1, column=col+1).fill = RED_FILL
                    results.append({"分頁": sheet_name, "日期": date_val, "項目": "全榖", "原因": f"❌ 份數不足：{val} (要求 {STANDARD['全榖']})"})
                elif "豆魚蛋肉" in label and val < STANDARD["蛋白質"]:
                    ws.cell(row=r_idx+1, column=col+1).fill = RED_FILL
                    results.append({"分頁": sheet_name, "日期": date_val, "項目": "蛋白質", "原因": f"❌ 份數不足：{val} (要求 {STANDARD['蛋白質']})"})
                elif "蔬菜" in label and val < STANDARD["蔬菜"]:
                    ws.cell(row=r_idx+1, column=col+1).fill = RED_FILL
                    results.append({"分頁": sheet_name, "日期": date_val, "項目": "蔬菜", "原因": f"❌ 份數不足：{val} (要求 {STANDARD['蔬菜']})"})

            # --- 2. 禁辣日審核 (原則四：週一二四) ---
            if any(d in day_val for d in ["週一", "週二", "週四"]):
                for r_idx in range(date_row + 2, date_row + 12):
                    if any(x in str(df.iloc[r_idx, 2]) for x in ["水果", "湯品", "熱量"]): continue
                    cell_val = str(df.iloc[r_idx, col])
                    if "●" in cell_val or "🌶️" in cell_val:
                        ws.cell(row=r_idx+1, column=col+1).fill = RED_FILL
                        results.append({"分頁": sheet_name, "日期": date_val, "項目": "禁辣日", "原因": f"🚫 禁止辣標：{cell_val}"})

            # --- 3. 食材重複性審核 (原則九) ---
            # A. 餐道間 (A餐 vs B餐) 重複比對
            main_A = clean_name(df.iloc[date_row + 3, col])
            main_B = ""
            for r_idx in range(date_row + 12, len(df)):
                if "輕食B餐" in str(df.iloc[r_idx, 2]):
                    main_B = clean_name(df.iloc[r_idx + 1, col])
                    break
            
            if main_A and main_B:
                core_A, core_B = main_A[:2], main_B[:2]
                if core_A == core_B:
                    # 標註這兩格為黃色
                    ws.cell(row=date_row + 4, column=col+1).fill = YELLOW_FILL
                    results.append({"分頁": sheet_name, "日期": date_val, "項目": "重複性", "原因": f"⚠️ A/B餐主食雷同({core_A})"})

            # B. 排除重複性的白名單 (水果、湯品、食材內容)
            # 這裡維持您之前要求的：不抓「水果」與「湯品」
            seen_today = {}
            for r_idx in range(date_row + 2, date_row + 20):
                row_label = str(df.iloc[r_idx, 2])
                if any(x in row_label for x in ["水果", "湯品", "食材內容", "熱量", "份數"]): continue
                
                item_name = clean_name(df.iloc[r_idx, col])
                if len(item_name) >= 2:
                    core = item_name[:2]
                    if core in seen_today:
                        ws.cell(row=r_idx+1, column=col+1).fill = YELLOW_FILL
                        results.append({"分頁": sheet_name, "日期": date_val, "項目": "重複性", "原因": f"⚠️ 食材重複({core})"})
                    seen_today[core] = r_idx

    wb.save(output)
    return results, output.getvalue()

# --- 介面呈現 ---
st.title("🛡️ 輕食區(一月初) 菜單自主稽核系統")
st.caption("製作者：Alison / 版本：V5 (全功能整合版)")
st.markdown("---")

up_file = st.file_uploader("👉 上傳週菜單 Excel (.xlsx)", type=["xlsx"])
if up_file:
    with st.spinner("深度稽核中，請稍候..."):
        logs, final_data = audit_process(up_file)
        if logs:
            st.error(f"🚩 發現 {len(logs)} 項違反原則或法規之處")
            st.download_button("📥 下載退件標註檔", final_data, f"稽核報告_{up_file.name}")
            st.table(pd.DataFrame(logs))
        else:
            st.success("🎉 完美！通過所有法規與重複性審核。")
