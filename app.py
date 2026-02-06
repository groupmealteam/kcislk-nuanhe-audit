import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# 1. 網頁基本設定 (行政專用最高權限)
st.set_page_config(page_title="林口康橋 - 校方行政稽核系統", layout="wide")

# 設定顏色標註
RED_FILL = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")    # 嚴重違規(禁辣/規格)
YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid") # 食材重複

# 合約規格庫 (新北食品)
CONTRACT_SPECS = {"現撈小卷": "80|100", "無刺白帶魚": "120|150", "手作漢堡排": "150", "手作烤肉串": "80"}

def clean_name(text):
    if pd.isna(text) or re.search(r"\d{1,2}/\d{1,2}", str(text)): return ""
    return "".join(re.findall(r'[\u4e00-\u9fa5]+', str(text)))

def admin_audit(file, mode):
    wb = load_workbook(file)
    all_sheets = pd.read_excel(file, sheet_name=None, header=None)
    results = []
    output = BytesIO()
    is_valid = False

    for sheet_name, df in all_sheets.items():
        df = df.fillna("")
        ws = wb[sheet_name]
        
        # 定位行座標
        date_row = next((i for i, row in df.iterrows() if any(re.search(r"\d{1,2}/\d{1,2}", str(c)) for c in row)), None)
        target_rows = [i for i, row in df.iterrows() if any(k in str(row[1]) for k in ["主食", "副菜", "主菜", "套餐", "麵食"])]
        
        if date_row is None or not target_rows: continue
        is_valid = True

        for col in range(2, len(df.columns)):
            date_val = str(df.iloc[date_row, col])
            day_val = str(df.iloc[date_row+1, col]) if (date_row+1) < len(df) else ""
            if not re.search(r"\d{1,2}/\d{1,2}", date_val): continue
            
            # --- 行政專用審核邏輯 ---
            is_restricted_day = any(d in day_val for d in ["週一", "週二", "週四"])
            seen_today = {}
            processed_cnt = 0
            fried_cnt = 0

            for r_idx in target_rows:
                cell_val = str(df.iloc[r_idx, col]).strip()
                if not cell_val or len(cell_val) < 2: continue

                # A. 禁辣日審核 (原則四)
                if is_restricted_day and ("●" in cell_val or "🌶️" in cell_val):
                    ws.cell(row=r_idx+1, column=col+1).fill = RED_FILL
                    results.append({"日期": date_val, "項目": cell_val, "原因": "🚫 禁辣日違規標示"})

                # B. 食材重複審核 (原則九)
                core = clean_name(cell_val)[:2]
                if len(core) >= 2:
                    if core in seen_today:
                        ws.cell(row=r_idx+1, column=col+1).fill = YELLOW_FILL
                        ws.cell(row=seen_today[core]+1, column=col+1).fill = YELLOW_FILL
                        results.append({"日期": date_val, "項目": cell_val, "原因": f"❌ 食材重複({core})"})
                    seen_today[core] = r_idx

                # C. 法規標示計數 (原則五/七)
                if "△" in cell_val: processed_cnt += 1
                if "◎" in cell_val: fried_cnt += 1

                # D. 新北食品專屬規格檢查
                for item, spec in CONTRACT_SPECS.items():
                    if item in cell_val and not re.search(spec, cell_val):
                        ws.cell(row=r_idx+1, column=col+1).fill = RED_FILL
                        results.append({"日期": date_val, "項目": cell_val, "原因": f"⚠️ 規格應為 {spec}g"})

            # E. 總結判斷
            if processed_cnt > 1: results.append({"日期": date_val, "原因": "🚫 加工品(△)超量"})
            if fried_cnt > 1: results.append({"日期": date_val, "原因": "🚫 油炸(◎)超量"})

    wb.save(output)
    return results, output.getvalue() if is_valid else (None, None)

# --- 行政介面 ---
st.title("🏫 林口康橋：校方行政專用稽核系統")
st.markdown("---")
st.sidebar.header("⚙️ 稽核權限設定")
mode = st.sidebar.selectbox("切換審核對象", ["新北食品 (NTCatering)", "一月初 (暖禾)"])

up = st.file_uploader(f"👉 請上傳【{mode}】待審核菜單", type=["xlsx"])

if up:
    logs, excel = admin_audit(up, mode)
    if logs:
        st.error(f"🚩 偵測到 {len(logs)} 項實質違規，建議退回修正。")
        st.download_button("📥 下載行政標註檔 (退件回傳用)", excel, f"校方複審_{up.name}")
        st.table(pd.DataFrame(logs))
    else:
        st.success("🎉 完美！該週菜單符合所有校內審閱原則。")
