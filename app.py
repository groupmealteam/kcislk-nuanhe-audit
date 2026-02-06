import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# 1. 網頁基本設定
st.set_page_config(page_title="林口康橋 - 行政終極稽核系統", layout="wide")

# 設定顏色標註
RED_FILL = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")    # 嚴重違規
YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid") # 食材重複

# 合約規格庫 (新北食品專屬)
XINBEI_SPECS = {"現撈小卷": "80|100", "無刺白帶魚": "120|150", "手作漢堡排": "150", "手作烤肉串": "80"}

def audit_process(file, vendor_mode):
    try:
        wb = load_workbook(file)
        all_sheets = pd.read_excel(file, sheet_name=None, header=None)
    except:
        return ["❌ 檔案讀取失敗，請確認是否為 Excel 檔。"], None

    results = []
    output = BytesIO()
    found_any_valid_sheet = False

    for sheet_name, df in all_sheets.items():
        df = df.fillna("")
        ws = wb[sheet_name]
        
        # 關鍵行定位
        date_row = next((i for i, row in df.iterrows() if any(re.search(r"\d{1,2}/\d{1,2}", str(c)) for c in row)), None)
        # 根據模式切換檢查關鍵字
        if vendor_mode == "新北食品 (NTCatering)":
            target_keywords = ["主食", "副菜", "主菜"]
        else: # 一月初 (暖禾)
            target_keywords = ["套餐", "麵食", "輕食", "副食"]
            
        target_rows = [i for i, row in df.iterrows() if any(k in str(row[1]) for k in target_keywords)]
        
        if date_row is None or not target_rows:
            continue
            
        found_any_valid_sheet = True

        for col in range(2, len(df.columns)):
            date_val = str(df.iloc[date_row, col])
            if not re.search(r"\d{1,2}/\d{1,2}", date_val): continue
            
            # --- 核心稽核邏輯 ---
            seen_today = {}
            processed_cnt = 0
            fried_cnt = 0

            for r_idx in target_rows:
                cell_val = str(df.iloc[r_idx, col]).strip()
                if not cell_val or len(cell_val) < 2: continue

                # 1. 共通規則：禁辣日 (週一二四)
                day_val = str(df.iloc[date_row+1, col]) if (date_row+1) < len(df) else ""
                if any(d in day_val for d in ["週一", "週二", "週四"]):
                    if "●" in cell_val or "🌶️" in cell_val:
                        ws.cell(row=r_idx+1, column=col+1).fill = RED_FILL
                        results.append({"日期": date_val, "原因": f"🚫 禁辣日違規標示：{cell_val}"})

                # 2. 一月初 (暖禾) 專屬：食材重複比對 (原則九)
                if vendor_mode == "一月初 (暖禾)":
                    core = "".join(re.findall(r'[\u4e00-\u9fa5]+', cell_val))[:2]
                    if len(core) >= 2:
                        if core in seen_today:
                            ws.cell(row=r_idx+1, column=col+1).fill = YELLOW_FILL
                            ws.cell(row=seen_today[core]+1, column=col+1).fill = YELLOW_FILL
                            results.append({"日期": date_val, "原因": f"❌ 食材重複：{core}"})
                        seen_today[core] = r_idx

                # 3. 新北食品 專屬：規格與標籤檢查
                if vendor_mode == "新北食品 (NTCatering)":
                    for item, spec in XINBEI_SPECS.items():
                        if item in cell_val and not re.search(spec, cell_val):
                            ws.cell(row=r_idx+1, column=col+1).fill = RED_FILL
                            results.append({"日期": date_val, "原因": f"⚠️ 規格應為 {spec}g"})
                    if "△" in cell_val: processed_cnt += 1
                    if "◎" in cell_val: fried_cnt += 1

            if vendor_mode == "新北食品 (NTCatering)":
                if processed_cnt > 1: results.append({"日期": date_val, "原因": "🚫 加工品(△)單日超標"})
                if fried_cnt > 1: results.append({"日期": date_val, "原因": "🚫 油炸(◎)單日超標"})

    if not found_any_valid_sheet:
        return [f"❌ 偵測失敗：您選擇的是『{vendor_mode}』，但上傳檔案格式不符。"], None

    wb.save(output)
    return results, output.getvalue()

# --- 介面設計 ---
st.title("🏫 林口康橋：校方行政專用稽核系統")
st.markdown("---")

# 強制要求使用者先選擇廠商
vendor_choice = st.radio("第一步：請選擇本次要審核的廠商對象", ["新北食品 (NTCatering)", "一月初 (暖禾)"], horizontal=True)

st.write(f"### 第二步：上傳【{vendor_choice}】的週菜單")
up_file = st.file_uploader("請選擇 Excel 檔案 (.xlsx)", type=["xlsx"])

if up_file:
    with st.spinner("正在執行深度稽核..."):
        logs, final_data = audit_process(up_file, vendor_choice)
        
        if logs and isinstance(logs[0], str) and "❌" in logs[0]:
            st.error(logs[0])
        elif logs:
            st.error(f"🚩 發現 {len(logs)} 項違規！請下載標註檔回傳廠商修正。")
            st.download_button(f"📥 下載【{vendor_choice}】稽核標註檔", final_data, f"行政覆核_{up_file.name}")
            st.table(pd.DataFrame(logs))
        else:
            st.success(f"🎉 驗證通過！該週菜單符合【{vendor_choice}】所有合約原則。")
