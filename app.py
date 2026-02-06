import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# 1. 網頁基本設定
st.set_page_config(page_title="一月初 - 輕食菜單稽核系統", layout="wide")
YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
RED_FILL = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")

def clean_name(text):
    if pd.isna(text): return ""
    # 提取中文核心字，過濾換行與雜訊
    return "".join(re.findall(r'[\u4e00-\u9fa5]+', str(text).split('\n')[0]))

def audit_logic(file):
    try:
        wb = load_workbook(file)
        all_sheets = pd.read_excel(file, sheet_name=None, header=None)
    except:
        return ["❌ 檔案讀取失敗，請確認格式。"], None

    results = []
    output = BytesIO()
    found_menu = False

    for sheet_name, df in all_sheets.items():
        df = df.fillna("")
        ws = wb[sheet_name]
        
        # 定位「日期Date」行 (在您的檔案中通常是第3或第4行)
        date_row = None
        for i, row in df.iterrows():
            if any("日期Date" in str(x) for x in row):
                date_row = i
                break
        
        if date_row is None: continue
        found_menu = True

        # 遍歷 C 到 G 欄 (對應原始數據中的週一至週五)
        for col_idx in range(3, 8):
            if col_idx >= len(df.columns): break
            
            date_val = str(df.iloc[date_row, col_idx]) # 例如 2026-03-30
            day_label = str(df.iloc[date_row + 1, col_idx]) # 例如 週一
            
            seen_today = {} # 記錄已出現的餐點
            
            # 一月初特色：餐點在特定的「主食」、「湯品」行
            # 我們直接掃描該欄所有「非食材內容」的格子
            for r_idx in range(date_row + 2, len(df)):
                label_cell = str(df.iloc[r_idx, 2]) # 檢查 B 欄標籤
                
                # 排除標註為「食材內容」、「熱量」、「份數」的行
                if any(x in label_cell for x in ["食材內容", "熱量", "份", "類"]): continue
                
                cell_val = str(df.iloc[r_idx, col_idx]).strip()
                if len(cell_val) < 2 or cell_val.isdigit(): continue

                # A. 禁辣日檢查 (週一、二、四)
                if any(d in day_label for d in ["週一", "週二", "週四"]):
                    if "●" in cell_val or "🌶️" in cell_val:
                        ws.cell(row=r_idx+1, column=col_idx+1).fill = RED_FILL
                        results.append({"分頁": sheet_name, "日期": date_val, "項目": cell_val, "原因": "🚫 禁辣日違規"})

                # B. 食材重複檢查 (抓核心前兩個字)
                core = clean_name(cell_val)[:2]
                if len(core) >= 2:
                    if core in seen_today:
                        ws.cell(row=r_idx+1, column=col_idx+1).fill = YELLOW_FILL
                        prev_r = seen_today[core]
                        ws.cell(row=prev_r+1, column=col_idx+1).fill = YELLOW_FILL
                        results.append({"分頁": sheet_name, "日期": date_val, "項目": cell_val, "原因": f"❌ 食材重複({core})"})
                    seen_today[core] = r_idx

    if not found_menu:
        return ["❌ 偵測失敗：格式不符（未偵測到日期Date標籤）。"], None

    wb.save(output)
    return results, output.getvalue()

# --- 網頁介面 ---
st.title("🛡️ 一月初 (暖禾) 輕食菜單稽核系統")
st.warning("⚠️ 系統已依據您上傳的原始格式完成『座標校對』，現在可以精準抓出重複餐點。")

up = st.file_uploader("👉 上傳『一月初』Excel 檔案", type=["xlsx"])

if up:
    logs, final_file = audit_logic(up)
    if logs and isinstance(logs[0], str) and "❌" in logs[0]:
        st.error(logs[0])
    elif logs:
        st.error(f"🚩 偵測到 {len(logs)} 項衝突項目！")
        st.download_button("📥 下載標註檔案", final_file, f"一月初審核_{up.name}")
        st.table(pd.DataFrame(logs))
    else:
        st.success("🎉 完美！該週菜單符合所有審閱原則。")
