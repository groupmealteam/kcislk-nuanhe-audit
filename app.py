import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# 1. 網頁基本設定
st.set_page_config(page_title="一月初 - 輕食菜單自檢系統", layout="wide")
YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

def clean_name(text):
    if pd.isna(text): return ""
    # 只提取中文字，過濾掉英文、數字與符號
    return "".join(re.findall(r'[\u4e00-\u9fa5]+', str(text)))

def audit_logic(file):
    try:
        # 強制讀取所有分頁，不設表頭
        wb = load_workbook(file)
        all_sheets = pd.read_excel(file, sheet_name=None, header=None)
    except:
        return ["❌ 檔案讀取失敗，請確認為 Excel 檔。"], None

    results = []
    output = BytesIO()
    found_menu = False

    for sheet_name, df in all_sheets.items():
        df = df.fillna("")
        ws = wb[sheet_name]
        
        # 尋找日期行 (支持 2026/3/30 或 3/30 格式)
        date_row_idx = None
        for i, row in df.iterrows():
            if any(re.search(r"(\d{4}/\d{1,2}/\d{1,2})|(\d{1,2}/\d{1,2})", str(x)) for x in row):
                date_row_idx = i
                break
        
        if date_row_idx is None: continue
        found_menu = True

        # 遍歷每一欄 (週一至週五)
        for col_idx in range(len(df.columns)):
            cell_val = str(df.iloc[date_row_idx, col_idx])
            # 確認這是一欄日期的開頭
            if not re.search(r"(\d{4}/\d{1,2}/\d{1,2})|(\d{1,2}/\d{1,2})", cell_val): continue
            
            seen_today = {} # 記錄當天核心字
            
            # 從日期行往下掃描 20 行 (確保涵蓋 A餐、B餐、湯品、水果)
            for r_offset in range(1, 25):
                r_idx = date_row_idx + r_offset
                if r_idx >= len(df): break
                
                dish_text = str(df.iloc[r_idx, col_idx]).strip()
                
                # 過濾無效字眼：太短的、純數字(熱量)、或是只有「週一/週二」
                if len(dish_text) < 2 or dish_text.isdigit() or "週" in dish_text: continue
                # 過濾掉成分說明 (通常很長且含逗號)
                if "、" in dish_text and len(dish_text) > 10: continue

                # 提取核心字眼 (例如：沙茶豬肉燴飯 -> 沙茶豬肉)
                core = clean_name(dish_text)[:4] 
                if len(core) >= 2:
                    if core in seen_today:
                        # 標註重複
                        ws.cell(row=r_idx+1, column=col_idx+1).fill = YELLOW_FILL
                        prev_r = seen_today[core]
                        ws.cell(row=prev_r+1, column=col_idx+1).fill = YELLOW_FILL
                        
                        results.append({
                            "分頁": sheet_name,
                            "日期": cell_val,
                            "項目": dish_text,
                            "問題": f"❌ 食材重複：與同日項目「{core}」衝突"
                        })
                    seen_today[core] = r_idx

    if not found_menu:
        return ["❌ 偵測失敗：找不到日期格式 (如 2026/3/30)。請檢查檔案。"], None

    wb.save(output)
    return results, output.getvalue()

# --- 介面 ---
st.title("🛡️ 一月初 (暖禾) 輕食菜單自主稽核系統")
st.markdown("---")
st.info("💡 **一月初專用版**：支援 2026/3/30 格式，自動掃描全欄位食材重複。")

up = st.file_uploader("👉 請上傳『一月初』菜單 Excel", type=["xlsx"])

if up:
    logs, final_file = audit_logic(up)
    if logs and isinstance(logs[0], str) and "❌" in logs[0]:
        st.error(logs[0])
    elif logs:
        st.error(f"🚩 偵測到 {len(logs)} 項重複！")
        st.download_button("📥 下載【一月初】標註結果", final_file, f"一月初審核_{up.name}")
        st.table(pd.DataFrame(logs))
    else:
        st.success("🎉 完美！這份菜單沒有食材重複衝突。")
