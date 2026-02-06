import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# 1. 網頁基本設定
st.set_page_config(page_title="一月初 - 輕食菜單自檢系統", layout="wide")
YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
RED_FILL = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")

def clean_name(text):
    if pd.isna(text): return ""
    # 僅提取中文字，過濾掉英文、標點符號與換行
    return "".join(re.findall(r'[\u4e00-\u9fa5]+', str(text)))

def audit_logic(file):
    try:
        wb = load_workbook(file)
        all_sheets = pd.read_excel(file, sheet_name=None, header=None)
    except:
        return ["❌ 檔案讀取失敗，請確認為 Excel (.xlsx) 格式。"], None

    results = []
    output = BytesIO()
    found_menu = False

    for sheet_name, df in all_sheets.items():
        df = df.fillna("")
        ws = wb[sheet_name]
        
        # 尋找「日期Date」所在行 (對應原始數據)
        date_row_idx = None
        for i, row in df.iterrows():
            if "日期Date" in str(row[2]): # 鎖定 C 欄附近的關鍵字
                date_row_idx = i
                break
        
        if date_row_idx is None: continue
        found_menu = True

        # 掃描週一到週五 (C 欄到 G 欄，即 index 3 到 7)
        for col_idx in range(3, 8):
            date_val = str(df.iloc[date_row_idx, col_idx]) # 抓取如 2026-03-30
            day_label = str(df.iloc[date_row_idx + 1, col_idx]) # 抓取週幾
            
            seen_today = {} # 記錄食材
            
            # 從「主食」行開始向下掃描至「水果」行為止 (約 10-15 行)
            for r_idx in range(date_row_idx + 3, date_row_idx + 15):
                if r_idx >= len(df): break
                
                cell_val = str(df.iloc[r_idx, col_idx]).strip()
                # 排除空值、過短字、或純熱量數字
                if len(cell_val) < 2 or cell_val.replace('.','').isdigit(): continue
                # 排除「食材內容」描述列
                if "、" in cell_val: continue

                # A. 禁辣日稽核 (原則四：週一二四)
                if any(d in day_label for d in ["週一", "週二", "週四"]):
                    if "●" in cell_val or "🌶️" in cell_val:
                        ws.cell(row=r_idx+1, column=col_idx+1).fill = RED_FILL
                        results.append({"分頁": sheet_name, "日期": date_val, "項目": cell_val, "原因": "🚫 禁辣日標示違規"})

                # B. 食材重複稽核 (原則九)
                core = clean_name(cell_val)[:2] # 抓取前兩個中文字
                if len(core) >= 2:
                    if core in seen_today:
                        ws.cell(row=r_idx+1, column=col_idx+1).fill = YELLOW_FILL
                        prev_r = seen_today[core]
                        ws.cell(row=prev_r+1, column=col_idx+1).fill = YELLOW_FILL
                        results.append({"分頁": sheet_name, "日期": date_val, "項目": cell_val, "原因": f"❌ 食材重複({core})"})
                    seen_today[core] = r_idx

    if not found_menu:
        return ["❌ 偵測失敗：找不到『日期Date』關鍵欄位，請檢查菜單格式。"], None

    wb.save(output)
    return results, output.getvalue()

# --- 介面呈現 ---
st.title("🛡️ 一月初 (暖禾) 輕食菜單自主稽核系統")
st.markdown("---")
st.info("💡 **一月初專用版**：已針對 2026-03-30 格式優化，自動檢查禁辣日與食材重複。")

up = st.file_uploader("👉 請上傳『一月初』週菜單 Excel", type=["xlsx"])

if up:
    with st.spinner("正在依照康橋審閱原則分析菜單..."):
        logs, final_file = audit_logic(up)
        if logs and isinstance(logs[0], str) and "❌" in logs[0]:
            st.error(logs[0])
        elif logs:
            st.error(f"🚩 偵測到 {len(logs)} 項違規，請下載標註檔：")
            st.download_button("📥 下載一月初稽核結果", final_file, f"一月初審核_{up.name}")
            st.table(pd.DataFrame(logs))
        else:
            st.success("🎉 完美！這份一月初菜單符合所有審閱原則。")
