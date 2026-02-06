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
    # 移除符號，只比對中文字核心
    return "".join(re.findall(r'[\u4e00-\u9fa5]+', str(text)))

def audit_logic(file):
    try:
        wb = load_workbook(file)
        # 一月初格式通常在第一個 Sheet，強行讀取前 15 欄
        df = pd.read_excel(file, header=None).fillna("")
    except:
        return ["❌ 檔案讀取失敗"], None

    results = []
    output = BytesIO()
    ws = wb.active # 直接針對當前工作表

    # 座標偵測：尋找含有日期 (M/D) 的那一行
    date_row_idx = None
    for i, row in df.iterrows():
        if any(re.search(r"\d{1,2}/\d{1,2}", str(x)) for x in row):
            date_row_idx = i
            break
    
    if date_row_idx is None:
        return ["❌ 偵測失敗：找不到日期欄位（例如 3/10），請確認檔案格式。"], None

    # 定義要檢查的「餐點內容行」（一月初格式：通常在日期下方 2-10 行內）
    # 我們擴大搜尋範圍，只要是有文字的格子都列入比對
    content_start = date_row_idx + 1
    content_end = min(date_row_idx + 15, len(df))

    # 遍歷週一到週五 (通常是 C 欄到 G 欄，即 index 2 到 6)
    for col_idx in range(2, len(df.columns)):
        date_cell = str(df.iloc[date_row_idx, col_idx])
        if not re.search(r"\d{1,2}/\d{1,2}", date_cell): continue
        
        seen_today = {} # 儲存當日核心字：行號

        for r_idx in range(content_start, content_end):
            cell_val = str(df.iloc[r_idx, col_idx]).strip()
            
            # 過濾掉太短的字（如：單個字、空白、或是純數字的過敏原標示）
            if len(cell_val) < 2 or cell_val.isdigit(): continue
            
            # 提取核心字（例如：雞肉沙拉 -> 雞肉）
            core = clean_name(cell_val)[:2]
            if len(core) >= 2:
                if core in seen_today:
                    # 發現重複！
                    ws.cell(row=r_idx+1, column=col_idx+1).fill = YELLOW_FILL
                    prev_r = seen_today[core]
                    ws.cell(row=prev_r+1, column=col_idx+1).fill = YELLOW_FILL
                    
                    results.append({
                        "日期": date_cell,
                        "衝突項目": cell_val,
                        "原因": f"❌ 食材重複：與同日項目「{core}」衝突"
                    })
                seen_today[core] = r_idx

    wb.save(output)
    return results, output.getvalue()

# --- 介面呈現 ---
st.title("🛡️ 一月初 (暖禾) 輕食菜單自主稽核系統")
st.markdown("---")
st.info("💡 **專門優化版**：此版本已針對『一月初』特殊排版進行座標適應。")

up = st.file_uploader("👉 請上傳『一月初』週菜單 Excel", type=["xlsx"])

if up:
    with st.spinner("系統分析中..."):
        logs, final_file = audit_logic(up)
        
        if logs and isinstance(logs[0], str) and "❌" in logs[0]:
            st.error(logs[0])
        elif logs:
            st.error(f"🚩 偵測到 {len(logs)} 項食材重複！請下載修正。")
            st.download_button("📥 下載【一月初】標註檔案", final_file, f"一月初審核_{up.name}")
            st.table(pd.DataFrame(logs))
        else:
            st.success("🎉 完美！該週菜單無食材重複衝突。")
