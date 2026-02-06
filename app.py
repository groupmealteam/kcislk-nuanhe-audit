import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# 1. 網頁基本設定
st.set_page_config(page_title="一月初 - 輕食餐道稽核系統", layout="wide")
YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

def clean_name(text):
    if pd.isna(text): return ""
    # 僅提取第一行中文字，過濾掉英文與括號雜訊
    return "".join(re.findall(r'[\u4e00-\u9fa5]+', str(text).split('\n')[0]))

def audit_logic(file):
    try:
        wb = load_workbook(file)
        all_sheets = pd.read_excel(file, sheet_name=None, header=None)
    except:
        return ["❌ 檔案讀取失敗"], None

    results = []
    output = BytesIO()
    found_menu = False

    for sheet_name, df in all_sheets.items():
        df = df.fillna("")
        ws = wb[sheet_name]
        
        # 定位「日期Date」行
        date_row = next((i for i, row in df.iterrows() if any("日期Date" in str(x) for x in row)), None)
        if date_row is None: continue
        found_menu = True

        # 遍歷 C 到 G 欄 (週一至週五)
        for col_idx in range(3, 8):
            if col_idx >= len(df.columns): break
            date_val = str(df.iloc[date_row, col_idx])
            if "2026" not in date_val: continue # 確認是日期欄位
            
            seen_ingredients = {} # 存放當天(A+B餐)已出現的核心詞

            # 搜尋當天所有餐點行
            for r_idx in range(date_row + 2, len(df)):
                row_label = str(df.iloc[r_idx, 2]) # 抓取 B 欄標籤
                
                # --- 排除不需要檢查的項目 ---
                # 1. 忽略水果 (您的要求)
                # 2. 忽略湯品 (您的要求：A/B 湯品本來就一樣)
                # 3. 排除熱量、成分內容等輔助資訊
                skip_keywords = ["水果", "湯品", "食材內容", "熱量", "份", "類", "星期"]
                if any(k in row_label for k in skip_keywords): continue
                
                cell_val = str(df.iloc[r_idx, col_idx]).strip()
                # 排除空值、純數字、或過短的字
                if len(cell_val) < 2 or cell_val.isdigit() or "輕食" in cell_val: continue

                # --- 食材重複比對 ---
                core = clean_name(cell_val)[:2] # 抓取核心詞前兩字 (例如：燒肉、排骨)
                if len(core) >= 2:
                    if core in seen_ingredients:
                        # 發現重複 (例如 A餐主食 跟 B餐主食 都有"豬肉")
                        ws.cell(row=r_idx+1, column=col_idx+1).fill = YELLOW_FILL
                        prev_r = seen_ingredients[core]
                        ws.cell(row=prev_r+1, column=col_idx+1).fill = YELLOW_FILL
                        
                        results.append({
                            "日期": date_val,
                            "衝突項目": f"「{cell_val}」與同日其他餐道重複",
                            "原因": f"❌ 食材核心字「{core}」重複 (原則九)"
                        })
                    seen_ingredients[core] = r_idx

    wb.save(output)
    return results, output.getvalue()

# --- 介面呈現 ---
st.title("🛡️ 一月初 (暖禾) 輕食餐道自主稽核系統")
st.markdown("---")
st.info("💡 **邏輯已更新**：已忽略水果與湯品，專注比對 A/B 餐道間的主食與副菜重複性。")

up = st.file_uploader("👉 請上傳『一月初』Excel 菜單", type=["xlsx"])

if up:
    logs, final_file = audit_logic(up)
    if logs:
        st.error(f"🚩 發現 {len(logs)} 項食材重複！(已標註黃色)")
        st.download_button("📥 下載標註檔案", final_file, f"一月初審核_{up.name}")
        st.table(pd.DataFrame(logs))
    else:
        st.success("🎉 完美！A/B 餐道配置多元，無食材重複衝突。")
