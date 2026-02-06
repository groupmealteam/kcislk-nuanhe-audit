import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# 1. 網頁基本設定 (使用招牌名稱：一月初)
st.set_page_config(page_title="一月初 - 輕食菜單稽核系統", layout="wide")

# 設定違規標色 (黃色：針對原則九食材重複)
YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

def clean_dish_name(text):
    """提取核心菜名，排除日期、括號、過敏原符號等雜訊"""
    if pd.isna(text) or re.search(r"\d{1,2}/\d{1,2}", str(text)): return ""
    return "".join(re.findall(r'[\u4e00-\u9fa5]+', str(text)))

def audit_logic(file):
    try:
        wb = load_workbook(file)
        all_sheets = pd.read_excel(file, sheet_name=None, header=None)
    except:
        return ["❌ 檔案無法讀取，請確保上傳的是 Excel (.xlsx) 格式。"], None

    results = []
    output = BytesIO()
    is_menu_found = False

    for sheet_name, df in all_sheets.items():
        df = df.fillna("")
        ws = wb[sheet_name]
        
        # 定位關鍵行：日期行與主副食行 (鎖定 B 欄的關鍵字)
        date_row = next((i for i, row in df.iterrows() if any(re.search(r"\d{1,2}/\d{1,2}", str(c)) for c in row)), None)
        target_rows = [i for i, row in df.iterrows() if any(k in str(row[1]) for k in ["主食", "副菜", "套餐", "麵食", "輕食"])]
        
        if date_row is None or not target_rows:
            continue
            
        is_menu_found = True

        # 逐欄審核 (週一至週五)
        for col in range(2, len(df.columns)):
            date_val = str(df.iloc[date_row, col])
            if not re.search(r"\d{1,2}/\d{1,2}", date_val): continue
            
            seen_today = {} # 儲存當日已出現的食材核心字
            daily_soups = [] # 儲存當日 A/B 餐的湯品

            for r_idx in target_rows:
                cell_val = str(df.iloc[r_idx, col]).strip()
                if not cell_val or len(cell_val) < 2: continue

                # A. 收集湯品資訊 (原則：湯品須一致)
                if "湯" in cell_val or "羹" in cell_val:
                    daily_soups.append(cell_val)

                # B. 食材重複稽核 (原則九：避免學生辨識出相同食物)
                dish_core = clean_dish_name(cell_val)[:2] # 抓取前兩個中文字作為核心
                if len(dish_core) >= 2:
                    if dish_core in seen_today:
                        # 標註當前格子與先前重複的格子
                        ws.cell(row=r_idx+1, column=col+1).fill = YELLOW_FILL
                        prev_row_idx = seen_today[dish_core]
                        ws.cell(row=prev_row_idx+1, column=col+1).fill = YELLOW_FILL
                        
                        results.append({
                            "分頁": sheet_name, 
                            "日期": date_val, 
                            "項目": cell_val, 
                            "問題": f"❌ 食材重複：與同日其他餐點「{dish_core}」重複 (原則九)"
                        })
                    seen_today[dish_core] = r_idx

            # C. 湯品一致性檢查
            if len(set(daily_soups)) > 1:
                results.append({
                    "分頁": sheet_name, 
                    "日期": date_val, 
                    "項目": "湯品比對", 
                    "問題": "⚠️ 湯品不一致：A/B餐當日湯品應保持相同，以利行政作業"
                })

    if not is_menu_found:
        return ["❌ 偵測失敗：上傳檔案不含日期或主食欄位，請確認是否為『一月初』正式菜單。"], None

    wb.save(output)
    return results, output.getvalue()

# --- 網頁畫面 ---
st.title("🛡️ 一月初 (暖禾) 輕食菜單自主稽核系統")
st.markdown("---")
st.info("💡 **自主檢查重點：** 1. 同日 A/B 餐食材重複避讓 (原則九)  2. 湯品一致性檢查。")

up = st.file_uploader("👉 請上傳『一月初』菜單 Excel (.xlsx)", type=["xlsx"])

if up:
    with st.spinner("一月初 稽核系統分析中..."):
        logs, final_excel = audit_logic(up)
        
        if logs and isinstance(logs[0], str) and logs[0].startswith("❌"):
            st.error(logs[0])
        elif logs:
            st.error(f"🚩 偵測到 {len(logs)} 項衝突或建議，請下載檔案查看：")
            st.download_button(
                label="📥 下載『一月初』標註版檔案",
                data=final_excel,
                file_name=f"一月初_審核_{up.name}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.table(pd.DataFrame(logs))
        else:
            st.success("🎉 審核通過！該週菜單食材配置完美，無重複或衝突。")
