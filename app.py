import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment

# 1. 網頁基本設定
st.set_page_config(page_title="輕食區(一月初) 菜單自主稽核系統", layout="wide")

# --- 定義規範顏色與大字體 ---
FONT_NAME = "微軟正黑體"
FONT_SIZE = 30

# 樣式定義
REPEAT_STYLE = {"fill": PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid"), "font": Font(name=FONT_NAME, size=FONT_SIZE, color="FF0000", bold=True)}
CALORIE_STYLE = {"fill": PatternFill(start_color="FFCCFF", end_color="FFCCFF", fill_type="solid"), "font": Font(name=FONT_NAME, size=FONT_SIZE, color="800000", bold=True)}
PORTION_STYLE = {"fill": PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid"), "font": Font(name=FONT_NAME, size=FONT_SIZE, color="FFFFFF", bold=True)}
SPICY_STYLE = {"fill": PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"), "font": Font(name=FONT_NAME, size=FONT_SIZE, color="000000", bold=True)}

CENTER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)

# 食材肉種識別字典
MEAT_MAP = {"豬": ["豬", "肉絲", "肉片", "排骨", "焢肉", "培根", "火腿", "叉燒"], "雞": ["雞", "翅", "鳳", "鳥"], "牛": ["牛"], "魚": ["魚", "吻仔魚", "柳葉魚", "鮮魚"], "蛋": ["蛋"], "豆": ["豆", "腐", "干"]}

def get_meat_type(text):
    if not text or "水果" in text or "Fruit" in text: return None
    for key, keywords in MEAT_MAP.items():
        if any(word in text for word in keywords): return key
    return text[:2]

def clean_name(text):
    if pd.isna(text): return ""
    return "".join(re.findall(r'[\u4e00-\u9fa5]+', str(text).split('\n')[0]))

def to_float(val):
    try:
        nums = re.findall(r"\d+\.?\d*", str(val))
        return float(nums[0]) if nums else 0.0
    except: return 0.0

def audit_process(file):
    try:
        wb = load_workbook(file)
        all_sheets = pd.read_excel(file, sheet_name=None, header=None)
        results = []
        output = BytesIO()

        for sheet_name, df in all_sheets.items():
            df = df.fillna("")
            ws = wb[sheet_name]
            date_row = next((i for i, row in df.iterrows() if "日期Date" in str(row[2])), None)
            if date_row is None: continue

            prev_meat_A = None 

            for col in range(3, 8): 
                if col >= len(df.columns): break
                date_val = str(df.iloc[date_row, col]).split(" ")[0]
                day_val = str(df.iloc[date_row+1, col])
                if "202" not in date_val: continue

                # --- 1. 食材重複性審核 (黃底紅字) ---
                idx_A_main = date_row + 3
                dish_A_main = clean_name(df.iloc[idx_A_main, col])
                
                idx_B_label = next((i for i in range(date_row + 5, len(df)) if "輕食B餐" in str(df.iloc[i, 2])), None)
                idx_B_main = idx_B_label + 1 if idx_B_label else None
                dish_B_main = clean_name(df.iloc[idx_B_main, col]) if idx_B_main else ""

                meat_A = get_meat_type(dish_A_main)
                meat_B = get_meat_type(dish_B_main)

                # 餐道間重複 (A vs B)
                if meat_A and meat_B and meat_A == meat_B:
                    for r in [idx_A_main, idx_B_main]:
                        if r:
                            ws.cell(row=r+1, column=col+1).fill = REPEAT_STYLE["fill"]
                            ws.cell(row=r+1, column=col+1).font = REPEAT_STYLE["font"]
                    results.append({"日期": date_val, "項目": "食材重複", "原因": f"⚠️ A/B餐主食重複({meat_A})"})

                # 跨日重複
                if meat_A and meat_A == prev_meat_A:
                    ws.cell(row=idx_A_main+1, column=col+1).fill = REPEAT_STYLE["fill"]
                    results.append({"日期": date_val, "項目": "跨日重複", "原因": f"⚠️ A餐連續重複({meat_A})"})
                prev_meat_A = meat_A

                # --- 2. 營養標示審核 (粉底/紅底) ---
                for r_idx in range(date_row + 10, len(df)):
                    label = str(df.iloc[r_idx, 2])
                    val = to_float(df.iloc[r_idx, col])
                    cell = ws.cell(row=r_idx+1, column=col+1)
                    if "熱量" in label and (val < 750 or val > 850):
                        cell.fill = CALORIE_STYLE["fill"]; cell.font = CALORIE_STYLE["font"]
                        results.append({"日期": date_val, "項目": "熱量", "原因": "粉底：熱量異常"})
                    elif any(x in label for x in ["全榖", "豆魚", "蔬菜"]):
                        standard_val = 4.0 if "全榖" in label or "豆魚" in label else 2.0
                        if val < standard_val:
                            cell.fill = PORTION_STYLE["fill"]; cell.font = PORTION_STYLE["font"]
                            results.append({"日期": date_val, "項目": "份數", "原因": "紅底白字：份數不足"})

                # --- 3. 禁辣日審核 (淺綠底) ---
                if any(d in day_val for d in ["週一", "週二", "週四"]):
                    for r_idx in range(date_row + 2, date_row + 12):
                        if "水果" in str(df.iloc[r_idx, 2]): continue
                        cell_val = str(df.iloc[r_idx, col])
                        if "●" in cell_val or "🌶️" in cell_val:
                            cell = ws.cell(row=r_idx+1, column=col+1)
                            cell.fill = SPICY_STYLE["fill"]; cell.font = SPICY_STYLE["font"]
                            results.append({"日期": date_val, "項目": "禁辣日", "原因": "淺綠底：違規"})

        wb.save(output)
        return results, output.getvalue()
    except Exception as e:
        return [f"❌ 程式發生錯誤：{str(e)}"], None

# --- 介面 ---
st.title("🛡️ 輕食區(一月初) 菜單自主稽核系統")
st.caption("製作者：Alison / 版本：V5.5")
up = st.file_uploader("👉 上傳 Excel", type=["xlsx"])
if up:
    logs, data = audit_process(up)
    if data:
        st.error(f"🚩 偵測到 {len(logs)} 項異常")
        st.download_button("📥 下載退件標註檔", data, f"稽核報告_{up.name}")
        st.table(pd.DataFrame(logs))
    else:
        st.error(logs[0])
