import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment

# 1. 網頁基本設定
st.set_page_config(page_title="輕食區(一月初) 菜單自主稽核系統", layout="wide")

# --- 定義【30級字 + 微軟正黑體】最終規範顏色 ---
FONT_NAME = "微軟正黑體"
FONT_SIZE = 30

# 顏色定義：嚴格遵循 Alison 要求
PORTION_FAIL = {"fill": PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid"), "font": Font(name=FONT_NAME, size=FONT_SIZE, color="FFFFFF", bold=True)}
CALORIE_FAIL = {"fill": PatternFill(start_color="FFCCFF", end_color="FFCCFF", fill_type="solid"), "font": Font(name=FONT_NAME, size=FONT_SIZE, color="800000", bold=True)}
REPEAT_FAIL  = {"fill": PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid"), "font": Font(name=FONT_NAME, size=FONT_SIZE, color="FF0000", bold=True)}
SPICY_FAIL   = {"fill": PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"), "font": Font(name=FONT_NAME, size=FONT_SIZE, color="000000", bold=True)}

# 食材肉種精確識別字典 (落實原則九)
MEAT_DICT = {
    "豬": ["豬", "肉絲", "肉片", "排骨", "焢肉", "培根", "火腿", "叉燒", "里肌", "貢丸"],
    "雞": ["雞", "翅", "鳳", "鳥", "咔啦", "柳", "腿"],
    "牛": ["牛"],
    "魚": ["魚", "吻", "柳葉", "海鮮", "蝦", "花枝"],
    "蛋": ["蛋"],
    "豆": ["豆", "腐", "干", "素肉", "麵腸"]
}

def get_meat(text):
    if not text or "水果" in text: return None
    for key, words in MEAT_DICT.items():
        if any(w in text for w in words): return key
    return text[:2]

def to_num(val):
    try:
        res = re.findall(r"\d+\.?\d*", str(val))
        return float(res[0]) if res else 0.0
    except: return 0.0

def audit_process(file):
    wb = load_workbook(file)
    sheets_df = pd.read_excel(file, sheet_name=None, header=None)
    logs = []
    output = BytesIO()

    for sn, df in sheets_df.items():
        ws = wb[sn]
        df = df.fillna("")
        # 定位日期列 (通常在 index 2 或 3)
        d_row = next((i for i, r in df.iterrows() if "日期Date" in str(r[2])), None)
        if d_row is None: continue

        prev_A, prev_B = None, None # 紀錄昨日肉種

        for col in range(3, 8): # 週一到週五
            day_name = str(df.iloc[d_row+1, col])
            date_str = str(df.iloc[d_row, col]).split(" ")[0]
            
            # --- A. 教育部營養法規審核 ---
            for r_idx in range(d_row + 10, len(df)):
                label = str(df.iloc[r_idx, 2])
                val = to_num(df.iloc[r_idx, col])
                cell = ws.cell(row=r_idx+1, column=col+1)
                
                if "熱量" in label and (val < 750 or val > 850):
                    cell.fill, cell.font = CALORIE_FAIL["fill"], CALORIE_FAIL["font"]
                    logs.append({"分頁": sn, "日期": date_str, "項目": "熱量", "原因": f"粉底：{val}不符750-850"})
                elif "全榖" in label and val < 4.0:
                    cell.fill, cell.font = PORTION_FAIL["fill"], PORTION_FAIL["font"]
                    logs.append({"分頁": sn, "日期": date_str, "項目": "全榖", "原因": "紅底白字：不足4份"})
                elif "豆魚" in label and val < 4.0:
                    cell.fill, cell.font = PORTION_FAIL["fill"], PORTION_FAIL["font"]
                    logs.append({"分頁": sn, "日期": date_str, "項目": "蛋白質", "原因": "紅底白字：不足4份"})
                elif "蔬菜" in label and val < 2.0:
                    cell.fill, cell.font = PORTION_FAIL["fill"], PORTION_FAIL["font"]
                    logs.append({"分頁": sn, "日期": date_str, "項目": "蔬菜", "原因": "紅底白字：不足2份"})

            # --- B. 食材重複性規範 (原則九) ---
            # 抓取 A/B 主食
            main_A_idx = d_row + 3
            main_A_name = str(df.iloc[main_A_idx, col])
            meat_A = get_meat(main_A_name)
            
            # 動態尋找 B 餐主食
            label_B = next((i for i in range(d_row+5, len(df)) if "輕食B餐" in str(df.iloc[i, 2])), None)
            main_B_idx = label_B + 1 if label_B else None
            meat_B = get_meat(str(df.iloc[main_B_idx, col])) if main_B_idx else None

            # 1. 餐道間 (A vs B)
            if meat_A and meat_B and meat_A == meat_B:
                for r in [main_A_idx, main_B_idx]:
                    ws.cell(row=r+1, column=col+1).fill = REPEAT_FAIL["fill"]
                    ws.cell(row=r+1, column=col+1).font = REPEAT_FAIL["font"]
                logs.append({"分頁": sn, "日期": date_str, "項目": "餐道衝突", "原因": f"黃底紅字：A/B餐主食皆為{meat_A}"})
            
            # 2. 跨日比對 (原則九：同餐道主食不連兩天)
            if meat_A and meat_A == prev_A:
                ws.cell(row=main_A_idx+1, column=col+1).fill = REPEAT_FAIL["fill"]
                logs.append({"分頁": sn, "日期": date_str, "項目": "跨日重複", "原因": f"黃底紅字：A餐與昨日重複({meat_A})"})
            
            prev_A, prev_B = meat_A, meat_B

            # --- C. 校內禁辣原則 ---
            if any(day in day_name for day in ["週一", "週二", "週四"]):
                for r_idx in range(d_row + 2, d_row + 15):
                    if "水果" in str(df.iloc[r_idx, 2]): continue
                    content = str(df.iloc[r_idx, col])
                    if "●" in content or "🌶️" in content:
                        ws.cell(row=r_idx+1, column=col+1).fill = SPICY_FAIL["fill"]
                        ws.cell(row=r_idx+1, column=col+1).font = SPICY_FAIL["font"]
                        logs.append({"分頁": sn, "日期": date_str, "項目": "禁辣日", "原因": "淺綠底：禁辣日標辣"})

    wb.save(output)
    return logs, output.getvalue()

# --- Streamlit 介面 ---
st.title("🛡️ 輕食區(一月初) 菜單自主稽核系統")
st.caption("製作者：Alison / 版本：V5.6 (合約規範對標版)")

st.info("📌 本系統嚴格執行：1. 營養法規(紅粉底) 2. 原則九食材多樣性(黃底紅字) 3. 禁辣日(淺綠底)")

up = st.file_uploader("👉 上傳菜單 Excel", type=["xlsx"])
if up:
    results, data = audit_process(up)
    if data:
        if results:
            st.error(f"🚩 偵測到 {len(results)} 項不符規範")
            st.download_button("📥 下載退件標註檔", data, f"稽核結果_{up.name}")
            st.table(pd.DataFrame(results))
        else:
            st.success("🎉 完美！該週菜單符合所有合約與法規。")
