import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment

# 1. 網頁基本設定
st.set_page_config(page_title="輕食區(一月初) 菜單自主稽核系統", layout="wide")

# --- 定義視覺規範 (30級字 + 微軟正黑體) ---
FONT_NAME = "微軟正黑體"
FONT_SIZE = 30

PORTION_STYLE = {"fill": PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid"), "font": Font(name=FONT_NAME, size=FONT_SIZE, color="FFFFFF", bold=True)}
CALORIE_STYLE = {"fill": PatternFill(start_color="FFCCFF", end_color="FFCCFF", fill_type="solid"), "font": Font(name=FONT_NAME, size=FONT_SIZE, color="800000", bold=True)}
REPEAT_STYLE  = {"fill": PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid"), "font": Font(name=FONT_NAME, size=FONT_SIZE, color="FF0000", bold=True)}
SPICY_STYLE   = {"fill": PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"), "font": Font(name=FONT_NAME, size=FONT_SIZE, color="000000", bold=True)}
# 真空漏填專用：黑底白字 (設定較小字級以免撐開表格)
VACUUM_STYLE  = {"fill": PatternFill(start_color="000000", end_color="000000", fill_type="solid"), "font": Font(name=FONT_NAME, size=14, color="FFFFFF", bold=True)}

MEAT_DICT = {"豬": ["豬", "肉絲", "肉片", "排骨", "焢肉", "培根", "火腿", "里肌"], "雞": ["雞", "翅", "鳳", "咔啦", "柳", "腿"], "牛": ["牛"], "魚": ["魚", "吻仔", "海鮮", "蝦"], "蛋": ["蛋"], "豆": ["豆", "腐", "干", "素肉"]}

def get_meat(text):
    if not text or "水果" in text or "Fruit" in text: return None
    for key, words in MEAT_DICT.items():
        if any(w in text for w in words): return key
    return text[:2] if len(text) >= 2 else None

def to_num(val):
    try:
        if pd.isna(val) or str(val).strip() == "": return -1.0 # 用 -1 代表真空
        res = re.findall(r"\d+\.?\d*", str(val))
        return float(res[0]) if res else 0.0
    except: return 0.0

def audit_process(file):
    # --- 強制關鍵字檢查 ---
    if "輕食" not in file.name:
        return ["❌ 錯誤：檔名不含『輕食』關鍵字，拒絕稽核"], None

    try:
        wb = load_workbook(file)
        sheets_df = pd.read_excel(file, sheet_name=None, header=None)
        logs = []
        output = BytesIO()

        # 🎯 只稽核這些標籤，解決 228 項 nan 異常
        TARGET_LABELS = ["熱量", "全榖", "豆魚", "蔬菜", "水果", "油脂", "乳品", "奶類"]

        for sn, df in sheets_df.items():
            ws = wb[sn]
            d_row = next((i for i, r in df.iterrows() if "日期Date" in str(r[2])), None)
            if d_row is None: continue

            for col in range(3, 8):
                if col >= len(df.columns): break
                date_val = df.iloc[d_row, col]
                date_str = str(date_val).split(" ")[0]
                if "202" not in date_str: continue
                day_name = str(df.iloc[d_row+1, col]) if (d_row+1) < len(df) else ""

                # --- 營養標示精準稽核 ---
                for r_idx in range(d_row + 10, len(df)):
                    label_raw = df.iloc[r_idx, 2]
                    label = str(label_cell) if not pd.isna(label_raw) else ""
                    
                    # 過濾非目標列：如果名稱沒在 TARGET_LABELS 裡，直接跳過
                    if not any(t in label for t in TARGET_LABELS): continue

                    raw_val = df.iloc[r_idx, col]
                    val = to_num(raw_val)
                    cell = ws.cell(row=r_idx+1, column=col+1)

                    # 1. 判定真空 (-1)
                    if val == -1.0:
                        cell.fill, cell.font = VACUUM_STYLE["fill"], VACUUM_STYLE["font"]
                        cell.value = "❌真空漏填"
                        logs.append({"分頁": sn, "日期": date_str, "項目": label, "原因": "真空漏填"})
                    
                    # 2. 判定數值 (0.0 會被視為合法，不噴色)
                    else:
                        if "熱量" in label and (val < 750 or val > 850):
                            cell.fill, cell.font = CALORIE_STYLE["fill"], CALORIE_STYLE["font"]
                            logs.append({"分頁": sn, "日期": date_str, "項目": "熱量", "原因": f"粉底：{val} Kcal"})
                        elif any(x in label for x in ["全榖", "豆魚", "蔬菜"]):
                            std = 2.0 if "蔬菜" in label else 4.0
                            # 關鍵：只有 0 < val < std 才報錯，排除 0.0
                            if 0 < val < std:
                                cell.fill, cell.font = PORTION_STYLE["fill"], PORTION_STYLE["font"]
                                logs.append({"分頁": sn, "日期": date_str, "項目": label, "原因": "紅底白字：份數不足"})

                # --- 2. 食材重複性 ---
                main_A_idx = d_row + 3
                meat_A = get_meat(str(df.iloc[main_A_idx, col])) if main_A_idx < len(df) else None
                label_B = next((i for i in range(d_row+5, len(df)) if "輕食B餐" in str(df.iloc[i, 2])), None)
                main_B_idx = label_B + 1 if label_B else None
                meat_B = get_meat(str(df.iloc[main_B_idx, col])) if main_B_idx and main_B_idx < len(df) else None

                if meat_A and meat_B and meat_A == meat_B:
                    for r in [main_A_idx, main_B_idx]:
                        if r and r < len(df):
                            c = ws.cell(row=r+1, column=col+1)
                            c.fill, c.font = REPEAT_STYLE["fill"], REPEAT_STYLE["font"]
                    logs.append({"分頁": sn, "日期": date_str, "項目": "餐道衝突", "原因": "黃底紅字：肉種重複"})

                # --- 3. 禁辣原則 ---
                if any(day in day_name for day in ["週一", "週二", "週四"]):
                    for r_idx in range(d_row + 2, d_row + 12):
                        if r_idx >= len(df) or "水果" in str(df.iloc[r_idx, 2]): continue
                        txt = str(df.iloc[r_idx, col])
                        if "●" in txt or "🌶️" in txt:
                            cell = ws.cell(row=r_idx+1, column=col+1)
                            cell.fill, cell.font = SPICY_STYLE["fill"], SPICY_STYLE["font"]
                            logs.append({"分頁": sn, "日期": date_str, "項目": "禁辣日", "原因": "淺綠底：違規標辣"})

        wb.save(output)
        return logs, output.getvalue()
    except Exception as e:
        return [f"發生錯誤：{str(e)}"], None

# --- UI 介面 (完全保留妳原本的視覺) ---
st.title("🛡️ 輕食區(一月初) 菜單自主稽核系統")
st.caption("製作者：Alison")
st.markdown("---")

up = st.file_uploader("👉 上傳菜單 Excel", type=["xlsx"])
if up:
    with st.spinner("稽核中..."):
        results, data = audit_process(up)
        if data:
            if results and not isinstance(results[0], str):
                st.error(f"🚩 偵測到 {len(results)} 項異常")
                st.download_button("📥 下載退件標註檔", data, f"稽核結果_{up.name}")
                st.table(pd.DataFrame(results))
            elif isinstance(results[0], str) and "錯誤" in results[0]:
                st.error(results[0])
            else:
                st.success("🎉 通過所有稽核項目！")
        elif results:
            st.error(results[0])
