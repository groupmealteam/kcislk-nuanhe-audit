import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# --- 樣式定義 ---
STYLE_VACUUM = {"fill": PatternFill("solid", fgColor="000000"), "font": Font(name="微軟正黑體", size=12, color="FFFFFF", bold=True)}
STYLE_LOW    = {"fill": PatternFill("solid", fgColor="FF0000"), "font": Font(name="微軟正黑體", size=12, color="FFFFFF", bold=True)}
STYLE_CAL    = {"fill": PatternFill("solid", fgColor="FFCCFF"), "font": Font(name="微軟正黑體", size=12, color="800000", bold=True)}

def clean_to_num(val):
    if pd.isna(val) or str(val).strip().lower() in ["", "nan", "none"]: return None
    try:
        res = re.findall(r"\d+\.?\d*", str(val))
        return float(res[0]) if res else 0.0
    except: return 0.0

def alison_nuanhe_audit_v3(file):
    wb = load_workbook(file)
    sheets_df = pd.read_excel(file, sheet_name=None, header=None)
    logs = []
    total_scan = 0
    
    # 暖禾高標 (全榖/豆魚: 4.0, 蔬菜: 2.0, 熱量: 750-850)
    LIMITS = {"熱量": (750, 850), "全榖": 4.0, "豆魚": 4.0, "蔬菜": 2.0}

    for sn, df in sheets_df.items():
        ws = wb[sn]
        # 1. 尋找日期行 (通常在第3-5列)
        d_row = None
        for i in range(len(df)):
            cell_content = str(df.iloc[i, 2]) # 檢查 C 欄
            if "日期" in cell_content or "Date" in cell_content:
                d_row = i
                break
        if d_row is None: continue

        # 2. 開始橫向掃描週一到週五 (C欄到G欄)
        for col in range(2, 7): # Index 2=C, 6=G
            if col >= len(df.columns): break
            
            # 抓取日期
            date_cell = str(df.iloc[d_row, col]).strip()
            if "202" not in date_cell: continue
            this_day = date_cell.split(" ")[0]

            # 3. 往下搜尋 30 列尋找目標標籤
            for r_off in range(1, 31):
                r_idx = d_row + r_off
                if r_idx >= len(df): break
                
                # 抓取標籤名稱 (B 欄或 C 欄)
                label = str(df.iloc[r_idx, 1]) + str(df.iloc[r_idx, 2])
                raw_val = df.iloc[r_idx, col]
                
                for key, limit in LIMITS.items():
                    if key in label:
                        total_scan += 1
                        val = clean_to_num(raw_val)
                        target_cell = ws.cell(row=r_idx+1, column=col+1)
                        
                        if val is None:
                            target_cell.fill, target_cell.font = STYLE_VACUUM["fill"], STYLE_VACUUM["font"]
                            target_cell.value = "❌真空"
                            logs.append({"分頁": sn, "日期": this_day, "項目": key, "原因": "真空漏填"})
                        elif key == "熱量":
                            if val < limit[0] or val > limit[1]:
                                target_cell.fill, target_cell.font = STYLE_CAL["fill"], STYLE_CAL["font"]
                                logs.append({"分頁": sn, "日期": this_day, "項目": "熱量", "原因": f"異常:{val}"})
                        else:
                            # 份數檢核 (0.0 安全過關)
                            if 0 < val < limit:
                                target_cell.fill, target_cell.font = STYLE_LOW["fill"], STYLE_LOW["font"]
                                logs.append({"分頁": sn, "日期": this_day, "項目": key, "原因": f"不足({val})"})
    
    out = BytesIO()
    wb.save(out)
    return logs, out.getvalue(), total_scan

# --- Streamlit UI ---
st.set_page_config(page_title="暖禾自主稽核", layout="wide")
st.title("🛡️ 輕食區(暖禾) 菜單自主稽核系統")
st.caption("製作者：Alison")

up = st.file_uploader("📂 上傳暖禾菜單 Excel", type=["xlsx"])
if up:
    try:
        results, excel_data, count = alison_nuanhe_audit_v3(up)
        st.info(f"📊 確實度：本次共深入核對了 {count} 個關鍵數據點。")
        
        if count < 10:
            st.error("❌ 格式不對！程式找不到足夠的營養標籤，請檢查表格排版。")
        elif results:
            st.error(f"🚩 發現 {len(results)} 項法規異常。")
            st.table(pd.DataFrame(results))
            st.download_button("📥 下載退件標註檔", excel_data, f"退件_{up.name}")
        else:
            st.success("🎉 完美！所有數據皆符合暖禾高標規範（已確認 0 值數據）。")
    except Exception as e:
        st.error(f"🔥 系統崩潰：{str(e)}。這通常是因為 Excel 格式變動太大。")
