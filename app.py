import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# 樣式定義
STYLE_VACUUM = {"fill": PatternFill("solid", fgColor="000000"), "font": Font(name="微軟正黑體", size=12, color="FFFFFF", bold=True)}
STYLE_LOW    = {"fill": PatternFill("solid", fgColor="FF0000"), "font": Font(name="微軟正黑體", size=12, color="FFFFFF", bold=True)}
STYLE_CAL    = {"fill": PatternFill("solid", fgColor="FFCCFF"), "font": Font(name="微軟正黑體", size=12, color="800000", bold=True)}

def to_num(val):
    if pd.isna(val) or str(val).strip() == "" or str(val).strip().lower() == "nan": return None
    try:
        res = re.findall(r"\d+\.?\d*", str(val))
        return float(res[0]) if res else 0.0
    except: return 0.0

def alison_nuanhe_audit(file):
    wb = load_workbook(file)
    sheets_df = pd.read_excel(file, sheet_name=None, header=None)
    logs = []
    scan_points = 0
    # 暖禾標準 (全榖/豆魚: 4.0, 蔬菜: 2.0)
    LIMITS = {"熱量": (750, 850), "全榖": 4.0, "豆魚": 4.0, "蔬菜": 2.0}

    for sn, df in sheets_df.items():
        ws = wb[sn]
        # 尋找日期行 (精確定位 C 欄)
        d_row = next((i for i, r in df.iterrows() if any(k in str(r[2]) for k in ["日期", "Date"])), None)
        if d_row is None: continue

        for col in range(3, 8):
            raw_date = str(df.iloc[d_row, col]).strip()
            if "202" not in raw_date: continue
            current_date = raw_date.split(" ")[0] # 修正變數名稱一致性

            # 往下掃描 25 列
            for r_off in range(1, 26):
                row_idx = d_row + r_off
                if row_idx >= len(df): break
                
                header = str(df.iloc[row_idx, 2])
                raw_val = df.iloc[row_idx, col]
                
                for key, limit in LIMITS.items():
                    if key in header: # 模糊匹配：全榖雜糧也算全榖
                        scan_points += 1
                        val = to_num(raw_val)
                        cell = ws.cell(row=row_idx+1, column=col+1)
                        
                        if val is None:
                            cell.fill, cell.font = STYLE_VACUUM["fill"], STYLE_VACUUM["font"]
                            cell.value = "❌真空"
                            logs.append({"分頁": sn, "日期": current_date, "項目": header, "原因": "真空漏填"})
                        elif key == "熱量":
                            if val < limit[0] or val > limit[1]:
                                cell.fill, cell.font = STYLE_CAL["fill"], STYLE_CAL["font"]
                                logs.append({"分頁": sn, "日期": current_date, "項目": header, "原因": f"熱量異常:{val}"})
                        else:
                            # 0.0 是合法數據，(0, limit) 才是不足
                            if 0 < val < limit:
                                cell.fill, cell.font = STYLE_LOW["fill"], STYLE_LOW["font"]
                                logs.append({"分頁": sn, "日期": current_date, "項目": header, "原因": f"份數不足({val})"})
    
    output = BytesIO()
    wb.save(output)
    return logs, output.getvalue(), scan_points

# UI
st.title("🛡️ 輕食區(暖禾) 菜單自主稽核系統")
st.caption("製作者：Alison")

up = st.file_uploader("📂 上傳菜單", type=["xlsx"])
if up:
    try:
        logs, data, count = alison_nuanhe_audit(up)
        st.info(f"📊 確實度報告：本次共深入稽核了 {count} 個營養數據點。")
        
        if count < 10:
            st.error("❌ 內容識別失敗！程式找不到標籤，請確認是否為『暖禾輕食』格式。")
        elif logs:
            st.error(f"🚩 偵測到 {len(logs)} 項異常。")
            st.table(pd.DataFrame(logs))
            st.download_button("📥 下載退件標註檔", data, f"退件_{up.name}")
        else:
            st.success("🎉 數據稽核確實！所有偵測到的欄位皆包含有效值（含 0 值）。")
    except Exception as e:
        st.error(f"🔥 系統崩潰：{str(e)}")
