import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# --- 樣式定義 ---
STYLE_VACUUM = {"fill": PatternFill("solid", fgColor="000000"), "font": Font(size=12, color="FFFFFF", bold=True)}
STYLE_LOW    = {"fill": PatternFill("solid", fgColor="FF0000"), "font": Font(size=12, color="FFFFFF", bold=True)}
STYLE_CAL    = {"fill": PatternFill("solid", fgColor="FFCCFF"), "font": Font(size=12, color="800000", bold=True)}

def to_num(val):
    if pd.isna(val) or str(val).strip() == "": return None
    try:
        res = re.findall(r"\d+\.?\d*", str(val))
        return float(res[0]) if res else 0.0
    except: return 0.0

def alison_nuanhe_audit(file):
    wb = load_workbook(file)
    sheets_df = pd.read_excel(file, sheet_name=None, header=None)
    logs = []
    scan_points = 0 # 確實度計數器
    
    # 暖禾標準門檻
    LIMITS = {"熱量": (750, 850), "全榖": 4.0, "豆魚": 4.0, "蔬菜": 2.0}

    for sn, df in sheets_df.items():
        ws = wb[sn]
        # 尋找日期 Date 行 (確保 C 欄有日期)
        d_row = None
        for i, r in df.iterrows():
            if any(k in str(r[2]) for k in ["日期", "Date"]):
                d_row = i
                break
        if d_row is None: continue

        for col in range(3, 8):
            date_val = str(df.iloc[d_row, col]).split(" ")[0]
            if "202" not in date_val: continue

            # 往下掃描 20 列尋找目標
            for r_offset in range(1, 25):
                curr_row = d_row + r_offset
                if curr_row >= len(df): break
                
                label = str(df.iloc[curr_row, 2])
                raw_val = df.iloc[curr_row, col]
                cell = ws.cell(row=curr_row+1, column=col+1)
                
                # --- 確實稽核：只有對中關鍵字才處理 ---
                for key, threshold in LIMITS.items():
                    if key in label:
                        scan_points += 1
                        val = to_num(raw_val)
                        
                        if val is None: # 真空
                            cell.fill, cell.font = STYLE_VACUUM["fill"], STYLE_VACUUM["font"]
                            cell.value = "❌漏填"
                            logs.append({"分頁": sn, "日期": date_label, "項目": label, "原因": "真空漏填"})
                        elif key == "熱量":
                            if val < threshold[0] or val > threshold[1]:
                                cell.fill, cell.font = STYLE_CAL["fill"], STYLE_CAL["font"]
                                logs.append({"日期": date_val, "項目": "熱量", "原因": f"熱量異常:{val}"})
                        else: # 份數
                            if 0 < val < threshold:
                                cell.fill, cell.font = STYLE_LOW["fill"], STYLE_LOW["font"]
                                logs.append({"日期": date_val, "項目": label, "原因": f"不足({val})"})
    
    output = BytesIO()
    wb.save(output)
    return logs, output.getvalue(), scan_count

# --- UI ---
st.set_page_config(page_title="暖禾自主稽核", layout="wide")
st.title("🛡️ 輕食區(暖禾) 菜單自主稽核系統")
st.caption("製作者：Alison")

up = st.file_uploader("📂 請上傳菜單", type=["xlsx"])
if up:
    logs, data, count = alison_nuanhe_audit(up)
    
    # 💡 確實度報告：如果 count 太低，直接報警
    st.info(f"📊 本次共深入稽核了 {count} 個營養數據點。")
    
    if count < 10:
        st.error("❌ 格式識別嚴重錯誤！程式找不到足夠的營養數據，請確認是否上傳了『正確格式』的輕食菜單。")
    elif logs:
        st.error(f"🚩 偵測到 {len(logs)} 項異常。")
        st.table(pd.DataFrame(logs))
        st.download_button("📥 下載退件標註檔", data, f"退件_{up.name}")
    else:
        st.success("🎉 完美！所有數據皆符合暖禾高標規範。")
