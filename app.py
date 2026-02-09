import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# --- 1. 標題與基礎設定 (完全遵照 Alison 規範) ---
ST_TITLE = "🛡️ 輕食區(一月初) 菜單自主稽核系統"
ST_AUTHOR = "製作者：Alison"

STYLE_ERR = {
    "fill": PatternFill("solid", fgColor="000000"), 
    "font": Font(name="微軟正黑體", size=12, color="FFFFFF", bold=True)
}

def alison_light_audit(file):
    fname = file.name
    
    # --- 2. 嚴格模式判定：關鍵字 "輕食" ---
    if "輕食" in fname:
        mode = "輕食區(暖禾)"
        # 輕食模式座標：通常 F-L 欄位為營養分析 (索引 5-11)
        nutri_indices = [5, 6, 7, 8, 9, 10, 11] 
        data_indices = [1, 2] # 輕食主項 B-C 欄
    else:
        # 如果沒看到輕食，直接阻斷，不亂套模式
        return None, "BLOCK", None

    wb = load_workbook(file)
    sheets_df = pd.read_excel(file, sheet_name=None, header=None)
    logs = []

    for sn, df in sheets_df.items():
        ws = wb[sn]
        # 【重要】只處理真正的 NaN，不准動廠商填的 '0'
        df_audit = df.astype(str).replace(['nan', 'NaN', 'None'], '')
        
        for r_idx in range(len(df_audit)):
            label = str(df_audit.iloc[r_idx, 0]).strip()
            
            # 只有日期行 (含有 / 和 ( ) 才稽核
            if "/" in label and "(" in label:
                # 檢查營養成分：只有「真空」才算缺失
                for n_idx in nutri_indices:
                    if n_idx < len(df_audit.columns):
                        val = df_audit.iloc[r_idx, n_idx].strip()
                        
                        # 0 絕對不噴黑，只有空字串才噴黑
                        if val == "":
                            cell = ws.cell(row=r_idx+1, column=n_idx+1)
                            cell.fill, cell.font = STYLE_ERR["fill"], STYLE_ERR["font"]
                            cell.value = "❌數據缺失"
                            logs.append({"日期": label, "缺失": f"欄位{n_idx+1}真空漏填"})

    output = BytesIO()
    wb.save(output)
    return logs, mode, output.getvalue()

# --- 3. Streamlit 介面啟動區 ---
st.set_page_config(page_title="輕食自主稽核系統", layout="wide")
st.title(ST_TITLE)
st.caption(ST_AUTHOR)

uploaded = st.file_uploader("📂 請上傳菜單檔案 (檔名須包含『輕食』)", type=["xlsx"])

if uploaded:
    logs, detected_mode, excel_data = alison_light_audit(uploaded)
    
    if detected_mode == "BLOCK":
        st.error(f"❌ 拒絕審核：『{uploaded.name}』非輕食區檔案。")
    else:
        st.success(f"✅ 啟動模式：{detected_mode}")
        if logs:
            st.warning(f"🚩 發現 {len(logs)} 處數據真空缺失（已保留 0 值的合格判定）。")
            st.table(pd.DataFrame(logs))
            st.download_button("📥 下載 Alison 標註退件檔", excel_data, f"退件_{uploaded.name}")
        else:
            st.success("🎉 檢查完畢！營養分析數據完整（包含 0 值均已正確識別）。")
