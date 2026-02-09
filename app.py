import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# --- 1. 標題與基礎設定 (Alison 規範) ---
ST_TITLE = "🛡️ 團膳區(新北食品) 菜單自主稽核系統"
ST_AUTHOR = "製作者：Alison"
STYLE_ERR = {
    "fill": PatternFill("solid", fgColor="000000"), 
    "font": Font(name="微軟正黑體", size=12, color="FFFFFF", bold=True)
}

def alison_complete_audit(file):
    fname = file.name
    mode = None
    
    # --- 2. 模式判定 (包含新北幼兒園、小學、暖禾輕食) ---
    if any(kw in fname for kw in ["小學", "幼兒園", "幼兒"]):
        mode = "新北食品-教育學部"
        nutri_indices = [9, 10, 11, 12, 13, 14, 15] # J-P 欄
        data_indices = [1, 2, 3, 4, 5, 6, 7]
    elif any(kw in fname for kw in ["輕食", "暖禾"]):
        mode = "新北食品-輕食專區(暖禾)"
        # 暖禾輕食座標校準：通常在 B-C 為主項，F-L 為營養分析
        nutri_indices = [5, 6, 7, 8, 9, 10, 11] 
        data_indices = [1, 2]
    elif any(kw in fname for kw in ["美食街", "素食"]):
        mode = "新北食品-美食/素食區"
        nutri_indices = [3, 4, 5, 6, 7]
        data_indices = [3, 4, 5, 6, 7]
    
    if mode is None:
        return None, "BLOCK", None

    # --- 3. 執行稽核 ---
    wb = load_workbook(file)
    sheets_df = pd.read_excel(file, sheet_name=None, header=None)
    logs = []

    for sn, df in sheets_df.items():
        ws = wb[sn]
        # 只處理 NaN，保留 0 和 0.0
        df_audit = df.astype(str).replace(['nan', 'NaN', 'None'], '')
        
        for r_idx in range(len(df_audit)):
            label = str(df_audit.iloc[r_idx, 0]).strip()
            
            # 日期行判定 (格式如: 1/05 (一))
            if "/" in label and "(" in label:
                # A. 營養成分全檢 (Alison 鐵律: 0 不是空白)
                has_content = any(df_audit.iloc[r_idx, c].strip() != "" for c in data_indices if c < len(df_audit.columns))
                
                if has_content:
                    for n_idx in nutri_indices:
                        if n_idx >= len(df_audit.columns): continue
                        val = df_audit.iloc[r_idx, n_idx].strip()
                        
                        # 只有當格完全沒有字(真空)時才噴黑
                        if val == "":
                            cell = ws.cell(row=r_idx+1, column=n_idx+1)
                            cell.fill, cell.font = STYLE_ERR["fill"], STYLE_ERR["font"]
                            cell.value = "❌漏填數據"
                            logs.append({"分頁": sn, "日期": label, "缺失": f"欄位 {n_idx+1} 營養數據漏填"})

                # B. 垂直菜名黑洞檢查 (防止明細有字但菜名空白)
                for d_idx in data_indices:
                    if d_idx >= len(df_audit.columns): continue
                    if df_audit.iloc[r_idx, d_idx] == "":
                        try:
                            # 檢查下一列(明細列)是否有字
                            if df_audit.iloc[r_idx+1, d_idx] != "":
                                cell = ws.cell(row=r_idx+1, column=d_idx+1)
                                cell.fill, cell.font = STYLE_ERR["fill"], STYLE_ERR["font"]
                                cell.value = "❌漏填菜名"
                                logs.append({"分頁": sn, "日期": label, "缺失": "有明細無菜名"})
                        except: pass

    output = BytesIO()
    wb.save(output)
    return logs, mode, output.getvalue()

# --- 4. Streamlit 介面 (Alison 專屬 UI) ---
st.set_page_config(page_title="新北食品稽核系統", layout="wide")
st.title(ST_TITLE)
st.caption(ST_AUTHOR)

uploaded = st.file_uploader("📂 上傳菜單 Excel (支援暖禾輕食/小學/幼兒園/美食街)", type=["xlsx"])

if uploaded:
    logs, detected_mode, excel_file = alison_complete_audit(uploaded)
    
    if detected_mode == "BLOCK":
        st.error(f"❌ 拒絕審核：『{uploaded.name}』檔名不符規範。請確認包含關鍵字。")
    else:
        st.info(f"✅ 模式識別成功：{detected_mode}")
        if logs:
            st.warning(f"🚩 發現 {len(logs)} 處數據真空缺失 (不包含數值為 0 之項目)")
            st.table(pd.DataFrame(logs))
            st.download_button("📥 下載 Alison 標註退件檔", excel_file, f"退件_{uploaded.name}")
        else:
            st.success("🎉 完美！所有數值（含 0）填寫均完整且對位正確。")
