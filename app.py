def audit_process(file):
    # --- 關鍵修正 1：檔名攔截 (沒關鍵字不准審) ---
    if "輕食" not in file.name:
        return ["❌ 錯誤：檔名不含『輕食』關鍵字，系統拒絕審核"], None

    try:
        wb = load_workbook(file)
        sheets_df = pd.read_excel(file, sheet_name=None, header=None)
        logs = []
        output = BytesIO()

        for sn, df in sheets_df.items():
            df = df.fillna("") # 將 NaN 轉為 ""
            ws = wb[sn]
            d_row = next((i for i, r in df.iterrows() if "日期Date" in str(r[2])), None)
            if d_row is None: continue

            for col in range(3, 8):
                if col >= len(df.columns): break
                
                # --- 關鍵修正 2：區分 0 與 空白 ---
                for r_idx in range(d_row + 10, len(df)):
                    label = str(df.iloc[r_idx, 2])
                    raw_val = str(df.iloc[r_idx, col]).strip() # 原始字串
                    val = to_num(raw_val) # 轉成數字
                    cell = ws.cell(row=r_idx+1, column=col+1)
                    
                    # A. 檢查熱量
                    if "熱量" in label:
                        if raw_val == "": # 真空空白
                            cell.fill, cell.font = PORTION_STYLE["fill"], PORTION_STYLE["font"] # 噴黑底或紅底
                            logs.append({"分頁": sn, "日期": "---", "項目": "熱量", "原因": "真空漏填"})
                        elif val < 750 or val > 850:
                            cell.fill, cell.font = CALORIE_STYLE["fill"], CALORIE_STYLE["font"]
                            logs.append({"分頁": sn, "日期": "---", "項目": "熱量", "原因": f"數據異常：{val}"})

                    # B. 檢查份數 (全榖/豆魚/蔬菜)
                    elif any(x in label for x in ["全榖", "豆魚", "蔬菜"]):
                        std = 2.0 if "蔬菜" in label else 4.0
                        if raw_val == "": # 只有真空才報漏填
                             cell.fill, cell.font = PORTION_STYLE["fill"], PORTION_STYLE["font"]
                             logs.append({"分頁": sn, "項目": label, "原因": "真空漏填"})
                        elif val < std: # 如果是 0，則按標準判定是否不足
                             # 如果妳的標準是 0 算合格，這裡就要加 if val != 0
                             if val != 0: 
                                 cell.fill, cell.font = PORTION_STYLE["fill"], PORTION_STYLE["font"]
                                 logs.append({"分頁": sn, "項目": label, "原因": "份數不足"})
