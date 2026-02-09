def audit_process(file):
    # --- 【交叉比對修正 1】：身分攔截 ---
    # 這是妳最在意的：沒"輕食"關鍵字，絕對不准進場審核
    if "輕食" not in file.name:
        return ["❌ 錯誤：檔名未包含『輕食』，系統拒絕誤審！"], None

    try:
        wb = load_workbook(file)
        sheets_df = pd.read_excel(file, sheet_name=None, header=None)
        logs = []
        output = BytesIO()

        for sn, df in sheets_df.items():
            ws = wb[sn]
            # 注意：這裡不准先做 fillna("")，否則會分不出 0 跟真空
            
            d_row = next((i for i, r in df.iterrows() if "日期Date" in str(r[2])), None)
            if d_row is None: continue

            for col in range(3, 8): # 週一到週五
                if col >= len(df.columns): break
                date_str = str(df.iloc[d_row, col]).split(" ")[0]
                
                # --- 【交叉比對修正 2】：營養標示「真空」判定 ---
                # 輕食區專用座標：d_row + 10 到最後
                for r_idx in range(d_row + 10, len(df)):
                    label = str(df.iloc[r_idx, 2])
                    raw_val = df.iloc[r_idx, col] # 抓取原始格位
                    cell = ws.cell(row=r_idx+1, column=col+1)

                    # A. 優先檢查「真空」 (這是妳抓到的 BUG：0 不是空白)
                    # pd.isna 會精確抓出「沒填」，而不會抓到「填 0」
                    if pd.isna(raw_val) or str(raw_val).strip() == "":
                        cell.fill, cell.font = PORTION_STYLE["fill"], PORTION_STYLE["font"]
                        cell.value = "❌漏填數據"
                        logs.append({"分頁": sn, "日期": date_str, "項目": label, "原因": "真空空白漏填"})
                        continue # 真空直接噴黑，不參與後續數值判斷

                    # B. 如果不是真空 (不管是 0 還是數字)，才進妳原本的數值規則
                    val = to_num(raw_val)
                    if "熱量" in label and (val < 750 or val > 850):
                        cell.fill, cell.font = CALORIE_STYLE["fill"], CALORIE_STYLE["font"]
                        logs.append({"分頁": sn, "日期": date_str, "項目": "熱量", "原因": f"異常：{val} Kcal"})
                    
                    elif any(x in label for x in ["全榖", "豆魚", "蔬菜"]):
                        std = 2.0 if "蔬菜" in label else 4.0
                        # 只有當 val < 標準 且 不等於 0 時才判定不足 (如果妳認為 0 是合理數據)
                        if val < std and val != 0:
                            cell.fill, cell.font = PORTION_STYLE["fill"], PORTION_STYLE["font"]
                            logs.append({"分頁": sn, "日期": date_str, "項目": label, "原因": "份數不足"})

        # ...其餘重複性審核與禁辣邏輯不變...
