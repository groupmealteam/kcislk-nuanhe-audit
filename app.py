# ... (前面的視覺規範與 MEAT_DICT 一字不差) ...

def audit_process(file):
    # --- 【修正 1】：嚴格攔截 (Alison 要求的認真) ---
    if "輕食" not in file.name:
        return ["❌ 錯誤：檔名不含『輕食』關鍵字，系統拒絕審核"], None

    try:
        wb = load_workbook(file)
        sheets_df = pd.read_excel(file, sheet_name=None, header=None)
        logs = []
        output = BytesIO()

        for sn, df in sheets_df.items():
            # 注意：這裡要把 fillna 拿掉，改在下面判斷，否則分不出 0 和真空
            ws = wb[sn]
            d_row = next((i for i, r in df.iterrows() if "日期Date" in str(r[2])), None)
            if d_row is None: continue

            for col in range(3, 8): 
                if col >= len(df.columns): break
                date_str = str(df.iloc[d_row, col]).split(" ")[0]
                if "202" not in date_str: continue
                day_name = str(df.iloc[d_row+1, col]) if (d_row+1) < len(df) else ""

                # --- 【修正 2】：營養標示審核 (精確處理 0 與 真空) ---
                for r_idx in range(d_row + 10, len(df)):
                    label = str(df.iloc[r_idx, 2])
                    raw_val = df.iloc[r_idx, col] # 抓取原始內容
                    cell = ws.cell(row=r_idx+1, column=col+1)
                    
                    # 1. 先判斷「真空」：如果是空的 (NaN 或 空字串)
                    if pd.isna(raw_val) or str(raw_val).strip() == "":
                        # 噴黑底白字 (妳指定針對漏填的噴色)
                        cell.fill = PatternFill(start_color="000000", end_color="000000", fill_type="solid")
                        cell.font = Font(name=FONT_NAME, size=FONT_SIZE, color="FFFFFF", bold=True)
                        cell.value = "❌漏填"
                        logs.append({"分頁": sn, "日期": date_str, "項目": label, "原因": "真空漏填"})
                        continue # 跳過，不進後續數字判斷

                    # 2. 如果不是空的 (可能是 0，或者是正常數字)
                    val = to_num(raw_val)
                    if "熱量" in label and (val < 750 or val > 850):
                        cell.fill, cell.font = CALORIE_STYLE["fill"], CALORIE_STYLE["font"]
                        logs.append({"分頁": sn, "日期": date_str, "項目": "熱量", "原因": f"粉底：{val} Kcal"})
                    elif any(x in label for x in ["全榖", "豆魚", "蔬菜"]):
                        std = 2.0 if "蔬菜" in label else 4.0
                        # 如果 val 是 0，它是廠商填的數據，不算份數不足 (依妳的 3/27 邏輯)
                        if 0 < val < std:
                            cell.fill, cell.font = PORTION_STYLE["fill"], PORTION_STYLE["font"]
                            logs.append({"分頁": sn, "日期": date_str, "項目": label, "原因": "紅底白字：份數不足"})

                # --- 3. 食材重複性與禁辣 (原本很正常的邏輯完全保留) ---
                # ... (以下代碼與妳貼的一模一樣) ...
