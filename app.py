# ... 網頁基本設定與視覺規範 (PORTION_STYLE, CALORIE_STYLE) 一字不差 ...

def audit_process(file):
    # --- 1. 唯一合法性檢查 (Alison 規範) ---
    if "輕食" not in file.name:
        return ["❌ 錯誤：檔名不含『輕食』關鍵字，系統拒絕審核"], None

    try:
        wb = load_workbook(file)
        sheets_df = pd.read_excel(file, sheet_name=None, header=None)
        logs = []
        output = BytesIO()

        for sn, df in sheets_df.items():
            ws = wb[sn]
            # --- 注意：不使用 df.fillna("")，否則會分不出 0 跟真空 ---
            
            d_row = next((i for i, r in df.iterrows() if "日期Date" in str(r[2])), None)
            if d_row is None: continue

            for col in range(3, 8):
                if col >= len(df.columns): break
                date_str = str(df.iloc[d_row, col]).split(" ")[0]

                # --- 2. 營養標示審核 (解決 0 與 空白 衝突的核心) ---
                for r_idx in range(d_row + 10, len(df)):
                    label = str(df.iloc[r_idx, 2])
                    # 抓取原始格位數據，不要先轉數字！
                    raw_val = df.iloc[r_idx, col] 
                    cell = ws.cell(row=r_idx+1, column=col+1)

                    # --- 關鍵：判斷是否為「真空空白」 ---
                    # 只有真的什麼都沒填 (NaN 或 空字串) 才算缺失
                    if pd.isna(raw_val) or str(raw_val).strip() == "":
                        cell.fill, cell.font = PORTION_STYLE["fill"], PORTION_STYLE["font"]
                        cell.value = "❌漏填數據"
                        logs.append({"分頁": sn, "日期": date_str, "項目": label, "原因": "真空空白漏填"})
                        continue # 是真空就噴色跳過，不准往下跑數字判定

                    # --- 如果不是真空 (可能是 0, 0.1, 800...) 才進這裡 ---
                    val = to_num(raw_val)
                    
                    if "熱量" in label and (val < 750 or val > 850):
                        cell.fill, cell.font = CALORIE_STYLE["fill"], CALORIE_STYLE["font"]
                        logs.append({"分頁": sn, "日期": date_str, "項目": "熱量", "原因": f"粉底：{val} Kcal"})
                    
                    elif any(x in label for x in ["全榖", "豆魚", "蔬菜"]):
                        std = 2.0 if "蔬菜" in label else 4.0
                        # 只有當「有填數字且非 0」且「低於標準」才噴色
                        # 這樣寫 0 就會被視為合法數據，不會被噴色！
                        if 0 < val < std: 
                            cell.fill, cell.font = PORTION_STYLE["fill"], PORTION_STYLE["font"]
                            logs.append({"分頁": sn, "日期": date_str, "項目": label, "原因": "份數不足"})

        # ...其餘重複性審核與禁辣邏輯完全保留...
