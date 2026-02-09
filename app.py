# ... (前面視覺規範、MEAT_DICT、to_num 全部保持不變) ...

def audit_process(file):
    if "輕食" not in file.name:
        return ["❌ 錯誤：檔名不含『輕食』關鍵字，系統拒絕審核"], None

    try:
        wb = load_workbook(file)
        sheets_df = pd.read_excel(file, sheet_name=None, header=None)
        logs = []
        output = BytesIO()

        # 定義我們要檢查的關鍵字
        TARGET_LABELS = ["熱量", "全榖", "豆魚", "蔬菜", "水果", "油脂", "乳品"]

        for sn, df in sheets_df.items():
            ws = wb[sn]
            d_row = next((i for i, r in df.iterrows() if "日期Date" in str(r[2])), None)
            if d_row is None: continue

            for col in range(3, 8):
                if col >= len(df.columns): break
                date_str = str(df.iloc[d_row, col]).split(" ")[0]
                if "202" not in date_str: continue

                # --- 核心修正：精確掃描營養標示 ---
                for r_idx in range(d_row + 10, len(df)):
                    # 抓取項目名稱 (例如：全榖雜糧類(份))
                    raw_label = str(df.iloc[r_idx, 2]) if not pd.isna(df.iloc[r_idx, 2]) else ""
                    
                    # 【關鍵過濾】：如果這一列的名稱不在我們的檢查目標內，直接跳過！
                    # 這樣就不會再抓到一堆 nan 了
                    if not any(target in raw_label for target in TARGET_LABELS):
                        continue

                    raw_val = df.iloc[r_idx, col]
                    cell = ws.cell(row=r_idx+1, column=col+1)
                    
                    # 1. 判定真空 (只有在標籤存在，但數值空白時才報錯)
                    if pd.isna(raw_val) or str(raw_val).strip() == "":
                        cell.fill = PatternFill(start_color="000000", end_color="000000", fill_type="solid")
                        cell.font = Font(name=FONT_NAME, size=14, color="FFFFFF", bold=True)
                        cell.value = "❌漏填"
                        logs.append({"分頁": sn, "日期": date_str, "項目": raw_label, "原因": "真空漏填"})
                        continue

                    # 2. 判定數值 (0 值會安全過關)
                    val = to_num(raw_val)
                    if "熱量" in raw_label:
                        if val < 750 or val > 850:
                            cell.fill, cell.font = CALORIE_STYLE["fill"], CALORIE_STYLE["font"]
                            logs.append({"分頁": sn, "日期": date_str, "項目": "熱量", "原因": f"粉底：{val} Kcal"})
                    
                    elif any(x in raw_label for x in ["全榖", "豆魚", "蔬菜"]):
                        std = 2.0 if "蔬菜" in raw_label else 4.0
                        # 0 < val < std 代表有填但不足；val == 0 則是妳說的合法數據
                        if 0 < val < std:
                            cell.fill, cell.font = PORTION_STYLE["fill"], PORTION_STYLE["font"]
                            logs.append({"分頁": sn, "日期": date_str, "項目": raw_label, "原因": "紅底白字：份數不足"})

                # ... (食材重複性與禁辣邏輯完全保留) ...
