def audit_process(file):
    # --- 1. 關鍵字攔截：沒"輕食"就不准跑 ---
    if "輕食" not in file.name:
        return ["❌ 錯誤：檔案非『輕食』類別，拒絕審核"], None

    wb = load_workbook(file)
    sheets_df = pd.read_excel(file, sheet_name=None, header=None)
    logs = []
    
    for sn, df in sheets_df.items():
        df = df.fillna("") # 這裡只是輔助，不影響原始數值判斷
        ws = wb[sn]
        
        # ... 定位日期 d_row 邏輯不變 ...

        for col in range(3, 8):
            # 營養標示審核區
            for r_idx in range(d_row + 10, len(df)):
                label = str(df.iloc[r_idx, 2])
                raw_value = df.iloc[r_idx, col] # 抓取原始格位，不轉數字
                
                # --- 核心修正：先看是不是「真空」 ---
                if str(raw_value).strip() == "":
                    # 只有真空空白才噴黑 (PORTION_STYLE 或者是妳指定的漏填樣式)
                    cell = ws.cell(row=r_idx+1, column=col+1)
                    cell.fill, cell.font = PORTION_STYLE["fill"], PORTION_STYLE["font"]
                    cell.value = "❌漏填"
                    logs.append({"分頁": sn, "項目": label, "原因": "真空空白缺失"})
                    continue # 跳過，不進數字判定
                
                # --- 如果不是空白(包含 0)，才進數字判定 ---
                val = to_num(raw_value)
                if "熱量" in label and (val < 750 or val > 850):
                    # 噴粉底樣式...
                elif any(x in label for x in ["全榖", "豆魚", "蔬菜"]):
                    std = 2.0 if "蔬菜" in label else 4.0
                    # 如果 val 是 0，它是有數據的，所以不應該因為 < std 就噴紅底
                    # 除非妳的規則是「0 也算份數不足」
                    if val < std and val != 0: 
                        # 噴紅底樣式...
