import pandas as pd
import json
import os
import sys
import re

# 欄位對應表：原始欄位名稱 → 標準欄位名稱
COLUMN_MAP = {
    "藥代": "藥品代碼",
    "藥代碼": "藥品代碼",
    "代碼": "藥品代碼",
    "code": "藥品代碼",

    "藥品": "藥品名稱",
    "品名": "藥品名稱",
    "name": "藥品名稱",

    "廠商": "廠商",
    "供應商": "廠商",
    "製造商": "廠商",
    "小廠": "廠商",
    "vendor": "廠商",

    "累計數量": "盤點數量",
    "數量": "盤點數量",
    "qty": "盤點數量"
}

def normalize_columns(df):
    """將欄位名稱標準化"""
    new_columns = {}
    for col in df.columns:
        col_clean = str(col).strip()
        new_columns[col] = COLUMN_MAP.get(col_clean, col_clean)
    return df.rename(columns=new_columns)

def extract_date_from_filename(filename):
    """從檔名中提取日期（格式：YYYYMMDD → YYYY-MM-DD）"""
    match = re.search(r"(\d{8})", filename)
    if match:
        raw = match.group(1)
        return f"{raw[:4]}-{raw[4:6]}-{raw[6:]}"
    return None

def fill_missing_vendor(df, default_vendor="未填廠商"):
    """補上缺漏或空白的廠商欄位"""
    if "廠商" not in df.columns:
        df["廠商"] = default_vendor
    else:
        df["廠商"] = df["廠商"].fillna(default_vendor)
        df.loc[df["廠商"].astype(str).str.strip() == "", "廠商"] = default_vendor
    return df

def add_inventory_date(df, date_str):
    """加入盤點日期欄位"""
    if date_str:
        df["盤點日期"] = date_str
    return df

def excel_to_json(excel_file, output_file=None):
    # 讀取 Excel
    df = pd.read_excel(excel_file)

    # 標準化欄位名稱
    df = normalize_columns(df)

    # 補上廠商欄位
    df = fill_missing_vendor(df)

    # 從檔名提取日期並加入盤點日期欄位
    date_str = extract_date_from_filename(excel_file)
    df = add_inventory_date(df, date_str)

    # 轉成 JSON
    records = df.to_dict(orient="records")

    # 設定輸出檔名
    if output_file is None:
        base = os.path.splitext(excel_file)[0]
        output_file = base + ".json"

    # 寫入 JSON 檔案
    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=2)

    print(f"✅ 已轉換完成：{output_file}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("❗ 用法：python excel_to_json.py <Excel檔案路徑> [輸出檔名]")
    else:
        excel_file = sys.argv[1]
        output_file = sys.argv[2] if len(sys.argv) > 2 else None
        excel_to_json(excel_file, output_file)

        print("❗ 用法：python excel_to_json.py <Excel檔案路徑> [輸出檔名]")
    else:
        excel_file = sys.argv[1]
        output_file = sys.argv[2] if len(sys.argv) > 2 else None
        excel_to_json(excel_file, output_file)
