import pandas as pd
import re
from datetime import datetime
import os

# === 參數設定 ===
file_path = "11411.xlsx"
sheet_name = "醫令統計"
target_col = "數量"   # 如果要改統計「使用量」，只要把這裡改成 "使用量"

# === 讀取 Excel ===
df = pd.read_excel(file_path, sheet_name=sheet_name, engine="openpyxl")

# 1) 欄位清理
df.columns = df.columns.str.strip()
need = ["藥品", "藥代", "廠商", target_col]
missing = [c for c in need if c not in df.columns]
if missing:
    raise ValueError(f"缺少必要欄位：{missing}")

# 2) 去除全空列與欄位字串標準化
def clean_text(s):
    if pd.isna(s):
        return ""
    s = str(s)
    s = s.replace("\u00a0", " ")            # 不換行空格
    s = re.sub(r"[\r\n\t]", " ", s)         # 控制字元
    s = re.sub(r"\s+", " ", s).strip()      # 多重空白
    return s

for col in ["藥品", "藥代", "廠商"]:
    df[col] = df[col].map(clean_text)

# 3) 藥代標準化（大寫）
df["藥代"] = df["藥代"].str.upper()

# 4) 數值化（去逗號/空白/中文字）
def to_num(x):
    if pd.isna(x):
        return 0.0
    s = str(x).strip().replace(",", "")
    s = re.sub(r"[^\d\.\-]", "", s)
    try:
        return float(s) if s != "" else 0.0
    except:
        return 0.0

df[target_col] = df[target_col].map(to_num)

# 5) 去掉顯然不是資料的行
df = df[(df["藥代"] != "") & (df["藥品"] != "")]

# 6) 原始合計核對：ALPR4
mask_alpr4 = df["藥代"].eq("ALPR4")
sum_raw = df.loc[mask_alpr4, target_col].sum()
cnt_raw = mask_alpr4.sum()
print(f"原始資料 ALPR4 筆數：{cnt_raw}，{target_col}合計：{sum_raw}")

# 7) 分組統計（藥品+藥代+廠商）
usage = (
    df.groupby(["藥品", "藥代", "廠商"], as_index=False)[target_col]
      .sum()
      .rename(columns={target_col: f"累計{target_col}"})
)

# 8) 統計表核對：ALPR4
sum_grp = usage.loc[usage["藥代"].eq("ALPR4"), f"累計{target_col}"].sum()
print(f"統計結果 ALPR4 {target_col}合計：{sum_grp}")

# 9) 差異檢查
if abs(sum_grp - sum_raw) > 1e-6:
    print("\n⚠️ 差異偵測：列出 ALPR4 原始前10列供檢查")
    print(df.loc[mask_alpr4, ["藥品", "藥代", "廠商", target_col]].head(10))
    print("\n⚠️ 差異偵測：列出 ALPR4 統計前10列供檢查")
    print(usage.loc[usage["藥代"].eq("ALPR4"), ["藥品", "藥代", "廠商", f"累計{target_col}"]].head(10))

# 10) 排序（依累計值由大到小）
usage = usage.sort_values(by=f"累計{target_col}", ascending=False)

# 10b) 調整欄位順序
usage = usage[["藥代", "藥品", "廠商", f"累計{target_col}"]]

# 11) 輸出 Excel
now = datetime.now().strftime("%Y%m%d_%H%M%S")
out_file = f"藥品藥代廠商統計_醫令統計_{target_col}_{now}.xlsx"
usage.to_excel(out_file, index=False)
print(f"✅ 已存檔：{os.path.abspath(out_file)}")
