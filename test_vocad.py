import pandas as pd
from datetime import datetime

# 讀取 Excel
file_path = "藥品藥代廠商統計_醫令統計_數量_20251211_114629.xlsx"
sheet_name = "Sheet1"
df = pd.read_excel(file_path, sheet_name=sheet_name, engine="openpyxl")

# 清理欄位名稱
df.columns = df.columns.str.strip()
df.rename(columns={"累計數量": "累計用量"}, inplace=True)

# 確保數值正確
df["累計用量"] = pd.to_numeric(df["累計用量"], errors="coerce").fillna(0)
df["廠商"] = df["廠商"].fillna("未標示廠商")

# 廠商前2碼合併
df["廠商代碼"] = df["廠商"].str[:2]

# 依廠商代碼統計總累計用量
summary = (
    df.groupby("廠商代碼", as_index=False)["累計用量"]
    .sum()
    .rename(columns={"累計用量": "總累計用量"})
)

# 先列出所有廠商代碼
print("\n🏭 可選擇的廠商代碼清單：")
for i, v in enumerate(summary["廠商代碼"], 1):
    print(f"{i}. {v} (累計用量:{summary.loc[i-1,'總累計用量']:.1f})")

# 使用者選擇廠商代碼
choice = input("\n請輸入要查詢的廠商代碼編號：").strip()
try:
    choice_idx = int(choice) - 1
    vendor_code = summary.loc[choice_idx, "廠商代碼"]
except:
    raise ValueError("輸入錯誤，請輸入正確的廠商代碼編號")

# 篩選該廠商代碼的藥品
vendor_data = df[df["廠商代碼"] == vendor_code].copy()

print(f"\n📋 廠商代碼 {vendor_code} 的藥品清單：")
stocks = []
for _, r in vendor_data.iterrows():
    prompt = f"{r['藥品'][:40]:<40} 累計用量:{r['累計用量']:.1f} → 庫存："
    val = input(prompt).strip()
    try:
        stocks.append(int(val) if val else 0)
    except ValueError:
        stocks.append(0)

vendor_data["庫存"] = stocks
vendor_data["缺口"] = (vendor_data["累計用量"] - vendor_data["庫存"]).clip(lower=0)
vendor_data["需採購"] = vendor_data["庫存"] < vendor_data["累計用量"]

# 顯示結果
print(f"\n🏭 廠商代碼 {vendor_code} 盤點結果：")
header = f"{'藥品名稱':<40} {'累計用量':>12} {'庫存':>8} {'缺口':>8} {'需採購':>8}"
print(header)
print("-" * len(header))
for _, r in vendor_data.iterrows():
    print(f"{r['藥品'][:40]:<40} {r['累計用量']:>12.1f} {int(r['庫存']):>8} {int(r['缺口']):>8} {str(r['需採購']):>8}")

# 存檔
now = datetime.now().strftime("%Y%m%d_%H%M%S")
out_file = f"{vendor_code}_盤點結果_{now}.xlsx"
vendor_data.to_excel(out_file, index=False)
print(f"\n✅ 結果已儲存：{out_file}")
