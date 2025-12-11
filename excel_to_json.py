mport os
import json
import pandas as pd

# 你的 Excel 資料夾（你提供的路徑）
FOLDER_PATH = r"D:\User\Desktop\purchase"

# 輸出 JSON 檔案名稱
OUTPUT_JSON = "merged.json"


def read_all_excels(folder_path):
    # 檢查資料夾是否存在
    if not os.path.exists(folder_path):
        print("❌ 路徑不存在：", folder_path)
        return []

    # 找所有 Excel 檔案
    excel_files = [
        f for f in os.listdir(folder_path)
        if f.lower().endswith(".xlsx") or f.lower().endswith(".xls")
    ]

    if not excel_files:
        print("❌ 沒找到任何 Excel (.xlsx/.xls)")
        return []

    print("📄 找到 Excel：", excel_files)

    data = []

    for filename in excel_files:
        file_path = os.path.join(folder_path, filename)
        print(f"📂 讀取：{file_path}")

        try:
            xls = pd.ExcelFile(file_path)  # 讀全部 sheet
        except Exception as e:
            print("⚠ 無法讀取：", file_path)
            print("原因：", e)
            continue

        for sheet in xls.sheet_names:
            df = pd.read_excel(file_path, sheet_name=sheet)
            data.append({
                "file": filename,
                "sheet": sheet,
                "rows": df.to_dict(orient="records")
            })

    return data


def main():
    all_data = read_all_excels(FOLDER_PATH)
    if not all_data:
        return  # 沒讀到資料則停止

    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(all_data, f, ensure_ascii=False, indent=2)

    print("\n✔ 完成！已輸出：", OUTPUT_JSON)


if __name__ == "__main__":
    main()

