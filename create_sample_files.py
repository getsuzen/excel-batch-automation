from pathlib import Path

from openpyxl import Workbook


INPUT_DIR = Path("input")


def create_excel(filename, rows):
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Data"

    for row in rows:
        worksheet.append(row)

    workbook.save(INPUT_DIR / filename)
    workbook.close()


def main():
    INPUT_DIR.mkdir(exist_ok=True)

    header = ["ID", "Name", "Department", "Amount"]

    create_excel(
        "sales_01.xlsx",
        [
            header,
            [1, "Alice", "Tokyo", 1200],
            [2, "Bob", "Osaka", 1500],
        ],
    )

    create_excel(
        "sales_02.xlsx",
        [
            header,
            [3, "Charlie", "Nagoya", 1800],
            [4, "David", "Tokyo", 2100],
        ],
    )

    create_excel(
        "sales_03.xlsx",
        [
            header,
            [5, "Emma", "Osaka", 1700],
            [6, "Frank", "Fukuoka", 1600],
            [1, "Alice", "Tokyo", 1200],  # sales_01.xlsx の1行と重複
        ],
    )

    print("テスト用 Excel ファイルの生成が完了しました（全7行のデータ、うち1行が重複）。")


if __name__ == "__main__":
    main()
