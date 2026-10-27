from pathlib import Path
from zipfile import BadZipFile
from xml.etree.ElementTree import ParseError

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.utils.exceptions import InvalidFileException


INPUT_DIR = Path("input")
OUTPUT_FILE = Path("output") / "summary.xlsx"


def is_empty_row(row):
    """行全体が空かどうかを判定する。"""
    return all(
        value is None or (isinstance(value, str) and not value.strip())
        for value in row
    )


def read_excel(file_path):
    """Excel の先頭シートを読み込み、ヘッダーとデータ行を返す。"""
    workbook = None
    try:
        workbook = load_workbook(file_path, read_only=True, data_only=True)
        if not workbook.worksheets:
            print(f"空ファイルのためスキップ: {file_path.name}（シートがありません）")
            return None

        worksheet = workbook.active
        rows = [
            tuple(row)
            for row in worksheet.iter_rows(values_only=True)
            if not is_empty_row(row)
        ]
    except (
        InvalidFileException,
        BadZipFile,
        OSError,
        ValueError,
        KeyError,
        EOFError,
        ParseError,
    ) as error:
        print(f"破損または読み取れないファイルのためスキップ: {file_path.name}（{error}）")
        return None
    finally:
        if workbook is not None:
            workbook.close()

    if not rows:
        print(f"空ファイルのためスキップ: {file_path.name}（先頭シートに内容がありません）")
        return None

    return rows[0], rows[1:]


def format_output(worksheet):
    """結果シートに表示・閲覧用の基本書式を適用する。"""
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)

    for cell in worksheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions

    for column_cells in worksheet.columns:
        max_length = max(
            len(str(cell.value)) if cell.value is not None else 0
            for cell in column_cells
        )
        column_letter = get_column_letter(column_cells[0].column)
        worksheet.column_dimensions[column_letter].width = min(max(max_length + 2, 10), 30)


def save_result(header, rows):
    """統合結果を output/summary.xlsx に保存する。"""
    output_workbook = Workbook()
    output_sheet = output_workbook.active
    output_sheet.title = "Summary"

    output_sheet.append(list(header))
    for row in rows:
        output_sheet.append(list(row))

    format_output(output_sheet)

    try:
        OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
        output_workbook.save(OUTPUT_FILE)
    except OSError as error:
        print(f"保存に失敗しました: {OUTPUT_FILE}（{error}）")
        return False
    finally:
        output_workbook.close()

    return True


def warn_if_old_output_exists():
    if OUTPUT_FILE.exists():
        print(f"注意：既存の {OUTPUT_FILE} は前回生成されたもので、今回は更新されていません。")


def main():
    if not INPUT_DIR.exists():
        print(f"入力フォルダーが見つかりません: {INPUT_DIR}")
        warn_if_old_output_exists()
        return 1

    excel_files = sorted(
        file_path
        for file_path in INPUT_DIR.glob("*.xlsx")
        if not file_path.name.startswith("~$")
    )

    if not excel_files:
        print(f"{INPUT_DIR} フォルダー内に .xlsx ファイルがありません。")
        warn_if_old_output_exists()
        return 1

    header = None
    unique_rows = []
    seen_rows = set()
    original_count = 0
    valid_file_count = 0

    for file_path in excel_files:
        print(f"読み込み中: {file_path.name}")
        result = read_excel(file_path)
        if result is None:
            continue

        current_header, data_rows = result

        if header is None:
            header = current_header
        elif current_header != header:
            print(f"ヘッダーが一致しないためスキップ: {file_path.name}")
            print(f"  期待されるヘッダー: {list(header)}")
            print(f"  実際のヘッダー: {list(current_header)}")
            continue

        valid_file_count += 1
        if not data_rows:
            print(f"ヒント： {file_path.name} はヘッダーのみで、データ行がありません。")

        original_count += len(data_rows)
        for row in data_rows:
            if row not in seen_rows:
                seen_rows.add(row)
                unique_rows.append(row)

    if header is None:
        print("統合可能な Excel データが見つからず、新しい出力ファイルは生成されませんでした。")
        warn_if_old_output_exists()
        return 1

    if not save_result(header, unique_rows):
        return 1

    unique_count = len(unique_rows)
    duplicate_count = original_count - unique_count

    print()
    print(f"有効ファイル数: {valid_file_count}/{len(excel_files)}")
    print(f"統合前の行数: {original_count}")
    print(f"統合後の行数: {unique_count}")
    print(f"除外した重複行数: {duplicate_count}")
    print(f"完了！出力ファイル: {OUTPUT_FILE}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
