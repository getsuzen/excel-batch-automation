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
    """判断一整行是否为空。"""
    return all(
        value is None or (isinstance(value, str) and not value.strip())
        for value in row
    )


def read_excel(file_path):
    """读取 Excel 第一张工作表，返回表头和数据行。"""
    workbook = None
    try:
        workbook = load_workbook(file_path, read_only=True, data_only=True)
        if not workbook.worksheets:
            print(f"跳过空文件: {file_path.name}（没有工作表）")
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
        print(f"跳过损坏或无法读取的文件: {file_path.name}（{error}）")
        return None
    finally:
        if workbook is not None:
            workbook.close()

    if not rows:
        print(f"跳过空文件: {file_path.name}（第一张工作表没有内容）")
        return None

    return rows[0], rows[1:]


def format_output(worksheet):
    """给结果表添加少量基础格式，便于展示和阅读。"""
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
    """把合并结果保存到 output/summary.xlsx。"""
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
        print(f"保存失败: {OUTPUT_FILE}（{error}）")
        return False
    finally:
        output_workbook.close()

    return True


def warn_if_old_output_exists():
    if OUTPUT_FILE.exists():
        print(f"注意: 已有的 {OUTPUT_FILE} 是之前生成的，本次没有更新。")


def main():
    if not INPUT_DIR.exists():
        print(f"未找到输入目录: {INPUT_DIR}")
        warn_if_old_output_exists()
        return 1

    excel_files = sorted(
        file_path
        for file_path in INPUT_DIR.glob("*.xlsx")
        if not file_path.name.startswith("~$")
    )

    if not excel_files:
        print(f"{INPUT_DIR} 目录中没有 .xlsx 文件。")
        warn_if_old_output_exists()
        return 1

    header = None
    unique_rows = []
    seen_rows = set()
    original_count = 0
    valid_file_count = 0

    for file_path in excel_files:
        print(f"正在读取: {file_path.name}")
        result = read_excel(file_path)
        if result is None:
            continue

        current_header, data_rows = result

        if header is None:
            header = current_header
        elif current_header != header:
            print(f"跳过表头不一致的文件: {file_path.name}")
            print(f"  期望表头: {list(header)}")
            print(f"  实际表头: {list(current_header)}")
            continue

        valid_file_count += 1
        if not data_rows:
            print(f"提示: {file_path.name} 只有表头，没有数据行。")

        original_count += len(data_rows)
        for row in data_rows:
            if row not in seen_rows:
                seen_rows.add(row)
                unique_rows.append(row)

    if header is None:
        print("没有找到可合并的 Excel 内容，未生成新输出文件。")
        warn_if_old_output_exists()
        return 1

    if not save_result(header, unique_rows):
        return 1

    unique_count = len(unique_rows)
    duplicate_count = original_count - unique_count

    print()
    print(f"有效文件数: {valid_file_count}/{len(excel_files)}")
    print(f"原始行数: {original_count}")
    print(f"去重后行数: {unique_count}")
    print(f"删除重复数: {duplicate_count}")
    print(f"完成！输出文件: {OUTPUT_FILE}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
