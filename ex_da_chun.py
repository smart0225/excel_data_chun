import pandas as pd
import numpy as np
import os
import logging
import re
import tempfile
from pathlib import Path

from openpyxl import load_workbook


LOGGER = logging.getLogger(__name__)


def _safe_sheet_filename(sheet_name):
    """Return a sheet name that is safe to use in a Windows filename."""
    filename = re.sub(r'[<>:"/\\|?*]', "_", sheet_name).rstrip(" .")
    return filename or "sheet"


def _keep_first_xlsx_sheet(source_path, temporary_path):
    workbook = load_workbook(source_path)
    try:
        for sheet_name in workbook.sheetnames[1:]:
            del workbook[sheet_name]
        workbook.save(temporary_path)
    finally:
        workbook.close()


def _keep_first_xls_sheet(source_path, temporary_path):
    import xlrd
    import xlwt

    source_workbook = xlrd.open_workbook(source_path)
    try:
        source_sheet = source_workbook.sheet_by_index(0)
        workbook = xlwt.Workbook(encoding="utf-8")
        sheet = workbook.add_sheet(source_sheet.name)
        date_style = xlwt.easyxf(num_format_str="YYYY-MM-DD HH:MM:SS")

        for row_index in range(source_sheet.nrows):
            for column_index in range(source_sheet.ncols):
                cell = source_sheet.cell(row_index, column_index)
                if cell.ctype == xlrd.XL_CELL_EMPTY:
                    continue
                if cell.ctype == xlrd.XL_CELL_DATE:
                    value = xlrd.xldate_as_datetime(
                        cell.value, source_workbook.datemode
                    )
                    sheet.write(row_index, column_index, value, date_style)
                elif cell.ctype == xlrd.XL_CELL_ERROR:
                    sheet.write(
                        row_index,
                        column_index,
                        xlrd.error_text_from_code.get(cell.value, str(cell.value)),
                    )
                else:
                    sheet.write(row_index, column_index, cell.value)

        workbook.save(str(temporary_path))
    finally:
        source_workbook.release_resources()


def split_workbooks(directory):
    """Export each sheet after the first as a values-only .xlsx workbook."""
    directory_path = Path(directory)
    if not directory_path.is_dir():
        raise NotADirectoryError(f"Not a directory: {directory_path}")

    failures = []
    workbook_count = 0
    split_workbook_count = 0
    output_file_count = 0
    for source_path in sorted(directory_path.iterdir()):
        if (
            not source_path.is_file()
            or source_path.suffix.lower() not in {".xls", ".xlsx"}
            or source_path.name.startswith("~$")
        ):
            continue

        workbook_count += 1
        engine = "xlrd" if source_path.suffix.lower() == ".xls" else "openpyxl"
        try:
            with tempfile.TemporaryDirectory(dir=directory_path) as temp_dir:
                temp_dir_path = Path(temp_dir)
                staged_outputs = []
                with pd.ExcelFile(source_path, engine=engine) as excel_file:
                    sheet_names = excel_file.sheet_names
                    if len(sheet_names) <= 1:
                        continue

                    output_paths = [
                        directory_path
                        / f"{source_path.stem} - {_safe_sheet_filename(sheet_name)}.xlsx"
                        for sheet_name in sheet_names[1:]
                    ]
                    normalized_paths = [
                        str(path).casefold() for path in output_paths
                    ]
                    if len(normalized_paths) != len(set(normalized_paths)):
                        raise FileExistsError(
                            "Sheet names produce duplicate output filenames: "
                            f"{source_path}"
                        )
                    existing_output = next(
                        (path for path in output_paths if path.exists()), None
                    )
                    if existing_output is not None:
                        raise FileExistsError(
                            f"Output already exists: {existing_output}"
                        )

                    for index, sheet_name in enumerate(sheet_names[1:], start=1):
                        frame = pd.read_excel(
                            excel_file,
                            sheet_name=sheet_name,
                            engine=engine,
                        )
                        staged_path = temp_dir_path / f"sheet-{index}.xlsx"
                        frame.to_excel(staged_path, index=False, sheet_name=sheet_name)
                        staged_outputs.append(staged_path)

                staged_source = temp_dir_path / source_path.name
                if engine == "openpyxl":
                    _keep_first_xlsx_sheet(source_path, staged_source)
                else:
                    LOGGER.warning(
                        "%s is .xls; keeping its first sheet will not preserve "
                        "formatting or other workbook features.",
                        source_path.name,
                    )
                    _keep_first_xls_sheet(source_path, staged_source)

                created_outputs = []
                try:
                    for staged_path, output_path in zip(staged_outputs, output_paths):
                        os.replace(staged_path, output_path)
                        created_outputs.append(output_path)
                    os.replace(staged_source, source_path)
                except OSError:
                    for output_path in created_outputs:
                        output_path.unlink(missing_ok=True)
                    raise

            LOGGER.info(
                "Split %s into %d workbook(s); kept the first sheet in the source.",
                source_path.name,
                len(output_paths),
            )
            split_workbook_count += 1
            output_file_count += len(output_paths)
        except Exception as error:
            LOGGER.exception("Could not split workbook %s", source_path)
            failures.append(f"{source_path.name}: {error}")

    if failures:
        raise RuntimeError("Some workbooks could not be split:\n" + "\n".join(failures))

    LOGGER.info(
        "處理成功：檢查 %d 個 Excel 檔案，分拆 %d 個活頁簿，建立 %d 個新檔案。",
        workbook_count,
        split_workbook_count,
        output_file_count,
    )

def load_data(file_path):
    """
    Load data from a CSV file into a pandas DataFrame.

    Parameters:
    file_path (str): The path to the CSV file.

    Returns:
    pd.DataFrame: A DataFrame containing the loaded data.
    """
    try:
        data = pd.read_csv(file_path)
        return data
    except Exception as e:
        print(f"Error loading data: {e}")
        return None



# 判斷資料夾檔案是否存在
def check_file_exists(file_path):
    """
    Check if a file exists at the given path.

    Parameters:
    file_path (str): The path to the file.

    Returns:
    bool: True if the file exists, False otherwise.
    """
    return os.path.isfile(file_path)


file_path = r'D:\MyData\BM' 


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    split_workbooks(file_path)