from openpyxl.styles import Alignment, Font
from openpyxl.worksheet.worksheet import Worksheet


def format_all_columns_with_hard_coded(work_sheet: Worksheet):
    work_sheet.column_dimensions["A"].width = 4.4
    work_sheet.column_dimensions["B"].width = 10.019
    work_sheet.column_dimensions["C"].width = 11.73
    work_sheet.column_dimensions["D"].width = 15.274
    work_sheet.column_dimensions["E"].width = 11.73
    work_sheet.column_dimensions["F"].width = 11.119
    work_sheet.column_dimensions["G"].width = 2.685
    work_sheet.column_dimensions["H"].width = 7.085
    work_sheet.column_dimensions["I"].width = 20.53
    work_sheet.column_dimensions["J"].width = 2.93


def check_all_columns_width(work_sheet: Worksheet):
    print(work_sheet.column_dimensions["A"].width)
    print(work_sheet.column_dimensions["B"].width)
    print(work_sheet.column_dimensions["C"].width)
    print(work_sheet.column_dimensions["D"].width)
    print(work_sheet.column_dimensions["E"].width)
    print(work_sheet.column_dimensions["F"].width)
    print(work_sheet.column_dimensions["G"].width)
    print(work_sheet.column_dimensions["H"].width)
    print(work_sheet.column_dimensions["I"].width)
    print(work_sheet.column_dimensions["J"].width)


def set_single_row_height(work_sheet: Worksheet, row: int, height: float):
    work_sheet.row_dimensions[row].height = height


def set_single_cell_alignment(work_sheet: Worksheet, position: str, alignment: Alignment):
    work_sheet[position].alignment = alignment


def set_single_cell_value(work_sheet: Worksheet, position: str, value: str, text_font: Font = Font()):
    cell = work_sheet[position]
    cell.value = value
    cell.font = text_font

