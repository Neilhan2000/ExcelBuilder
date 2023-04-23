from openpyxl.styles import Font, Alignment, Border
from openpyxl.worksheet.worksheet import Worksheet


def init_title_item(
        work_sheet: Worksheet,
        title: str,
        font_style: Font = Font(),
        alignment: Alignment = Alignment(),
        border: Border = Border(),
        border_row: int = 1,
        is_merged_title: bool = False,
        start_position: str = "A1",
        end_position: str = "A1"
):
    work_sheet.append([title])

    # initialize title cell
    if is_merged_title:
        work_sheet.merge_cells(f"${start_position}:${end_position}")
    work_sheet[start_position].font = font_style
    work_sheet[start_position].alignment = alignment

    # set merged cell border
    for i in range(convert_position_to_col_number(end_position) + 1):
        work_sheet[convert_col_number_to_position(col_number=i, row_number=border_row)].border = border


def init_multiple_merged_cell_item(
        work_sheet: Worksheet,
        position_value: list,
        split_position: list,
        font_style: Font = Font(),
        alignment: Alignment = Alignment(),
        border: Border = Border(),
        border_row: int = 1
):
    # initialize all merged cells
    for cells_need_merged in split_position:
        work_sheet.merge_cells(cells_need_merged)
        work_sheet[find_start_cell(cells_need_merged)].value = position_value[split_position.index(cells_need_merged)]
        work_sheet[find_start_cell(cells_need_merged)].font = font_style
        work_sheet[find_start_cell(cells_need_merged)].alignment = alignment

    # find last element and use it to set up border
    for i in range(convert_position_to_col_number(find_end_cell(split_position[-1])) + 1):
        work_sheet[convert_col_number_to_position(col_number=i, row_number=border_row)].border = border

def convert_col_number_to_position(col_number: int, row_number: int) -> str:  # make it private
    match col_number:
        case 0:
            return f"A${row_number}"
        case 1:
            return f"B${row_number}"
        case 2:
            return f"C${row_number}"
        case 3:
            return f"D${row_number}"
        case 4:
            return f"E${row_number}"
        case 5:
            return f"F${row_number}"
        case 6:
            return f"G${row_number}"
        case 7:
            return f"H${row_number}"
        case 8:
            return f"I${row_number}"
        case 9:
            return f"J${row_number}"


def convert_position_to_col_number(position: str) -> int:
    if position.__contains__("A"):
        return 0
    if position.__contains__("B"):
        return 1
    if position.__contains__("C"):
        return 2
    if position.__contains__("D"):
        return 3
    if position.__contains__("E"):
        return 4
    if position.__contains__("F"):
        return 5
    if position.__contains__("G"):
        return 6
    if position.__contains__("H"):
        return 7
    if position.__contains__("I"):
        return 8
    if position.__contains__("J"):
        return 9


def find_start_cell(cells_str: str):
    return cells_str.split(":")[0]


def find_end_cell(cells_str: str):
    return cells_str.split(":")[-1]
