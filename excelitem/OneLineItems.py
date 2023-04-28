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
        work_sheet[convert_col_and_row_to_position(col_number=i, row_number=border_row)].border = border


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
        work_sheet[convert_col_and_row_to_position(col_number=i, row_number=border_row)].border = border


def init_three_column_fee_item(
        work_sheet: Worksheet,
        left_column_range: str,
        left_title: str,
        middle_columns_value: list,
        right_columns_value: list,
        left_column_font: Font = Font(),
        left_column_alignment=Alignment(),
        other_font: Font = Font(),
        middle_alignment=Alignment(),
        right_alignment=Alignment(),
        border=Border()
):
    start_position = find_start_cell(left_column_range)
    start_column = start_position[0]
    start_row = start_position[-1]
    total_row_number = int(find_cell_digit(find_end_cell(left_column_range))) - int(find_cell_digit(start_position)) + 1
    middle_row_count = 0
    right_row_count = 0
    print(total_row_number)

    def get_left_position(increase_row: int) -> str:
        return convert_col_and_row_to_position(
            col_number=convert_column_str_to_int(column_str=start_column),
            row_number=int(start_row) + increase_row
        )

    def get_middle_position(increase_row: int) -> str:
        return convert_col_and_row_to_position(
            col_number=convert_column_str_to_int(column_str=start_column) + 1,
            row_number=int(start_row) + increase_row
        )

    def get_right_position(increase_row: int) -> str:
        return convert_col_and_row_to_position(
            col_number=convert_column_str_to_int(column_str=start_column) + 2,
            row_number=int(start_row) + increase_row
        )

    # initialize left column
    work_sheet.merge_cells(left_column_range)
    work_sheet[start_position].value = left_title
    work_sheet[start_position].font = left_column_font
    work_sheet[start_position].alignment = left_column_alignment

    for i in range(total_row_number):
        left_position_cell = get_left_position(increase_row=i)
        work_sheet[left_position_cell].border = border

    # initialize middle column
    for middle_element in middle_columns_value:
        middle_position_cell = work_sheet[get_middle_position(increase_row=middle_row_count)]

        middle_position_cell.value = middle_element
        middle_position_cell.font = other_font
        middle_position_cell.alignment = middle_alignment
        middle_row_count += 1

    for i in range(total_row_number):
        middle_position_cell = get_middle_position(increase_row=i)
        work_sheet[middle_position_cell].border = border

    # initialize right column
    for middle_element in right_columns_value:
        right_position_cell = work_sheet[get_right_position(increase_row=right_row_count)]

        right_position_cell.value = middle_element
        right_position_cell.font = other_font
        right_position_cell.alignment = right_alignment
        right_row_count += 1

    for i in range(total_row_number):
        right_position_cell = get_right_position(increase_row=i)
        work_sheet[right_position_cell].border = border


def convert_col_and_row_to_position(col_number: int, row_number: int) -> str:  # make it private
    match col_number:
        case 0:
            return f"A{row_number}"
        case 1:
            return f"B{row_number}"
        case 2:
            return f"C{row_number}"
        case 3:
            return f"D{row_number}"
        case 4:
            return f"E{row_number}"
        case 5:
            return f"F{row_number}"
        case 6:
            return f"G{row_number}"
        case 7:
            return f"H{row_number}"
        case 8:
            return f"I{row_number}"
        case 9:
            return f"J{row_number}"


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


def convert_column_str_to_int(column_str) -> int:
    match column_str:
        case "A":
            return 0
        case "B":
            return 1
        case "C":
            return 2
        case "D":
            return 3
        case "E":
            return 4
        case "F":
            return 5
        case "G":
            return 6
        case "H":
            return 7
        case "I":
            return 8
        case "J":
            return 9


def find_start_cell(cells_str: str):
    return cells_str.split(":")[0]


def find_end_cell(cells_str: str):
    return cells_str.split(":")[-1]


def find_cell_digit(cell: str) -> int:
    return int("".join(filter(str.isdigit, cell)))

