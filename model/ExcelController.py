import asyncio

from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from mapper.DataMapper import DataMapper
from openpyxl.styles import Font, Alignment, Border, Side
from model.LoadTextModel import LoadTextModel
from model.LoadExcelModel import LoadExcelModel
from model.Result import Success, Error
from model.dataclass.Fee import Fee
from model.dataclass.Student import Student
from excelutils.CustomWorksheet import format_all_columns_with_hard_coded, set_single_cell_alignment, set_single_row_height, \
    set_single_cell_value


class ExcelController:
    _text_data: Fee = None
    _student_data: Student = None

    def read_text_data_from_model(self, model: LoadTextModel):
        result = asyncio.run(model.read_all_text_from_note())
        if isinstance(result, Error):
            return

        if isinstance(result, Success):
            self._text_data = result.data

    def read_excel_data_from_model(self, model: LoadExcelModel):
        result = asyncio.run(model.read_all_data_from_excel())
        if isinstance(result, Error):
            return

        if isinstance(result, Success):
            self._student_data = result.data

    def initialize_excel_file(self, new_work_book: Workbook):
        new_work_sheet = new_work_book.active
        if self._text_data is not None:
            DataMapper.init_title_item(
                work_sheet=new_work_sheet,
                title="嘉義市私立菁英幼兒園(準公共幼兒園)",
                font_style=Font(size=12),
                alignment=Alignment(horizontal="center"),
                border=Border(
                    top=Side(style='thin'),
                    bottom=Side(style='thin'),
                    left=Side(style='thin'),
                    right=Side(style='thin')
                ),
                is_merged_title=True,
                start_position="A1",
                end_position="I1"
            )

            DataMapper.init_title_item(
                work_sheet=new_work_sheet,
                title="111學年度第    學期繳費收據",
                font_style=Font(size=12),
                alignment=Alignment(horizontal="center"),
                border=Border(
                    top=Side(style='thin'),
                    bottom=Side(style='thin'),
                    left=Side(style='thin'),
                    right=Side(style='thin')
                ),
                border_row=2,
                is_merged_title=True,
                start_position="A2",
                end_position="I2"
            )

            DataMapper.init_title_item(
                work_sheet=new_work_sheet,
                title="幼生姓名：      　　  班別：小班               111學年度    第一學期：111年8月1日至112年1月31日",
                font_style=Font(size=10),
                alignment=Alignment(horizontal="center"),
                border=Border(
                    top=Side(style='thin'),
                    bottom=Side(style='thin'),
                    left=Side(style='thin'),
                    right=Side(style='thin')
                ),
                border_row=3,
                is_merged_title=True,
                start_position="A3",
                end_position="I3"
            )

            DataMapper.init_title_item(
                work_sheet=new_work_sheet,
                title="年      月   費用      繳費日期：     年    月     日        年齡：3歲",
                font_style=Font(size=10),
                alignment=Alignment(horizontal="center"),
                border=Border(
                    top=Side(style="thin"),
                    bottom=Side(style="thin"),
                    left=Side(style="thin"),
                    right=Side(style="thin")
                ),
                border_row=4,
                is_merged_title=True,
                start_position="A4",
                end_position="I4"
            )

            DataMapper.init_multiple_merged_cell_item(
                work_sheet=new_work_sheet,
                position_value=["園所收費標準", "幼兒屬性", "家長每月繳費", "備註"],
                split_position=["A5:C5", "D5", "E5", "F5:I5"],
                font_style=Font(size=10),
                alignment=Alignment(horizontal="center"),
                border=Border(
                    top=Side(style="thin"),
                    bottom=Side(style="thin"),
                    left=Side(style="thin"),
                    right=Side(style="thin")
                )
            )

            DataMapper.init_three_column_fee_item(
                work_sheet=new_work_sheet,
                left_column_range="A6:A8",
                left_title="學期\n收費",
                middle_columns_value=["學費", "雜費"],
                right_columns_value=[self._text_data.tuition_fee, self._text_data.miscellaneous_fee_term],
                left_column_font=Font(size=9),
                other_font=Font(size=10),
                left_column_alignment=Alignment(horizontal="center", vertical="center"),
                middle_alignment=Alignment(horizontal="center", vertical="center"),
                right_alignment=Alignment(horizontal="right", vertical="center"),
                border=Border(
                    top=Side(style="thin"),
                    bottom=Side(style="thin"),
                    left=Side(style="thin"),
                    right=Side(style="thin")
                )
            )

            DataMapper.init_three_column_fee_item(
                work_sheet=new_work_sheet,
                left_column_range="A9:A14",
                left_title="月收費",
                middle_columns_value=["午餐費", "點心費", "材料費", "活動費", "雜費"],
                right_columns_value=[
                    self._text_data.lunch_fee,
                    self._text_data.dessert_fee,
                    self._text_data.material_fee,
                    self._text_data.activity_fee,
                    self._text_data.miscellaneous_fee_month
                ],
                left_column_font=Font(size=9),
                other_font=Font(size=10),
                left_column_alignment=Alignment(horizontal="center", vertical="center"),
                middle_alignment=Alignment(horizontal="center", vertical="center"),
                right_alignment=Alignment(horizontal="right", vertical="center"),
                border=Border(
                    top=Side(style="thin"),
                    bottom=Side(style="thin"),
                    left=Side(style="thin"),
                    right=Side(style="thin")
                )
            )

            DataMapper.init_multiple_merged_cell_item(
                work_sheet=new_work_sheet,
                position_value=["全學期收費", "37,164"],
                split_position=["A15:B15", "C15"],
                font_style=Font(size=10),
                alignment=Alignment(horizontal="center", vertical="center"),
                specify_position_alignment=Alignment(horizontal="right", vertical="center"),
                specify_position="C15",
                border=Border(
                    top=Side(style="thin"),
                    bottom=Side(style="thin"),
                    left=Side(style="thin"),
                    right=Side(style="thin")
                )
            )

            DataMapper.init_row_and_column_merged_item(
                work_sheet=new_work_sheet,
                merged_range="A16:B17",
                value="月平均收費",
                text_font=Font(size=10),
                text_alignment=Alignment(horizontal="center", vertical="center"),
                border=Border(
                    top=Side(style="thin"),
                    bottom=Side(style="thin"),
                    left=Side(style="thin"),
                    right=Side(style="thin")
                )
            )

            DataMapper.init_row_and_column_merged_item(
                work_sheet=new_work_sheet,
                merged_range="C16:C17",
                value="6,194",
                text_font=Font(size=10),
                text_alignment=Alignment(horizontal="right", vertical="center"),
                border=Border(
                    top=Side(style="thin"),
                    bottom=Side(style="thin"),
                    left=Side(style="thin"),
                    right=Side(style="thin")
                )
            )

            DataMapper.init_three_column_fee_item(
                work_sheet=new_work_sheet,
                left_column_range="A18:A20",
                left_title="其他代收",
                middle_columns_value=["交通費", "保險費", "課後拖延/月"],
                right_columns_value=[" ", " ", "750"],
                left_column_font=Font(size=9),
                other_font=Font(size=10),
                left_column_alignment=Alignment(horizontal="center", vertical="center"),
                middle_alignment=Alignment(horizontal="center", vertical="center"),
                right_alignment=Alignment(horizontal="right", vertical="center"),
                border=Border(
                    top=Side(style="thin"),
                    bottom=Side(style="thin"),
                    left=Side(style="thin"),
                    right=Side(style="thin")
                )
            )

            DataMapper.init_two_column_merged_row_item(
                work_sheet=new_work_sheet,
                left_split_row=["D6:D7", "D8:D9", "D10:D11", "D12:D13", "D14:D15", "D16:D17"],
                right_split_row=["E6:E7", "E8:E9", "E10:E11", "E12:E13", "E14:E15", "E16:E17"],
                left_column_values=["第1胎子女", "第2胎子女", "第3胎(含)以上子女", "低收入戶或\n中低收入", "", ""],
                right_column_values=[
                    self._text_data.first_child,
                    self._text_data.second_child,
                    self._text_data.third_child,
                    self._text_data.low_income_households,
                    "",
                    "-"
                ],
                left_alignment=Alignment(horizontal="center", vertical="center"),
                right_alignment=Alignment(horizontal="right", vertical="center"),
                text_font=Font(size=10),
                border=Border(
                    top=Side(style="thin"),
                    bottom=Side(style="thin"),
                    left=Side(style="thin"),
                    right=Side(style="thin")
                )
            )

            DataMapper.init_row_and_column_merged_item(
                work_sheet=new_work_sheet,
                merged_range="F6:I17",
                value="\n1.全學期以6個月計。\n\n2.自111年8月起，第1胎子女家長每月繳費不超過3,000元，第2胎不超過2,000元，第3胎以上不超過1,000元，低收入戶及中低收戶家長「免繳費用」，與幼兒園原收費間之差額，由行政院協助家長支付園方。",
                text_alignment=Alignment(horizontal="left", vertical="top"),
                text_font=Font(size=8),
                border=Border(
                    top=Side(style="thin"),
                    bottom=Side(style="thin"),
                    left=Side(style="thin"),
                    right=Side(style="thin")
                )
            )

            DataMapper.init_multiple_merged_cell_item(
                work_sheet=new_work_sheet,
                position_value=["幼兒屬性為：", "第1胎子女", "每月應繳", "1,000"],
                split_position=["D18", "E18:F18", "G18:H18", "I18"],
                font_style=Font(size=10),
                alignment=Alignment(vertical="center", horizontal="right"),
            )
            set_single_cell_alignment(work_sheet=new_work_sheet, position="E18",
                                      alignment=Alignment(vertical="center", horizontal="left"))

            DataMapper.init_row_and_column_merged_item(
                work_sheet=new_work_sheet,
                merged_range="D19:E21",
                value="收費金額合計新臺幣",
                text_alignment=Alignment(horizontal="right", vertical="center"),
                text_font=Font(size=10)
            )

            DataMapper.init_row_and_column_merged_item(
                work_sheet=new_work_sheet,
                merged_range="F19:G21",
                value=""
            )

            DataMapper.init_row_and_column_merged_item(
                work_sheet=new_work_sheet,
                merged_range="H19:H21",
                value="元整",
                text_alignment=Alignment(horizontal="center", vertical="center"),
                text_font=Font(size=10)
            )

            DataMapper.init_row_and_column_merged_item(
                work_sheet=new_work_sheet,
                merged_range="A21:A22",
                value="請假\n減收",
                text_alignment=Alignment(horizontal="center", vertical="center"),
                text_font=Font(size=9),
                border=Border(
                    top=Side(style="thin"),
                    bottom=Side(style="thin"),
                    left=Side(style="thin"),
                    right=Side(style="thin")
                )
            )

            DataMapper.init_row_and_column_merged_item(
                work_sheet=new_work_sheet,
                merged_range="B21:B22",
                value="午餐點心",
                text_alignment=Alignment(horizontal="center", vertical="center"),
                text_font=Font(size=9),
                border=Border(
                    top=Side(style="thin"),
                    bottom=Side(style="thin"),
                    left=Side(style="thin"),
                    right=Side(style="thin")
                )
            )

            DataMapper.init_row_and_column_merged_item(
                work_sheet=new_work_sheet,
                merged_range="C21:C22",
                value="",
                border=Border(
                    top=Side(style="thin"),
                    bottom=Side(style="thin"),
                    left=Side(style="thin"),
                    right=Side(style="thin")
                )
            )

            DataMapper.init_multiple_merged_cell_item(
                work_sheet=new_work_sheet,
                position_value=["園長：　　　　　　　　　　會計：　　　　　　　　　　　　經辦："],
                split_position=["A23:I23"],
                font_style=Font(size=10),
                alignment=Alignment(horizontal="center", vertical="center"),
            )
            set_single_row_height(work_sheet=new_work_sheet, row=23, height=26)

            set_single_cell_value(
                work_sheet=new_work_sheet,
                position="A24",
                value="——————————————————————————————————————",
                text_font=Font(size=12)
            )

            DataMapper.init_multiple_row_item(
                work_sheet=new_work_sheet,
                vertical_string="第二聯存根聯",
                start_position="J5"
            )

            format_all_columns_with_hard_coded(work_sheet=new_work_sheet)

            new_work_book.save("receipt.xlsx")
            return

        print("Text data or student data has not been loaded, call read_text_data_from_model & read_excel_data_from_model first.")

    # TODO
    # def set_text_data_to_excel_file(self, data: Fee):

