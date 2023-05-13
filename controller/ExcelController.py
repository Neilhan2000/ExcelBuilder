import asyncio

from openpyxl.workbook import Workbook

from excelutils.opserverpattern.Subject import Subject
from mapper.DataMapper import DataMapper
from openpyxl.styles import Font, Alignment, Border, Side
from model.LoadTextModel import LoadTextModel
from model.LoadExcelModel import LoadExcelModel
from model.Result import Success, Error
from model.dataclass.Fee import Fee
from model.dataclass.Date import Date
from excelutils.CustomWorksheet import format_all_columns_with_hard_coded, set_single_cell_alignment, \
    set_single_row_height, \
    set_single_cell_value
from datetime import datetime

from model.dataclass.Student import Student
from view import ReceiptView


class ExcelController(Subject):
    # _student_data: Student = None
    text_data: Fee = None
    student_list: Student  # now it is a single student, It will be a list in the future
    receipt_issue_date = None

    def __init__(self):
        self.__get_system_date()

    def read_text_data_from_model(self, model: LoadTextModel):
        result = asyncio.run(model.read_all_text_from_note())
        if isinstance(result, Error):
            return

        if isinstance(result, Success):
            self.text_data = result.data

    def read_excel_data_from_model(self, model: LoadExcelModel):
        result = asyncio.run(model.read_all_data_from_excel())
        if isinstance(result, Error):
            return

        if isinstance(result, Success):
            self.student_list = result.data

    # TODO
    # def set_text_data_to_excel_file(self, data: Fee):

    def __get_system_date(self):
        system_date = datetime.now()
        self.receipt_issue_date = Date(
            year=system_date.year - 1911,
            month=system_date.month,
            day=system_date.day
        )
        print(f"Issued date is {self.receipt_issue_date}")

    def is_model_initialized(self) -> bool:
        return self.text_data and self.student_list and self.receipt_issue_date is not None

    # below function should be moved to model class or somewhere
    def map_class_type_to_age(self, class_type) -> int:
        match class_type:
            case "大班":
                return 5
            case "中班":
                return 4
            case "小班":
                return 3

    def map_date_to_term_period(self) -> str:
        current_year = self.receipt_issue_date.year
        current_month = self.receipt_issue_date.month
        if 2 <= current_month <= 7:
            return f"{current_year - 1}學年度    第二學期：{current_year}年2月1日至{current_year}年7月31日"
        return f"{current_year}學年度    第一學期：{current_year}年8月1日至{current_year + 1}年1月31日"

    def add_view(self, view: ReceiptView):
        self.attach(view)

