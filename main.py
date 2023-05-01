import asyncio

from openpyxl.reader.excel import load_workbook
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from model.ExcelController import ExcelController
from model.ExcelModel import ExcelModel

work_book: Workbook = load_workbook(filename="ReceiptExample.xlsx")
work_sheet: Worksheet = work_book.active
sheet_names: list = work_book.sheetnames

controller = ExcelController()
model = ExcelModel()

controller.read_text_data_from_model(model=model)
controller.initialize_excel_file()


# need open the finished file function
