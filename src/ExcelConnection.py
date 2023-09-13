import win32com.client as win32
import os


class ExcelConnection:
	def __init__(self) -> None:
		self.excel = win32.gencache.EnsureDispatch('Excel.Application')

	
	def open_workbook(self, path):
		return self.excel.Workbooks.Open(f'{os.getcwd()}\\test.xlsx')
