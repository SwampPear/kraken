import win32com.client as win32
import os


class ExcelConnection:
	def __init__(self, path) -> None:
		self.excel = win32.gencache.EnsureDispatch('Excel.Application')
		self.workbook = self.excel.Workbooks.Open(path)

	
	def open_sheet(self):
		return self.workbook.Worksheets(1)
