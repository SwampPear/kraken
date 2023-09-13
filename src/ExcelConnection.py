import win32com.client as win32
import os


class ExcelConnection:
	def __init__(self) -> None:
		excel = win32.gencache.EnsureDispatch('Excel.Application')
		wb = excel.Workbooks.Open(f'{os.getcwd()}\\test.xlsx')
		print(excel.Workbooks.Count)
		print(os.getcwd())
		
