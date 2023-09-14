import win32com.client as win32
import os


class Excel:
	def __init__(self, path) -> None:
		self.excel = win32.gencache.EnsureDispatch('Excel.Application')
		self.workbook = self.excel.Workbooks.Open(path)

	
	def save(self) -> None:
		self.workbook.Save()

	
	def close(self) -> None:
		self.save()
		self.workbook.Close()


	def exit(self) -> None:
		self.excel.Quit()

	
	def sheet(self, index):
		return self.workbook.Worksheets(index)
	

	def sheet(self, name):
		for _sheet in self.workbook.Worksheets():
			if _sheet.Name == name:
				return _sheet
			
		return  None
