from typing import Union, Optional
import win32com.client as win32



class Excel:
	def __init__(self, path: str) -> None:
		self.excel = win32.gencache.EnsureDispatch('Excel.Application')
		self.workbook = self.excel.Workbooks.Open(path)

	def save(self) -> None:
		self.workbook.Save()


	def close(self) -> None:
		self.save()
		self.workbook.Close()


	def exit(self) -> None:
		self.excel.Quit()

	
	def sheets(self) -> object:
		return self.workbook.Worksheets

	
	def sheet(self, index: Optional[int]=None, name: Optional[str]=None) -> Union[object, None]:
		if name:
			for sheet in self.sheets():
				if sheet.Name == name:
					return sheet
				
			return None
		
		else:
			return self.workbook.Worksheets(index)
