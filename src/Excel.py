from typing import Optional, List
import win32com.client as win32



class Excel:
	def __init__(self, path: str) -> None:
		self.app = win32.gencache.EnsureDispatch('Excel.Application')
		self.workbook = self.app.Workbooks.Open(path)


	def save(self) -> None:
		self.workbook.Save()


	def close(self) -> None:
		self.save()
		self.workbook.Close()


	def exit(self) -> None:
		self.app.Quit()

	
	def sheets(self) -> List[object]:
		return self.workbook.Worksheets

	
	def sheet(self, index: Optional[int]=None, name: Optional[str]=None) -> object:
		if name:
			for sheet in self.sheets():
				if sheet.Name == name:
					return sheet
				
			return None
		
		else:
			return self.workbook.Worksheets(index)
