import win32com.client


class ExcelConnection:
	def __init__(self) -> None:
		excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
		print(excel.Workbooks.Count)

