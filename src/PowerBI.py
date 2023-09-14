import win32com.client as win32



class PowerBI:
	def __init__(self, path: str) -> None:
		self.excel = win32.gencache.EnsureDispatch('PowerBI.Application')