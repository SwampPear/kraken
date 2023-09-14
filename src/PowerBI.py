import win32com.client as win32



class PowerBI:
	def __init__(self, path: str) -> None:
		self.app = win32.gencache.EnsureDispatch('PowerBI.Application')