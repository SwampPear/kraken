import win32com.client as win32



class PowerBI:
	def __init__(self) -> None:
		self.app = win32.gencache.EnsureDispatch('PowerBIDesktop.Application')