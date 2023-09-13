import win32com.client

class SASConnection:
	def __init__(self) -> None:
		self.obj_factory = win32com.client.Dispatch("SASObjectManager.ObjectFactoryMulti2")
