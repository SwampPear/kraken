import win32com.client


class SASConnection:
	def __init__(self) -> None:
		self.factory = win32com.client.Dispatch("SASObjectManager.ObjectFactoryMulti2")
		self.server_def = self.init_server_def()
		self.sas = self.init_sas

	def init_server_def(self):
		server_def = win32com.client.Dispatch("SASObjectManager.ServerDef")
		#server_def.MachineDNSName = "servername"
		#server_def.Port = 8591    # workspace server port
		#server_def.Protocol = 2   # 2 = IOM protocol
		#server_def.BridgeSecurityPackage = "Username/Password"
		#server_def.ClassIdentifier = "workspace server id"

		return server_def

	def init_sas(self):
		return self.factory.CreateObjectByServer("SASApp", True, self.server_def, "uid", "pw")

