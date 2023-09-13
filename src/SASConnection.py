import win32com.client


class SASConnection:
	def __init__(self, server, server_port, iom_protocol, creds, server_id) -> None:
		self.factory = win32com.client.Dispatch('SASObjectManager.ObjectFactoryMulti2')
		self.server_def = self.init_server_def(server, server_port, iom_protocol, creds, server_id)
		self.sas = self.init_sas()

	def init_server_def(self, server, server_port, iom_protocol, creds, server_id):
		server_def = win32com.client.Dispatch('SASObjectManager.ServerDef')
		
		server_def.MachineDNSName = server			# server name
		server_def.Port = server_port				# workspace server port
		server_def.Protocol = iom_protocol			# 2 = IOM protocol
		server_def.BridgeSecurityPackage = creds	# username/password
		server_def.ClassIdentifier = server_id		# workspace server id

		return server_def

	def init_sas(self):
		return self.factory.CreateObjectByServer("SASApp", True, self.server_def, "uid", "pw")

