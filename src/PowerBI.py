from typing import List
from subprocess import check_call, DEVNULL, CalledProcessError
import io
import zipfile



class PowerBI:
	def __init__(self, path: str) -> None:
		self.path = path

	
	def decompress(self) -> None:
		with open(self.path, 'rb') as c_file:
			d_file = zipfile.ZipFile(io.BytesIO(c_file.read()))

			self.data = {
				'[Content_Types].xml': 	d_file.read('[Content_Types].xml').decode(),
				'DataModel': 			d_file.read('DataModel').decode(), # needs to be further decompressed
				'DiagramLayout': 		d_file.read('DiagramLayout').decode(),
				'Metadata': 			d_file.read('Metadata').decode(),
				'SecurityBindings': 	d_file.read('SecurityBindings').decode(), # needs to be further decompressed
				'Settings': 			d_file.read('Settings').decode(),
				'Version': 				d_file.read('Version').decode(),
			}