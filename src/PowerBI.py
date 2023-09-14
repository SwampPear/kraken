from typing import List
from subprocess import check_call, DEVNULL, STDOUT, CalledProcessError
import os



class PowerBI:
	def __init__(self, path: str) -> None:
		self.path = path

		self.root = os.getcwd()
		self.temp_dir = f'{self.root}\\temp'

		self.init_temp_dir()


	def call_cmd(self, cmd: List[str]) -> None:
		try:
			check_call(cmd, shell=True, stdout=DEVNULL, stderr=DEVNULL)

		except CalledProcessError:
			pass # should change later

	
	def move_temp(self, file: str) -> None:
		self.call_cmd(['move', f'{self.root}\\{file}', f'{self.temp_dir}'])

	
	def init_temp_dir(self):
		# clear contents of dir
		self.call_cmd(['rd', f'{self.temp_dir}', '/s', '/q'])
		self.call_cmd(['mkdir', f'{self.temp_dir}'])

		# copy and unzip pbix
		self.call_cmd(['copy', f'{self.root}\\{self.path}.pbix', f'{self.temp_dir}\\'])
		self.call_cmd(['unzip', f'{self.temp_dir}\\{self.path}.pbix'])

		# move contents to temp dir
		self.move_temp('[Content_Types].xml')
		self.move_temp('DataModel')
		self.move_temp('DiagramLayout')
		self.move_temp('MetaData')
		self.move_temp('SecurityBindings')
		self.move_temp('Settings')
		self.move_temp('Version')
		self.move_temp('Report')
	
