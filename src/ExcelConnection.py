from BaseAPIConnection import BaseAPIConnection


class ExcelConnection(BaseAPIConnection):
	def __init__(self, endpoint: str) -> None:
		super().__init__(endpoint)
