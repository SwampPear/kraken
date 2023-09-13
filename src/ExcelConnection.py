from APIConnection import APIConnection


class ExcelConnection(APIConnection):
	def __init__(self, endpoint: str) -> None:
		super().__init__(endpoint)
		