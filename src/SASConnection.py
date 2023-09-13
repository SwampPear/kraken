from APIConnection import APIConnection


class SASConnection(APIConnection):
	def __init__(self, endpoint: str) -> None:
		super().__init__(endpoint)
