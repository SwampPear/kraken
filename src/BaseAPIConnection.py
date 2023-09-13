class BaseAPIConnection:
	def __init__(self, endpoint: str) -> None:
		self.endpoint = endpoint
		