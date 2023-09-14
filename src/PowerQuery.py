import win32com.client as win32


class PowerQuery:
    def __init__(self) -> None:
        self.app = win32.gencache.EnsureDispatch('PowerQuery.Application')