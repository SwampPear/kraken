import win32com.client as win32


class PowerPoint:
    def __init__(self, path: str) -> None:
        self.app = win32.gencache.EnsureDispatch('PowerPoint.Application')
        self.workbook = self.app.Workbooks.Open(path)