from SASConnection import SASConnection
from ExcelConnection import ExcelConnection
from dotenv import load_dotenv
import os


load_dotenv()


def setup_sas():
    _server = os.getenv('SERVER')
    _server_port = os.getenv('SERVER_PORT')
    _iom_protocol = os.getenv('IOM_PROTOCOL')
    _creds = os.getenv('CREDS')
    _server_id = os.getenv('SERVER_ID')

    return SASConnection(_server, _server_port, _iom_protocol, _creds, _server_id)


def setup_excel(path):
    return ExcelConnection(path)


if __name__ == '__main__':
    #sas = setup_sas()
    excel = setup_excel(f'{os.getcwd()}\\test.xlsx')
    sheet = excel.open_sheet()

    for i in range(1, 11):
        for j in range(1, 11):
            print(int(sheet.Cells(i, j).Value))
