from SASConnection import SASConnection
from Excel import Excel
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


if __name__ == '__main__':
    #sas = setup_sas()
    excel = Excel(f'{os.getcwd()}\\test2.xlsx')
    sheet = excel.sheet(1)

    for i in range(1, 11):
        for j in range(1, 11):
            print(sheet.Cells(i, j).Value)

    excel.close()
