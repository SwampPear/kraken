from SASConnection import SASConnection
from Excel import Excel
from dotenv import load_dotenv
import os


load_dotenv()

"""
def setup_sas():
    _server = os.getenv('SERVER')
    _server_port = os.getenv('SERVER_PORT')
    _iom_protocol = os.getenv('IOM_PROTOCOL')
    _creds = os.getenv('CREDS')
    _server_id = os.getenv('SERVER_ID')

    return SASConnection(_server, _server_port, _iom_protocol, _creds, _server_id)
"""


if __name__ == '__main__':
    file = f'{os.getcwd()}\\test.xlsx'

    excel = Excel(file)

    print(type(excel.sheet(name='Sheet1')))

    #excel.close()
    #excel.exit()
