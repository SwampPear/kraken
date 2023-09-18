from Excel import Excel
from PowerBI import PowerBI
import os

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
    """
    file = f'{os.getcwd()}\\test.xlsx'

    excel = Excel(file)

    sheet = excel.sheet(name='Sheet1')

    for i in range(1, 6):
        for j in range(1, 6):
            print(sheet.Cells(i, j))

    excel.save()
    """
    
    test = PowerBI('test')
    #stest.save()

