import pandas as pd
from win32com.client import Dispatch, DispatchEx
import win32com.client
import logging
from log_config import log_location, log_filemode, log_format, log_datefmt

logger = logging.getLogger(__name__)


# https://stackoverflow.com/questions/3961007/passing-an-array-list-into-python
# https://stackoverflow.com/questions/36901/what-does-double-star-asterisk-and-star-asterisk-do-for-parameters
def run_access_macro(access_path, macros=[], *args):
    try:
        # https://stackoverflow.com/questions/9177984/how-to-run-a-ms-access-macro-from-python

        access_app = DispatchEx("Access.Application")
        access_app.Visible = False
        access_app.OpenCurrentDatabase(access_path)
        access_db = access_app.CurrentDb()

        for macro in macros:
            access_app.DoCmd.RunMacro(macro)

        access_app.Application.Quit()
        del access_app
        del access_db
    except:
        logger.error(f'Exception occurred', exc_info=True)


# https://stackoverflow.com/questions/3961007/passing-an-array-list-into-python
# https://stackoverflow.com/questions/36901/what-does-double-star-asterisk-and-star-asterisk-do-for-parameters
def run_excel_macro(excel_path, macros=[], *args):
    try:
        # https://redoakstrategic.com/pythonexcelmacro/
        # https://stackoverflow.com/questions/47608506/issue-in-using-win32com-to-access-excel-file

        xlapp = win32com.client.DispatchEx('Excel.Application')
        xlbook = xlapp.Workbooks.Open(excel_path)

        for macro in macros:
            xlapp.Application.Run(macro)

        xlbook.Save()
        xlbook.Close()
        xlapp.Quit()
        del xlbook
        del xlapp
    except:
        logger.error(f'Exception occurred', exc_info=True)
