#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      Brackston Land
#
# Created:     12/05/2020
# Copyright:   (c) Brackston Land 2020
# Licence:     <your licence>
#-------------------------------------------------------------------------------
import datetime
import time
import traceback

import win32com.client as win32

try:
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    # Open an existing workbook
    wb = excel.Workbooks.Open(r'Path to your Script Run Times and Database Lock Durations.xlsx')
    ws = wb.Worksheets('Table')
    ws.Range("K8").Value = datetime.datetime.now()
    time.sleep(20)
    #This portion writes the start time to excel sheet
    #writes time to specified cells
    ws.Range("L8").Value = datetime.datetime.now()
    wb.Save()
    excel.Application.Quit()
except Exception as ex:
    print(traceback.format_exc())