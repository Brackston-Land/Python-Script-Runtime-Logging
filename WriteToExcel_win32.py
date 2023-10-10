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
try:

    import time
    # Open an existing workbook
    import win32com.client as win32
    from datetime import datetime
    from pytz import timezone

except ImportError:
    print("Modules cannot be found")


class WriteToExcel:

    def __init__(self, ScriptName):
        self.ScriptName = ScriptName



    def colnum_string(self, n):
        string = ""
        while n > 0:
            n, remainder = divmod(n - 1, 26)
            string = chr(65 + remainder) + string
        return string

    def startTime(self):

        # define date format
        fmt = '%Y-%m-%d %H:%M:%S'
        # define eastern timezone
        central = timezone('US/Central')
        # localized datetime
        loc_dt = datetime.now(central)


        excel = win32.gencache.EnsureDispatch('Excel.Application')
        wb = excel.Workbooks.Open(r"Path to your Script Run Times and Database Lock Durations.xlsx")
        ws = wb.Worksheets('Table')

        for col_num in range(ws.UsedRange.Columns.Count):
            row_num = 1

            # Note Python's range() counts from 0 and Excel counts from 1
            column_value = ws.Cells(row_num, col_num + 1).Value

            if column_value == "Name":
                colnum = col_num

            if column_value == "Start_Time":
                column_string = self.colnum_string(col_num +1)


        for row_num in range(ws.UsedRange.Rows.Count):
            # Note Python's range() counts from 0 and Excel counts from 1
            value = ws.Cells(row_num+1, colnum + 1).Value


            if value == self.ScriptName:
                cells = ("{}{}").format(column_string, row_num+1)
                ws.Range(cells).Value = loc_dt.strftime(fmt)
        wb.Save()
        excel.Application.Quit()


    def endTime(self):

        # define date format
        fmt = '%Y-%m-%d %H:%M:%S'
        # define eastern timezone
        central = timezone('US/Central')
        # localized datetime
        loc_dt = datetime.now(central)


        excel = win32.gencache.EnsureDispatch('Excel.Application')
        wb = excel.Workbooks.Open(r"Path to your Script Run Times and Database Lock Durations.xlsx")
        ws = wb.Worksheets('Table')

        for col_num in range(ws.UsedRange.Columns.Count):
            row_num = 1

            # Note Python's range() counts from 0 and Excel counts from 1
            column_value = ws.Cells(row_num, col_num + 1).Value

            if column_value == "Name":
                colnum = col_num

            if column_value == "End_Time":
                column_string = self.colnum_string(col_num +1)


        for row_num in range(ws.UsedRange.Rows.Count):
            # Note Python's range() counts from 0 and Excel counts from 1
            value = ws.Cells(row_num+1, colnum + 1).Value


            if value == self.ScriptName:
                cells = ("{}{}").format(column_string, row_num+1)
                ws.Range(cells).Value = loc_dt.strftime(fmt)
        wb.Save()
        excel.Application.Quit()


    def grabErrors(self, status):


        excel = win32.gencache.EnsureDispatch('Excel.Application')
        wb = excel.Workbooks.Open(r"Path to your Script Run Times and Database Lock Durations.xlsx")
        ws = wb.Worksheets('Table')


        for col_num in range(ws.UsedRange.Columns.Count):
            row_num = 1

            # Note Python's range() counts from 0 and Excel counts from 1
            column_value = ws.Cells(row_num, col_num + 1).Value

            if column_value == "Name":
                colnum = col_num

            if column_value == "Error":
                column_string = self.colnum_string(col_num +1)

        for row_num in range(ws.UsedRange.Rows.Count):
            # Note Python's range() counts from 0 and Excel counts from 1
            value = ws.Cells(row_num+1, colnum + 1).Value

            if value == self.ScriptName:
                cells = ("{}{}").format(column_string, row_num+1)
                if status == False:
                    ws.Range(cells).Value ="No"
                if status == True:
                    ws.Range(cells).Value ="Yes"

        wb.Save()
        excel.Application.Quit()

