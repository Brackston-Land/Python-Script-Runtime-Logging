#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      Brackston Land
#
# Created:     18/05/2020
# Copyright:   (c) Brackston Land 2020
# Licence:     <your licence>
#-------------------------------------------------------------------------------
import WriteToExcel_win32
import time
import sys
import os

fullScriptPath = sys.argv[0]

#this finds the name of the script being run: ExampleScript.py
scriptname = os.path.basename(fullScriptPath)
print(scriptname)


#Write start time to excel sheet
startTime = WriteToExcel_win32.WriteToExcel(scriptname) #WriteToExcel_win32.WriteToExcel(ScriptName)
startTime.startTime()

time.sleep(5)

try:
    #kk

    #Write errors to excel sheet
    writeError = WriteToExcel_win32.WriteToExcel(scriptname)   #WriteToExcel_win32.WriteToExcel(column_letter, ScriptName)
    #Since the try statement worked, this will write "False" in the error column
    writeError.grabErrors(False)

except:

    #Write errors to excel sheet
    writeError = WriteToExcel_win32.WriteToExcel(scriptname) #WriteToExcel_win32.WriteToExcel(column_letter, ScriptName)
    #Since the try statement DID NOT work, this will write "True" in the error column
    writeError.grabErrors(True)


#Write end time to excel sheet
endTime = WriteToExcel_win32.WriteToExcel(scriptname) #WriteToExcel_win32.WriteToExcel(column_letter, ScriptName)
endTime.endTime()

print("All Done")

