# Python-Script-Runtime-Logging
A way to track script successes and errors in an easy to read excel workbook. Designed for automated batch script runs.



## Installation
git clone https://github.com/giswqs/python-geospatial.git


## Libraries Used

* :zap: time
* :zap: pywin32
* :zap: datetime
* :zap: pytz
* :zap: attrs


## Description of scripts

### SuperSimple_WriteToExcel_win32.py

This is just a super simple script
showing how to write to cells with win32 module. 


### WriteToExcel_win32.py
Final script for writting start and end times of scripts to excel sheet.
Also writes to excel sheet if there is an error in the script. This script can
be placed in the site-libs location of the python environment being used so that 
it can be simply imported and used as a module. Ie.  Import WriteToExcel_win32.py



### ExampleScript.py
A script demonstrating how to call  and use the
WriteToExcel_win32.py script 



## Other Considerations

If you have difficulty installing pywin32, please try pypiwin32: 
https://stackoverflow.com/questions/23864234/importerror-no-module-named-win32com-client
If you have difficulty installing attrs, please try attr: 
https://stackoverflow.com/questions/49228744/attributeerror-module-attr-has-no-attribute-s
