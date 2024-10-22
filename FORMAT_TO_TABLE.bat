@echo off
:: Set the path to your .xlsx file
set "xlsxFile=U:\AJ Dev\TEST\Funding Transactions_1729262777420.xlsx"

:: Set the correct path to the VBScript file
set "vbScriptPath=U:\AJ Dev\vbscript_test.vbs"

:: Call the VBScript to process the .xlsx file and convert data to table
cscript //nologo "%vbScriptPath%" "%xlsxFile%"

echo Done! The .xlsx file has been converted and saved.
pause
