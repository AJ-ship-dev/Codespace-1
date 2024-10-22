' VBScript to open an .xlsx file, convert the data to a table, and save it

Dim excelApp, workbook, sheet, usedRange, table, xlsxPath

' Get the .xlsx file path from the command-line argument
xlsxPath = WScript.Arguments(0)

' Create an Excel application object
Set excelApp = CreateObject("Excel.Application")
excelApp.Visible = False ' Set to True if you want to see Excel during the process

' Open the .xlsx file in Excel
Set workbook = excelApp.Workbooks.Open(xlsxPath)

' Get the first sheet (you can modify this to target a different sheet if needed)
Set sheet = workbook.Sheets(1)

' Select the used range (all data in the sheet)
Set usedRange = sheet.UsedRange

' Check if the used range is not empty
If Not usedRange Is Nothing Then
    ' Convert the used range into a Table (like Ctrl+L does in Excel)
    Set table = sheet.ListObjects.Add(1, usedRange, , 1)
    table.Name = "MyTable"
End If

' Save the workbook back as .xlsx (it will overwrite the original file)
excelApp.DisplayAlerts = False  ' Prevent prompts about overwriting
workbook.SaveAs xlsxPath, 51  ' 51 is the format code for .xlsx files

' Close the workbook and Excel application
workbook.Close False
excelApp.Quit

' Clean up
Set usedRange = Nothing
Set sheet = Nothing
Set workbook = Nothing
Set excelApp = Nothing
