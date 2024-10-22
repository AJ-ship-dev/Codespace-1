
Dim excelApp, workbook, sheet, usedRange, table, xlsxPath


xlsxPath = WScript.Arguments(0)


Set excelApp = CreateObject("Excel.Application")
excelApp.Visible = False ' Set to True if you want to see Excel during the process


Set workbook = excelApp.Workbooks.Open(xlsxPath)


Set sheet = workbook.Sheets(1)

Set usedRange = sheet.UsedRange

If Not usedRange Is Nothing Then
   
    Set table = sheet.ListObjects.Add(1, usedRange, , 1)
    table.Name = "MyTable"
End If

excelApp.DisplayAlerts = False  
workbook.SaveAs xlsxPath, 51  


workbook.Close False
excelApp.Quit

Set usedRange = Nothing
Set sheet = Nothing
Set workbook = Nothing
Set excelApp = Nothing
