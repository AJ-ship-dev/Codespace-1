Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

try {
    Import-Module ImportExcel

    $xlsxFolder = "D:\OpCon-Files\MANTL\PROD\FUNDING"
    $date = Get-Date -Format "yyyyMMdd"  

    Write-Host "Searching for files in: $xlsxFolder"
    
    
    $csvFile = Get-ChildItem -Path $xlsxFolder -Filter "simplified_reports__funding.csv"

    
    $xlsxFiles = Get-ChildItem -Path $xlsxFolder -Filter "Funding Transactions_*.xlsx"

    
    if ($csvFile) {
        Write-Host "Found CSV file: $($csvFile.FullName)"

        $newCsvFileName = "Funding Transactions_$date.xlsx"
        $newCsvFilePath = Join-Path $xlsxFolder $newCsvFileName

        
        $csvData = Import-Csv -Path $csvFile.FullName
        $csvData | Export-Excel -Path $newCsvFilePath -AutoSize -TableName "SimplifiedTable"

        Write-Host "CSV data saved as Excel: $newCsvFilePath"
    } else {
        Write-Host "Error: No CSV file found matching pattern 'simplified_reports__funding.csv' in $xlsxFolder"
    }

    
    if ($xlsxFiles.Count -gt 0) {
        foreach ($xlsxFile in $xlsxFiles) {
            Write-Host "Found Excel file: $($xlsxFile.FullName)"

            
            $excelData = Import-Excel -Path $xlsxFile.FullName

            
            $excelData | Export-Excel -Path $xlsxFile.FullName -AutoSize -TableName "MyTable"

            Write-Host "Table created and saved in $xlsxFile.FullName"

            
            $newFileName = [System.IO.Path]::GetFileNameWithoutExtension($xlsxFile.Name) + "_$date" + $xlsxFile.Extension
            $newFilePath = Join-Path $xlsxFolder $newFileName

            Rename-Item -Path $xlsxFile.FullName -NewName $newFileName

            Write-Host "File renamed to: $newFileName"
        }
    } else {
        Write-Host "Error: No Excel files found matching pattern 'Funding Transactions_*.xlsx' in $xlsxFolder"
    }

} catch {
    Write-Host "An error occurred: $_"
} finally {
    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Restricted
    Write-Host "Execution policy restored to Restricted."
}
