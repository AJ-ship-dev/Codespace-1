Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

try {
    Import-Module ImportExcel

    $xlsxFolder = ""
    $date = Get-Date -Format "yyyyMMdd"  

    Write-Host "Searching for files in: $xlsxFolder"
    
    $csvFile = Get-ChildItem -Path $xlsxFolder -Filter "simplified_reports__funding.csv"

    if ($csvFile) {
        Write-Host "Found CSV file: $($csvFile.FullName)"

        $newCsvFileName = "Funding Transactions_$date.xlsx"
        $newCsvFilePath = Join-Path $xlsxFolder $newCsvFileName

        $csvData = Import-Csv -Path $csvFile.FullName
        $csvData | Export-Excel -Path $newCsvFilePath -AutoSize -TableName "SimplifiedTable"

        Write-Host "CSV data saved as Excel: $newCsvFilePath"
        
        $destinationPath = ""
        Write-Host "Moving file to destination: $destinationPath"

        Move-Item -Path $newCsvFilePath -Destination $destinationPath

        Write-Host "File moved successfully to $destinationPath"
    } else {
        Write-Host "Error: No CSV file found matching pattern 'simplified_reports__funding.csv' in $xlsxFolder"
    }

} catch {
    Write-Host "An error occurred: $_"
} finally {
    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Restricted
    Write-Host "Execution policy restored to Restricted."
