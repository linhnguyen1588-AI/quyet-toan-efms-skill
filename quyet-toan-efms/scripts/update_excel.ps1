$data = @{
    "CAIU7238923" = @{ price="1650000"; inv="00003839"; date="27/05/2026" }
    "SKHU9419310" = @{ price="1650000"; inv="21026"; date="30/05/2026" }
    "SKHU9810583" = @{ price="1650000"; inv="21057"; date="30/05/2026" }
    "EGSU6559029" = @{ price="1488889"; inv="14663"; date="30/05/2026" }
}

$baseDir = Join-Path $env:USERPROFILE "OneDrive"
$targetDir = Get-ChildItem -Path $baseDir -Directory | Where-Object { $_.Name -like "M*y t*nh" } | Select-Object -First 1
$excelPath = Join-Path $targetDir.FullName "TOAN LUC T5 NLP\BUYING TOAN LUC PARITAS.xlsx"

Write-Host "Opening file: $excelPath"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$workbook = $excel.Workbooks.Open($excelPath)
$worksheet = $workbook.Sheets.Item(1)
$lastRow = $worksheet.UsedRange.Rows.Count

for ($i=2; $i -le $lastRow; $i++) {
    $cont = $worksheet.Cells.Item($i, 2).Text
    if ($cont -and $data.ContainsKey($cont)) {
        $worksheet.Cells.Item($i, 9).Value2 = $data[$cont].price
        $worksheet.Cells.Item($i, 15).Value2 = $data[$cont].inv
        $worksheet.Cells.Item($i, 16).Value2 = $data[$cont].date
    }
}

$workbook.Save()
$workbook.Close($true)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
Write-Host "Updated missing containers successfully"
