$extraData = @{
    "HMMU4284158" = @{ price="132407"; inv="3255"; date="15/05/2026" }
    "HMMU6016133" = @{ price="132407"; inv="3257"; date="15/05/2026" }
    "HMMU6852576" = @{ price="132407"; inv="3343"; date="18/05/2026" }
    "TXGU4122556" = @{ price="132407"; inv="3347"; date="18/05/2026" }
    "FITU5519190" = @{ price="1782407"; inv="173752"; date="16/05/2026" }
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

for ($i = $lastRow; $i -ge 2; $i--) {
    $cont = $worksheet.Cells.Item($i, 2).Text.Trim()
    
    if ($cont -and $extraData.ContainsKey($cont)) {
        # Insert new row below
        $insertRow = $i + 1
        $worksheet.Rows.Item($insertRow).Insert()
        
        # Copy basic columns (1 to 8)
        for ($c = 1; $c -le 8; $c++) {
            $worksheet.Cells.Item($insertRow, $c).Value2 = $worksheet.Cells.Item($i, $c).Value2
        }
        
        # Fill new data
        $worksheet.Cells.Item($insertRow, 9).Value2 = $extraData[$cont].price
        $worksheet.Cells.Item($insertRow, 15).Value2 = $extraData[$cont].inv
        $worksheet.Cells.Item($insertRow, 16).Value2 = $extraData[$cont].date
        
        # Highlight new row
        $worksheet.Rows.Item($insertRow).Interior.Color = 65535
        
        Write-Host "Inserted and highlighted row for $cont"
        
        # Remove from dictionary so we only do it once per container
        $extraData.Remove($cont)
    }
}

$workbook.Save()
$workbook.Close($true)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
Write-Host "Duplicated and highlighted successfully"
