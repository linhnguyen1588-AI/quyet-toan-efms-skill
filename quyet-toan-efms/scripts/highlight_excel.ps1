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

$processedConts = @("TRHU6112075", "CCLU7677844", "CSNU5510286", "HMMU4284158", "HMMU6016133", "FITU5519190", "HMMU6852576", "TXGU4122556", "BSIU9671780", "CAIU7238923", "SKHU9419310", "SKHU9810583", "EGSU6559029")

for ($i=2; $i -le $lastRow; $i++) {
    $cont = $worksheet.Cells.Item($i, 2).Text.Trim()
    
    if ($cont -ne "" -and $processedConts -notcontains $cont) {
        $worksheet.Rows.Item($i).Interior.Color = 65535 # Bôi vàng dòng
        Write-Host "Highlighted row $i for cont $cont"
    }
}

$workbook.Save()
$workbook.Close($true)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
Write-Host "Highlighted successfully"
