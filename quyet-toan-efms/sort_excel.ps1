$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$filePath = "D:\workspace-ai\HOA DON\HAOHUA\HAOHUA XUAT EPORT.xlsx"

try {
    $workbook = $excel.Workbooks.Open($filePath)
    $worksheet = $workbook.Sheets.Item(1)

    $lastRow = $worksheet.UsedRange.Rows.Count

    # Fill down Booking (Cột 2)
    for ($i = 2; $i -le $lastRow; $i++) {
        $val = $worksheet.Cells.Item($i, 2).Value2
        if ([string]::IsNullOrWhiteSpace($val)) {
            $worksheet.Cells.Item($i, 2).Value2 = $worksheet.Cells.Item($i - 1, 2).Value2
        }
    }

    # Sort Range by Booking
    $range = $worksheet.Range("A1:I$lastRow")
    # 1=xlAscending, 1=xlYes (Header)
    $range.Sort($worksheet.Range("B2"), 1, $null, $null, 1, $null, 1, 1)

    $workbook.Save()
    $workbook.Close()
    Write-Host "Success"
} catch {
    Write-Host "Error: $_"
    if ($workbook) { $workbook.Close($false) }
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
