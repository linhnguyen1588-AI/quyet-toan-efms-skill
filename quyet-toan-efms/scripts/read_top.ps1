$ErrorActionPreference = "Stop"

$dirPath = "D:\SOTRANS LINH 2\LINH\BUYING VT"
$files = Get-ChildItem -Path $dirPath -Filter "*T*I XU*T*.xlsx"
$filePath = $files[0].FullName

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $workbook = $excel.Workbooks.Open($filePath)
    $sheet = $workbook.Worksheets.Item(1)
    
    for ($r = 1; $r -le 7; $r++) {
        $rowStr = ""
        for ($c = 1; $c -le 10; $c++) {
            $val = $sheet.Cells.Item($r, $c).Text
            if ($val -eq $null) { $val = "" }
            $rowStr += "[$val]`t"
        }
        Write-Host "Row $($r) - $rowStr"
    }
} finally {
    if ($workbook) { $workbook.Close($false) }
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
