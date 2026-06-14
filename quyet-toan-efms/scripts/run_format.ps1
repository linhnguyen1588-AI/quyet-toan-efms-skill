$ErrorActionPreference = "Stop"

$dirPath = "D:\SOTRANS LINH 2\LINH\BUYING VT"
$files = Get-ChildItem -Path $dirPath -Filter "*T*I XU*T*.xlsx"

if ($files.Count -eq 0) {
    Write-Host "File not found!"
    exit 1
}

$filePath = $files[0].FullName
Write-Host "Found file: $filePath"

& ".\format_excel.ps1" -filePath $filePath
