# Tự động lưu và đồng bộ dự án lên GitHub
$gitPath = "d:\workspace-ai\git-portable\cmd\git.exe"

Write-Host "Dang kiem tra trang thai Git..." -ForegroundColor Cyan
& $gitPath status

Write-Host "Dang them cac thay doi vao Git..." -ForegroundColor Cyan
& $gitPath add .

Write-Host "Dang commit cac thay doi..." -ForegroundColor Cyan
$commitMessage = "Auto sync: " + (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
& $gitPath commit -m $commitMessage

Write-Host "Dang day len GitHub..." -ForegroundColor Cyan
& $gitPath push origin main

Write-Host "Dong bo thanh cong!" -ForegroundColor Green
