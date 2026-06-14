param (
    [Parameter(Mandatory=$true)][string]$filePath
)
$ErrorActionPreference = "Stop"

Write-Host "Opening Excel file: $filePath"
if (!(Test-Path $filePath)) {
    Write-Host "File not found!"
    exit 1
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $workbook = $excel.Workbooks.Open($filePath)
    
    foreach ($sheet in $workbook.Worksheets) {
        Write-Host "Processing sheet: $($sheet.Name)"
        $sheet.Activate()
        
        $excel.ActiveWindow.DisplayGridlines = $false
        $excel.ActiveWindow.FreezePanes = $false
        
        # Unmerge everything first (except if we want to preserve some structure, but it's safer to just clear formatting on the whole sheet and rebuild)
        $usedRange = $sheet.UsedRange
        
        # Reset row heights and column widths to default before auto-fitting
        $usedRange.WrapText = $false
        $usedRange.VerticalAlignment = -4108 # Center vertically
        
        # Find header row
        $headerRow = 1
        for ($r = 1; $r -le 10; $r++) {
            $count = 0
            for ($c = 1; $c -le 15; $c++) {
                if ($sheet.Cells.Item($r, $c).Value() -ne $null) { $count++ }
            }
            if ($count -gt 2) {
                $headerRow = $r
                break
            }
        }
        
        $maxCol = 1
        for ($c = 15; $c -ge 1; $c--) {
            if ($sheet.Cells.Item($headerRow, $c).Value() -ne $null) {
                $maxCol = $c
                break
            }
        }
        
        $maxRow = $usedRange.Rows.Count
        if ($maxRow -lt $headerRow) { $maxRow = $headerRow + 10 }
        
        # 1. Format Title
        $titleText = [System.Uri]::UnescapeDataString('B%E1%BA%A2NG%20K%C3%8A%20T%C3%81I%20XU%E1%BA%A4T%20GN-KH%E1%BB%90I%20C%C6%AF%E1%BB%9BC')
        # Clear rows 1 to headerRow - 1
        if ($headerRow -gt 1) {
            for ($r = 1; $r -lt $headerRow; $r++) {
                for ($c = 1; $c -le $maxCol; $c++) {
                    $sheet.Cells.Item($r, $c).Value = $null
                    $sheet.Cells.Item($r, $c).UnMerge()
                    $sheet.Cells.Item($r, $c).Borders.LineStyle = -4142 # xlNone
                    $sheet.Cells.Item($r, $c).Interior.ColorIndex = 0 # No fill
                }
            }
        }
        
        # Set new title at row 1
        $sheet.Cells.Item(1, 1).Value = $titleText
        $titleRange = $sheet.Range($sheet.Cells.Item(1, 1), $sheet.Cells.Item(1, $maxCol))
        $titleRange.Merge()
        $titleRange.Font.Size = 16
        $titleRange.Font.Bold = $true
        $titleRange.Font.Color = 7884319 # Dark Blue (1F4E78)
        $titleRange.HorizontalAlignment = -4108 # Center
        $titleRange.VerticalAlignment = -4108 # Center
        $titleRange.RowHeight = 30
        
        # 2. Format Table Header
        $headerRange = $sheet.Range($sheet.Cells.Item($headerRow, 1), $sheet.Cells.Item($headerRow, $maxCol))
        $headerRange.Font.Bold = $true
        $headerRange.Font.Color = 16777215 # White
        $headerRange.Interior.Color = 7884319 # Dark Blue
        $headerRange.HorizontalAlignment = -4108 # Center
        $headerRange.VerticalAlignment = -4108 # Center
        $headerRange.WrapText = $true
        $headerRange.RowHeight = 25
        
        # Find special columns
        $dateColIdx = -1
        $tkColIdx = -1
        for ($c = 1; $c -le $maxCol; $c++) {
            $colHeader = $sheet.Cells.Item($headerRow, $c).Value()
            if ($colHeader -ne $null) {
                $str = $colHeader.ToString().ToLower()
                if ($str -match "ngày") { $dateColIdx = $c }
                if ($str -match "số tk") { $tkColIdx = $c }
            }
        }
        
        # 3. Format Data Rows
        if ($maxRow -gt $headerRow) {
            $dataRange = $sheet.Range($sheet.Cells.Item($headerRow + 1, 1), $sheet.Cells.Item($maxRow, $maxCol))
            
            # Apply Thin Borders
            $dataRange.Borders.LineStyle = 1
            $headerRange.Borders.LineStyle = 1
            
            for ($r = $headerRow + 1; $r -le $maxRow; $r++) {
                $rowRange = $sheet.Range($sheet.Cells.Item($r, 1), $sheet.Cells.Item($r, $maxCol))
                
                # Zebra striping
                if (($r - $headerRow) % 2 -eq 0) {
                    $rowRange.Interior.Color = 16445670 # Light Blue E6F0FA
                } else {
                    $rowRange.Interior.Color = 16777215 # White
                }
                
                # Specific formats
                for ($c = 1; $c -le $maxCol; $c++) {
                    $cell = $sheet.Cells.Item($r, $c)
                    $colHeader = $sheet.Cells.Item($headerRow, $c).Value()
                    
                    if ($colHeader -ne $null) {
                        $colHeaderStr = $colHeader.ToString().ToLower()
                        
                        # Format Money
                        if ($colHeaderStr -match "phí" -or $colHeaderStr -match "thu" -or $colHeaderStr -match "tiền" -or $colHeaderStr -match "tờ khai") {
                            $cell.HorizontalAlignment = -4152 # Right
                            if ($cell.Value() -is [ValueType]) {
                                $cell.NumberFormat = "#,##0"
                            }
                        } 
                        # Center STT, Số lượng
                        elseif ($colHeaderStr -match "stt" -or $colHeaderStr -match "số lượng") {
                            $cell.HorizontalAlignment = -4108 # Center
                        }
                    }
                }
                
                # Format Date Column
                if ($dateColIdx -ne -1) {
                    $dateCell = $sheet.Cells.Item($r, $dateColIdx)
                    if ($dateCell.Value() -ne $null) {
                        # Ensure date is stored as DateTime if it's text
                        if ($dateCell.Value() -is [string]) {
                            try {
                                $parsedDate = [datetime]::ParseExact($dateCell.Value().Trim(), @('d/M/yyyy', 'dd/MM/yyyy', 'MM/dd/yyyy', 'd/M/yyyy h:mm:ss tt'), [System.Globalization.CultureInfo]::InvariantCulture)
                                $dateCell.Value = $parsedDate
                            } catch { }
                        }
                        $dateCell.NumberFormat = "dd/MM/yyyy"
                        $dateCell.HorizontalAlignment = -4108 # Center
                    }
                }
                
                # Format SỐ TK Column
                if ($tkColIdx -ne -1) {
                    $sheet.Cells.Item($r, $tkColIdx).WrapText = $true
                }
            }
            
            # Thick Border around Table
            $tableRange = $sheet.Range($sheet.Cells.Item($headerRow, 1), $sheet.Cells.Item($maxRow, $maxCol))
            $tableRange.BorderAround(1, 3) # xlContinuous, xlThick
        }
        
        # 4. Pure AutoFit (No extra paddings)
        $usedRange.Columns.AutoFit()
        $usedRange.Rows.AutoFit()
        
        # Freeze panes
        $excel.ActiveWindow.SplitRow = $headerRow
        $excel.ActiveWindow.SplitColumn = 0
        $excel.ActiveWindow.FreezePanes = $true
    }
    
    $workbook.Save()
    Write-Host "File saved successfully."

} catch {
    Write-Host "An error occurred: $_"
} finally {
    if ($workbook) {
        $workbook.Close()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
