# save_product.ps1
# Save product lookup record to "Product lookup" sheet, with duplicate check

param(
    [string]$PN          = "",
    [string]$Description = "",
    [Parameter(Mandatory=$true)][string]$ProductType,
    [string]$ExcelPath   = $env:PRODUCT_CODE_EXCEL
)

if (-not $ExcelPath) {
    $ExcelPath = "C:\Users\saran\OneDrive - Alisa intersupply CO.,LTD\Desktop\Product code\Product code and Brand by Claude.xlsx"
}
if (-not (Test-Path $ExcelPath)) {
    Write-Error "ERROR: Excel file not found at: $ExcelPath"; exit 1
}

$PN          = $PN.Trim().ToUpper()
$Description = $Description.Trim()
$ProductType = $ProductType.Trim().ToUpper()

if (-not $PN -and -not $Description) {
    Write-Error "ERROR: ต้องระบุ P/N หรือ Description อย่างน้อย 1 อย่าง"; exit 1
}

# --- Monochrome white/gray theme (BGR format for Excel COM) ---
function RGBtoExcel($r, $g, $b) { return [int]($b * 65536 + $g * 256 + $r) }
$colorHeader    = RGBtoExcel 31  31  31   # #1F1F1F dark charcoal
$colorHeaderTxt = RGBtoExcel 255 255 255  # #FFFFFF white
$colorOddRow    = RGBtoExcel 255 255 255  # #FFFFFF pure white
$colorEvenRow   = RGBtoExcel 238 238 238  # #EEEEEE light gray
$colorBorder    = RGBtoExcel 192 192 192  # #C0C0C0 silver
$colorText      = RGBtoExcel 26  26  26   # #1A1A1A near black
$fontName       = "Aptos"

function Apply-SheetTheme($s) {
    if (-not $s) { return }
    $totalRows = $s.UsedRange.Rows.Count
    $totalCols = $s.UsedRange.Columns.Count
    if ($totalRows -lt 1 -or $totalCols -lt 1) { return }
    # Header
    for ($c = 1; $c -le $totalCols; $c++) {
        $h = $s.Cells.Item(1, $c)
        $h.Font.Name = $fontName; $h.Font.Size = 12; $h.Font.Bold = $true
        $h.Font.Color = $colorHeaderTxt; $h.Interior.Color = $colorHeader
        $h.HorizontalAlignment = -4108; $h.VerticalAlignment = -4108
        $h.Borders.Weight = 2; $h.Borders.Color = $colorBorder
    }
    # Data rows — alternating white/gray
    for ($r = 2; $r -le $totalRows; $r++) {
        $bg = if (($r % 2) -eq 0) { $colorEvenRow } else { $colorOddRow }
        for ($c = 1; $c -le $totalCols; $c++) {
            $cell = $s.Cells.Item($r, $c)
            $cell.Font.Name = $fontName; $cell.Font.Size = 10; $cell.Font.Bold = $true
            $cell.Font.Color = $colorText; $cell.Interior.Color = $bg
            $cell.VerticalAlignment = -4108; $cell.HorizontalAlignment = -4131
            $cell.Borders.Weight = 2; $cell.Borders.Color = $colorBorder
        }
    }
}

$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Open($ExcelPath, 0, $false)

    # --- Find or create "Product lookup" sheet ---
    $sheet = $null
    foreach ($s in $wb.Sheets) { if ($s.Name.Trim() -eq "Product lookup") { $sheet = $s; break } }

    if (-not $sheet) {
        $sheet = $wb.Sheets.Add()
        $sheet.Name = "Product lookup"
        # Write header row
        $headers = @("P/N", "Description", "Product Type")
        for ($c = 1; $c -le 3; $c++) {
            $h = $sheet.Cells.Item(1, $c)
            $h.Value2              = $headers[$c - 1]
            $h.Font.Name           = $fontName
            $h.Font.Size           = 12
            $h.Font.Bold           = $true
            $h.Font.Color          = $colorHeaderTxt
            $h.Interior.Color      = $colorHeader
            $h.HorizontalAlignment = -4108
            $h.VerticalAlignment   = -4108
            $h.Borders.Weight      = 2
            $h.Borders.Color       = $colorBorder
        }
        # Auto-fit columns
        $sheet.Columns.Item(1).ColumnWidth = 20
        $sheet.Columns.Item(2).ColumnWidth = 40
        $sheet.Columns.Item(3).ColumnWidth = 15
    }

    # --- Duplicate check ---
    $lastRow = $sheet.UsedRange.Rows.Count
    for ($r = 2; $r -le $lastRow; $r++) {
        $existPN   = $sheet.Cells.Item($r, 1).Text.Trim().ToUpper()
        $existDesc = $sheet.Cells.Item($r, 2).Text.Trim().ToUpper()

        # Duplicate if P/N matches (when provided), or description matches exactly
        if ($PN -and $existPN -eq $PN) {
            Write-Output "DUPLICATE: P/N '$PN' มีอยู่แล้วที่แถว $r (Type: $($sheet.Cells.Item($r, 3).Text.Trim()))"
            $wb.Close($false); $excel.Quit(); exit 0
        }
        if (-not $PN -and $Description -and $existDesc -eq $Description.ToUpper()) {
            Write-Output "DUPLICATE: Description '$Description' มีอยู่แล้วที่แถว $r (Type: $($sheet.Cells.Item($r, 3).Text.Trim()))"
            $wb.Close($false); $excel.Quit(); exit 0
        }
    }

    # --- Append new row ---
    $nextRow = $lastRow + 1
    $sheet.Cells.Item($nextRow, 1).Value2 = $PN
    $sheet.Cells.Item($nextRow, 2).Value2 = $Description
    $sheet.Cells.Item($nextRow, 3).Value2 = $ProductType

    # --- Reapply white/gray theme to ALL sheets ---
    foreach ($s in $wb.Sheets) { Apply-SheetTheme $s }

    $wb.Save()
    $wb.Close($false)
    $excel.Quit()

    Write-Output "SUCCESS: บันทึก P/N='$PN' Desc='$Description' Type='$ProductType' ที่แถว $nextRow"

} catch {
    Write-Error "ERROR: $($_.Exception.Message)"
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    exit 1
} finally {
    if ($null -ne $excel) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
    }
}
