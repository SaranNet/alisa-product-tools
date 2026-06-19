# add_brand.ps1
# Add new brand abbreviation, sort A-Z, reapply monochrome handwriting theme to all rows

param(
    [Parameter(Mandatory=$true)][string]$BrandName,
    [Parameter(Mandatory=$true)][string]$Abbr,
    [string]$ProductGroup = "",
    [string]$ExcelPath = $env:PRODUCT_CODE_EXCEL
)

if (-not $ExcelPath) {
    $ExcelPath = "C:\Users\saran\OneDrive - Alisa intersupply CO.,LTD\Desktop\Product code\Product code and Brand by Claude.xlsx"
}
if (-not (Test-Path $ExcelPath)) {
    Write-Error "ERROR: Excel file not found at: $ExcelPath"
    exit 1
}

$BrandName = $BrandName.Trim().ToUpper()
$Abbr      = $Abbr.Trim().ToUpper()

if ($Abbr -notmatch '^[A-Z]{3}$') {
    Write-Error "ERROR: Abbreviation '$Abbr' must be exactly 3 uppercase letters (A-Z only)."
    exit 1
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

    $sheet = $null
    foreach ($s in $wb.Sheets) {
        if ($s.Name.Trim() -eq "Brand abbr") { $sheet = $s; break }
    }
    if (-not $sheet) {
        Write-Error "ERROR: Sheet 'Brand abbr' not found."
        $wb.Close($false); $excel.Quit(); exit 1
    }

    $rowCount = $sheet.UsedRange.Rows.Count

    # --- Duplicate check ---
    for ($r = 2; $r -le $rowCount; $r++) {
        $existBrand = $sheet.Cells.Item($r, 1).Text.Trim().ToUpper()
        $existAbbr  = $sheet.Cells.Item($r, 2).Text.Trim().ToUpper()
        if ($existBrand -eq $BrandName) {
            Write-Error "ERROR: Brand '$BrandName' already exists with abbreviation '$existAbbr' at row $r."
            $wb.Close($false); $excel.Quit(); exit 1
        }
        if ($existAbbr -eq $Abbr) {
            Write-Error "ERROR: Abbreviation '$Abbr' is already used by '$existBrand' at row $r."
            $wb.Close($false); $excel.Quit(); exit 1
        }
    }

    # --- Append new brand at next row ---
    $nextRow = $rowCount + 1
    $sheet.Cells.Item($nextRow, 1).Value2 = $BrandName
    $sheet.Cells.Item($nextRow, 2).Value2 = $Abbr
    $sheet.Cells.Item($nextRow, 3).Value2 = $ProductGroup

    # --- Sort all data A-Z by Brand name (column A), excluding header ---
    $lastRow  = $sheet.UsedRange.Rows.Count
    $dataRange = $sheet.Range("A2:C$lastRow")
    $keyRange  = $sheet.Range("A2:A$lastRow")
    $dataRange.Sort($keyRange, 1) | Out-Null  # 1 = xlAscending

    # --- Reapply white/gray theme to ALL sheets ---
    foreach ($s in $wb.Sheets) { Apply-SheetTheme $s }

    $wb.Save()
    $wb.Close($false)
    $excel.Quit()

    Write-Output "SUCCESS: Added $BrandName ($Abbr) — sorted A-Z — saved to '$ExcelPath'"

} catch {
    Write-Error "ERROR: $($_.Exception.Message)"
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    exit 1
} finally {
    if ($null -ne $excel) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
    }
}
