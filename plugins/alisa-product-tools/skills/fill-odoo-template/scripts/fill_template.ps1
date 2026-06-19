param(
    [string]$Products = "",
    [string]$Name = "",
    [string]$BrandAbbr = "",
    [string]$PN = "",
    [string]$Description = "",
    [string]$Units = "",
    [string]$TemplateExcel = "C:\Users\saran\OneDrive - Alisa intersupply CO.,LTD\Desktop\Product code\product_template to Odoo.xls",
    [string]$BrandExcel    = "C:\Users\saran\OneDrive - Alisa intersupply CO.,LTD\Desktop\Product code\Product code and Brand by Claude.xlsx"
)

if (-not (Test-Path $TemplateExcel)) { Write-Error "Template file not found: $TemplateExcel"; exit 1 }
if (-not (Test-Path $BrandExcel))    { Write-Error "Brand file not found: $BrandExcel"; exit 1 }

$productList = @()
if ($Products.Trim() -ne "") {
    try { $productList = ConvertFrom-Json $Products }
    catch { Write-Error "Invalid JSON in -Products: $($_.Exception.Message)"; exit 1 }
} elseif ($Name.Trim() -ne "") {
    $productList = @([PSCustomObject]@{
        Name        = $Name.Trim()
        BrandAbbr   = $BrandAbbr.Trim().ToUpper()
        PN          = $PN.Trim().ToUpper()
        Description = $Description.Trim()
        Units       = $Units.Trim()
    })
} else {
    Write-Error "Provide -Products (JSON) or individual params"
    exit 1
}

# Copy template to new timestamped file — original stays untouched as blank template
$dir        = [System.IO.Path]::GetDirectoryName($TemplateExcel)
$ext        = [System.IO.Path]::GetExtension($TemplateExcel)
$timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
$newName    = ("product_template to Odoo {0}{1}" -f $timestamp, $ext)
$outputPath = Join-Path $dir $newName
Copy-Item -Path $TemplateExcel -Destination $outputPath -Force

function RGBtoExcel($r, $g, $b) { return [int]($b * 65536 + $g * 256 + $r) }
$colorHeader    = RGBtoExcel 31  31  31
$colorHeaderTxt = RGBtoExcel 255 255 255
$colorOddRow    = RGBtoExcel 255 255 255
$colorEvenRow   = RGBtoExcel 238 238 238
$colorBorder    = RGBtoExcel 192 192 192
$colorText      = RGBtoExcel 26  26  26
$fontName       = "Aptos"

function Apply-SheetTheme($s) {
    if (-not $s) { return }
    $totalRows = $s.UsedRange.Rows.Count
    $totalCols = $s.UsedRange.Columns.Count
    if ($totalRows -lt 1 -or $totalCols -lt 1) { return }
    for ($c = 1; $c -le $totalCols; $c++) {
        $h = $s.Cells.Item(1, $c)
        $h.Font.Name = $fontName; $h.Font.Size = 12; $h.Font.Bold = $true
        $h.Font.Color = $colorHeaderTxt; $h.Interior.Color = $colorHeader
        $h.HorizontalAlignment = -4108; $h.VerticalAlignment = -4108
        $h.Borders.Weight = 2; $h.Borders.Color = $colorBorder
    }
    for ($r = 2; $r -le $totalRows; $r++) {
        $bg = if (($r % 2) -eq 0) { $colorEvenRow } else { $colorOddRow }
        for ($c = 1; $c -le $totalCols; $c++) {
            $cell = $s.Cells.Item($r, $c)
            $cell.Font.Name = $fontName; $cell.Font.Size = 10; $cell.Font.Bold = $true
            $cell.Font.Color = $colorText; $cell.Interior.Color = $bg
            $cell.VerticalAlignment = -4107; $cell.HorizontalAlignment = -4131
            $cell.Borders.Weight = 2; $cell.Borders.Color = $colorBorder
        }
    }
}

$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $wbBrand = $excel.Workbooks.Open($BrandExcel, 0, $true)
    $brandSheet = $null
    foreach ($s in $wbBrand.Sheets) {
        if ($s.Name.Trim() -eq "Brand abbr") { $brandSheet = $s; break }
    }
    if (-not $brandSheet) {
        $wbBrand.Close($false); $excel.Quit()
        Write-Error "Sheet Brand abbr not found"; exit 1
    }
    $brandMap = @{}
    $bRows = $brandSheet.UsedRange.Rows.Count
    for ($r = 2; $r -le $bRows; $r++) {
        $abbr = $brandSheet.Cells.Item($r, 2).Text.Trim().ToUpper()
        $full = $brandSheet.Cells.Item($r, 1).Text.Trim().ToUpper()
        if ($abbr) { $brandMap[$abbr] = $full }
    }
    $wbBrand.Close($false)

    $wbT = $excel.Workbooks.Open($outputPath, 0, $false)
    $tSheet = $null
    foreach ($s in $wbT.Sheets) {
        if ($s.Name.Trim() -eq "Template") { $tSheet = $s; break }
    }
    if (-not $tSheet) {
        $wbT.Close($false); $excel.Quit()
        Write-Error "Sheet Template not found"; exit 1
    }

    $maxID = 0
    $lastDataRow = 1
    $scanRows = $tSheet.UsedRange.Rows.Count
    for ($r = 2; $r -le $scanRows; $r++) {
        $extID = $tSheet.Cells.Item($r, 1).Text.Trim()
        if ($extID -match "product_template_(\d+)") {
            $num = [int]$Matches[1]
            if ($num -gt $maxID) { $maxID = $num; $lastDataRow = $r }
        }
    }
    $nextID  = $maxID + 1
    $nextRow = $lastDataRow + 1

    $results = @()
    foreach ($p in $productList) {
        $pName  = $p.Name.Trim()
        $pAbbr  = $p.BrandAbbr.Trim().ToUpper()
        $pPN    = $p.PN.Trim().ToUpper()
        $pDesc  = $p.Description.Trim()
        $pUnits = $p.Units.Trim()

        if (-not $brandMap.ContainsKey($pAbbr)) {
            Write-Warning ("SKIP: brand {0} not found, skipping {1}" -f $pAbbr, $pName)
            continue
        }
        $brandFull = $brandMap[$pAbbr]
        $descBlock = ("BRAND: {0}`nP/N: {1}`nDESCRIPTION: {2}" -f $brandFull, $pPN, $pDesc)

        $tSheet.Cells.Item($nextRow,  1).Value2 = ("product_template_{0}" -f $nextID)
        $tSheet.Cells.Item($nextRow,  2).Value2 = $pName
        $tSheet.Cells.Item($nextRow,  3).Value2 = "Goods"
        $tSheet.Cells.Item($nextRow,  4).Value2 = 1
        $tSheet.Cells.Item($nextRow,  5).Value2 = ""
        $tSheet.Cells.Item($nextRow,  6).Value2 = 0
        $tSheet.Cells.Item($nextRow,  7).Value2 = $pUnits
        $tSheet.Cells.Item($nextRow,  8).Value2 = 0
        $tSheet.Cells.Item($nextRow,  9).Value2 = 0
        $tSheet.Cells.Item($nextRow, 10).Value2 = $descBlock
        $tSheet.Cells.Item($nextRow, 11).Value2 = $descBlock
        $tSheet.Cells.Item($nextRow, 12).Value2 = $descBlock
        $tSheet.Cells.Item($nextRow, 13).Value2 = $descBlock
        $tSheet.Cells.Item($nextRow, 14).Value2 = "manassaporn.s"
        for ($c = 10; $c -le 13; $c++) { $tSheet.Cells.Item($nextRow, $c).WrapText = $true }

        $results += [PSCustomObject]@{ ID=$nextID; Name=$pName; Brand=$brandFull; PN=$pPN; Units=$pUnits; Row=$nextRow }
        $nextID++
        $nextRow++
    }

    Apply-SheetTheme $tSheet
    $wbT.Save()
    $wbT.Close($false)
    $excel.Quit()

    Write-Output ("SUCCESS: {0} product(s) added" -f $results.Count)
    foreach ($r in $results) {
        Write-Output ("  [{0}] {1}  Brand={2}  PN={3}  Units={4}  Row={5}" -f $r.ID, $r.Name, $r.Brand, $r.PN, $r.Units, $r.Row)
    }
    Write-Output ("FILE: {0}" -f $newName)

} catch {
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    Write-Error ("ERROR: {0}" -f $_.Exception.Message)
    exit 1
} finally {
    if ($null -ne $excel) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
    }
}
