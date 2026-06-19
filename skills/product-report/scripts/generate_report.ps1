# generate_report.ps1 — Generate summary report from both Excel sheets
# Output: JSON with summary stats, all brands, all groups, brands grouped by product group

param(
    [string]$ExcelPath = $env:PRODUCT_CODE_EXCEL
)

if (-not $ExcelPath) {
    $ExcelPath = "C:\Users\saran\OneDrive - Alisa intersupply CO.,LTD\Desktop\Product code\Product code and Brand by Claude.xlsx"
}

if (-not (Test-Path $ExcelPath)) {
    Write-Error "ERROR: Excel file not found at: $ExcelPath"
    exit 1
}

$tempPath = [System.IO.Path]::Combine(
    [System.IO.Path]::GetTempPath(),
    "report_$([System.Guid]::NewGuid().ToString('N')).xlsx"
)
Copy-Item $ExcelPath $tempPath -Force

$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Open($tempPath, 0, $true)

    # --- Read Brand abbr sheet ---
    $brandSheet = $null
    foreach ($s in $wb.Sheets) { if ($s.Name.Trim() -eq "Brand abbr") { $brandSheet = $s; break } }

    $brands = @()
    $brandsByGroup = @{}

    if ($brandSheet) {
        $rowCount = $brandSheet.UsedRange.Rows.Count
        for ($r = 2; $r -le $rowCount; $r++) {
            $brand = $brandSheet.Cells.Item($r, 1).Text.Trim()
            if (-not $brand) { continue }
            $abbr  = $brandSheet.Cells.Item($r, 2).Text.Trim()
            $group = $brandSheet.Cells.Item($r, 3).Text.Trim()

            $brands += [PSCustomObject]@{
                brand         = $brand
                abbr_brand    = $abbr
                product_group = $group
            }

            # Group brands by product_group
            if (-not $brandsByGroup.ContainsKey($group)) {
                $brandsByGroup[$group] = @()
            }
            $brandsByGroup[$group] += $brand
        }
    }

    # --- Read Product group sheet ---
    $pgSheet = $null
    foreach ($s in $wb.Sheets) { if ($s.Name.Trim() -eq "Product group") { $pgSheet = $s; break } }

    $groups = @()
    if ($pgSheet) {
        $rowCount = $pgSheet.UsedRange.Rows.Count
        for ($r = 2; $r -le $rowCount; $r++) {
            $mainType = $pgSheet.Cells.Item($r, 1).Text.Trim()
            if (-not $mainType) { continue }
            $abbrType = $pgSheet.Cells.Item($r, 2).Text.Trim()

            # Count brands in this group
            $brandCount = 0
            if ($brandsByGroup.ContainsKey($abbrType)) {
                $brandCount = $brandsByGroup[$abbrType].Count
            }

            $groups += [PSCustomObject]@{
                main_type    = $mainType
                abbr_type    = $abbrType
                category     = $pgSheet.Cells.Item($r, 3).Text.Trim()
                product_list = $pgSheet.Cells.Item($r, 4).Text.Trim()
                brand_count  = $brandCount
            }
        }
    }

    $wb.Close($false)
    $excel.Quit()

    # Build final report object
    $report = [PSCustomObject]@{
        summary = [PSCustomObject]@{
            total_brands = $brands.Count
            total_groups = $groups.Count
        }
        brands         = $brands
        groups         = $groups
        brands_by_group = $brandsByGroup
    }

    $report | ConvertTo-Json -Depth 5

} catch {
    Write-Error "ERROR: $($_.Exception.Message)"
    exit 1
} finally {
    if ($null -ne $excel) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
    }
    Remove-Item $tempPath -Force -ErrorAction SilentlyContinue
}
