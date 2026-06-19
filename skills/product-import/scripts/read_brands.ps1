# read_brands.ps1 — Read Brand abbr sheet from Excel
# Output: JSON array of all brands with their abbreviations

param(
    [string]$ExcelPath = $env:PRODUCT_CODE_EXCEL
)

if (-not $ExcelPath) {
    $ExcelPath = "C:\Users\saran\OneDrive - Alisa intersupply CO.,LTD\Desktop\Product code\Product code and Brand by Claude.xlsx"
}

if (-not (Test-Path $ExcelPath)) {
    Write-Error "ERROR: Excel file not found at: $ExcelPath"
    Write-Error "Set env var PRODUCT_CODE_EXCEL to the correct path."
    exit 1
}

$tempPath = [System.IO.Path]::Combine(
    [System.IO.Path]::GetTempPath(),
    "br_read_$([System.Guid]::NewGuid().ToString('N')).xlsx"
)
Copy-Item $ExcelPath $tempPath -Force

$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Open($tempPath, 0, $true)

    $sheet = $null
    foreach ($s in $wb.Sheets) {
        if ($s.Name.Trim() -eq "Brand abbr") { $sheet = $s; break }
    }
    if (-not $sheet) {
        Write-Error "Sheet 'Brand abbr' not found."
        $wb.Close($false); $excel.Quit(); exit 1
    }

    $ur = $sheet.UsedRange
    $rowCount = $ur.Rows.Count
    $results = @()

    for ($r = 2; $r -le $rowCount; $r++) {
        $brand = $sheet.Cells.Item($r, 1).Text.Trim()
        if (-not $brand) { continue }
        $results += [PSCustomObject]@{
            brand         = $brand
            abbr_brand    = $sheet.Cells.Item($r, 2).Text.Trim()
            product_group = $sheet.Cells.Item($r, 3).Text.Trim()
            row           = $r
        }
    }

    $wb.Close($false)
    $excel.Quit()
    $results | ConvertTo-Json -Depth 3

} catch {
    Write-Error "ERROR: $($_.Exception.Message)"
    exit 1
} finally {
    if ($null -ne $excel) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
    }
    Remove-Item $tempPath -Force -ErrorAction SilentlyContinue
}
