# read_products.ps1 — Read Product group sheet from Excel
# Output: JSON array of all product groups

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
    "pg_read_$([System.Guid]::NewGuid().ToString('N')).xlsx"
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
        if ($s.Name.Trim() -eq "Product group") { $sheet = $s; break }
    }
    if (-not $sheet) {
        Write-Error "Sheet 'Product group' not found."
        $wb.Close($false); $excel.Quit(); exit 1
    }

    $rowCount = $sheet.UsedRange.Rows.Count
    $results = @()

    for ($r = 2; $r -le $rowCount; $r++) {
        $mainType = $sheet.Cells.Item($r, 1).Text.Trim()
        if (-not $mainType) { continue }
        $results += [PSCustomObject]@{
            main_type    = $mainType
            abbr_type    = $sheet.Cells.Item($r, 2).Text.Trim()
            category     = $sheet.Cells.Item($r, 3).Text.Trim()
            product_list = $sheet.Cells.Item($r, 4).Text.Trim()
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
