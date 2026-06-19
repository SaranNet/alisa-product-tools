---
name: product-report
description: Generate a summary report of all brands and product groups from Alisa Intersupply Excel. Use when asked to "summarize products", "list all brands", "how many brands", "show all product types", or any overview/report request.
---

# Product Report

สรุปและแสดงรายงานข้อมูลแบรนด์และ product group ทั้งหมดจาก Excel

## Excel File

```
C:\Users\saran\OneDrive - Alisa intersupply CO.,LTD\Desktop\Product code\Product code and Brand by Claude.xlsx
```
(หรือ `$env:PRODUCT_CODE_EXCEL` ถ้าตั้งค่าไว้)

---

## Read All Data

รัน script นี้เพื่อดึงข้อมูลทั้งหมด:

```powershell
$xlPath = if ($env:PRODUCT_CODE_EXCEL) { $env:PRODUCT_CODE_EXCEL } else { "C:\Users\saran\OneDrive - Alisa intersupply CO.,LTD\Desktop\Product code\Product code and Brand by Claude.xlsx" }
$xl = New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false
$wb = $xl.Workbooks.Open($xlPath,0,$true)

# Read brands
$wsBrand=$null; foreach ($s in $wb.Sheets) { if ($s.Name -eq "Brand abbr") { $wsBrand=$s; break } }
$bRows=$wsBrand.UsedRange.Rows.Count; $brands=@()
for ($r=2;$r -le $bRows;$r++) {
    $brands += [PSCustomObject]@{
        brand=$wsBrand.Cells.Item($r,1).Text.Trim()
        abbr=$wsBrand.Cells.Item($r,2).Text.Trim()
        product_group=$wsBrand.Cells.Item($r,3).Text.Trim()
    }
}

# Read product groups
$wsGroup=$null; foreach ($s in $wb.Sheets) { if ($s.Name -eq "Product group") { $wsGroup=$s; break } }
$gRows=$wsGroup.UsedRange.Rows.Count; $groups=@()
for ($r=2;$r -le $gRows;$r++) {
    $groups += [PSCustomObject]@{
        abbr=$wsGroup.Cells.Item($r,1).Text.Trim()
        main_type=$wsGroup.Cells.Item($r,2).Text.Trim()
        category=$wsGroup.Cells.Item($r,3).Text.Trim()
    }
}

$wb.Close($false); $xl.Quit(); [Runtime.InteropServices.Marshal]::ReleaseComObject($xl)|Out-Null
Write-Output "=== BRANDS ($($brands.Count)) ==="; $brands | ConvertTo-Json
Write-Output "=== PRODUCT GROUPS ($($groups.Count)) ==="; $groups | ConvertTo-Json
```

---

## รูปแบบรายงาน

หลังได้ข้อมูลแล้ว แสดงผลดังนี้:

### สรุปภาพรวม
```
📊 Alisa Intersupply — Product Database Summary
แบรนด์ทั้งหมด: [N] รายการ
Product Group: [M] กลุ่ม
```

### ตารางแบรนด์ (เรียง A-Z)
| Brand | Abbr | Product Group |
|-------|------|---------------|
| ABB | ABB | LV equipments, Cable gland |
| ... | ... | ... |

### ตาราง Product Group
| Abbr | Main Type | Category |
|------|-----------|----------|
| CA | Cable | Wiring |
| ... | ... | ... |

---

## ตัวกรองพิเศษ

ถ้าผู้ใช้ถามเฉพาะเจาะจง เช่น "แบรนด์ประเภท Cable มีอะไรบ้าง" → filter จาก product_group field
ถ้าถามว่า "brand code ของ ABB คืออะไร" → ตอบเฉพาะแบรนด์นั้น ไม่ต้องแสดงทั้งหมด
