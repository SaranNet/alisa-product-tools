---
name: product-import
description: Bulk-add multiple new brands at once to Alisa Intersupply Excel. Use when adding several brands from a price list or PO — creates abbreviations for all, saves all in one session.
---

# Product Import (Bulk Brand Add)

นำเข้าแบรนด์ใหม่หลายรายการพร้อมกันในครั้งเดียว

## Excel File

```
C:\Users\saran\OneDrive - Alisa intersupply CO.,LTD\Desktop\Product code\Product code and Brand by Claude.xlsx
```
(หรือ `$env:PRODUCT_CODE_EXCEL` ถ้าตั้งค่าไว้)

---

## ขั้นตอน

1. รับรายการแบรนด์จากผู้ใช้ (ชื่อเต็ม / ประเภทสินค้า)
2. อ่านรายการแบรนด์ที่มีอยู่ทั้งหมด (Read All Brands)
3. สร้างรหัสย่อสำหรับแบรนด์ใหม่ทุกตัว (ตรวจไม่ซ้ำ)
4. บันทึกทั้งหมดในครั้งเดียว (Write All Brands)
5. แจ้งผลสรุป

---

## Step 1 — Read All Brands

```powershell
$xlPath = if ($env:PRODUCT_CODE_EXCEL) { $env:PRODUCT_CODE_EXCEL } else { "C:\Users\saran\OneDrive - Alisa intersupply CO.,LTD\Desktop\Product code\Product code and Brand by Claude.xlsx" }
$xl = New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false
$wb = $xl.Workbooks.Open($xlPath,0,$true)
$ws = $null; foreach ($s in $wb.Sheets) { if ($s.Name -eq "Brand abbr") { $ws=$s; break } }
$rows=$ws.UsedRange.Rows.Count; $existing=@()
for ($r=2;$r -le $rows;$r++) {
    $existing += [PSCustomObject]@{
        brand=$ws.Cells.Item($r,1).Text.Trim().ToUpper()
        abbr=$ws.Cells.Item($r,2).Text.Trim().ToUpper()
    }
}
$wb.Close($false); $xl.Quit(); [Runtime.InteropServices.Marshal]::ReleaseComObject($xl)|Out-Null
$existing | ConvertTo-Json
```

---

## Step 2 — สร้างรหัสย่อสำหรับแบรนด์ใหม่

สำหรับแต่ละแบรนด์ใหม่ (ที่ยังไม่มีในรายการ):
- ข้ามแบรนด์ที่มีอยู่แล้ว — ไม่ต้องสร้างรหัสซ้ำ
- ใช้อัลกอริทึมเดียวกับ brand-abbr skill (3 ตัว A-Z ไม่ซ้ำ)

| ลำดับ | วิธี | ตัวอย่าง |
|-------|------|---------|
| 1 | 3 ตัวแรก | SAMSUNG → SAM |
| 2 | อักษรตัวแรกแต่ละคำ | ALLEN BRADLEY LTD → ABL |
| 3 | 2 ตัวแรก + ตัวสุดท้าย | SAMSUNG → SAG |
| 4 | ตัวแรก + สระ + พยัญชนะ | SAMSUNG → SMG |
| 5 | ลองผสมจนไม่ซ้ำ | - |

---

## Step 3 — Write All New Brands

```powershell
$xlPath = if ($env:PRODUCT_CODE_EXCEL) { $env:PRODUCT_CODE_EXCEL } else { "C:\Users\saran\OneDrive - Alisa intersupply CO.,LTD\Desktop\Product code\Product code and Brand by Claude.xlsx" }

# --- แทนที่ด้วยรายการแบรนด์ใหม่ที่สร้างจาก Step 2 ---
$newBrands = @(
    [PSCustomObject]@{ brand="BRAND_NAME_1"; abbr="AB1"; product_group="Cable" },
    [PSCustomObject]@{ brand="BRAND_NAME_2"; abbr="AB2"; product_group="Junction box" }
)

$xl = New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false
$wb = $xl.Workbooks.Open($xlPath,0,$false)
$ws = $null; foreach ($s in $wb.Sheets) { if ($s.Name -eq "Brand abbr") { $ws=$s; break } }
$lastRow = $ws.UsedRange.Rows.Count

foreach ($b in $newBrands) {
    $lastRow++
    $ws.Cells.Item($lastRow,1).Value2=$b.brand
    $ws.Cells.Item($lastRow,2).Value2=$b.abbr
    $ws.Cells.Item($lastRow,3).Value2=$b.product_group
}

# Sort A-Z after all additions
$dataRange=$ws.Range($ws.Cells.Item(2,1),$ws.Cells.Item($ws.UsedRange.Rows.Count,3))
$dataRange.Sort($ws.Cells.Item(1,1),1)|Out-Null
$wb.Save(); $wb.Close($false); $xl.Quit(); [Runtime.InteropServices.Marshal]::ReleaseComObject($xl)|Out-Null
Write-Output "SUCCESS: Added $($newBrands.Count) brands"
```

---

## รูปแบบสรุปผล

```
✅ นำเข้าสำเร็จ [N] แบรนด์:
  • BRAND1 → AB1 (Product group)
  • BRAND2 → AB2 (Product group)
  
⏭ ข้ามแบรนด์ที่มีอยู่แล้ว [M] รายการ:
  • EXISTINGBRAND (รหัส EXI มีอยู่แล้ว)
```
