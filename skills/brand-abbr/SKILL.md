---
name: brand-abbr
description: Look up or create a 3-letter brand abbreviation in Alisa Intersupply's Brand abbr Excel sheet. Use when asked about brand codes, registering new brands, or listing all brands. Never duplicate existing codes. Always saves immediately without asking.
---

# Brand Abbreviation Manager

จัดการรหัสย่อแบรนด์ใน **Brand abbr** sheet — ทั้งค้นหาและสร้างใหม่

## Excel File

```
C:\Users\saran\OneDrive - Alisa intersupply CO.,LTD\Desktop\Product code\Product code and Brand by Claude.xlsx
```
(หรือ `$env:PRODUCT_CODE_EXCEL` ถ้าตั้งค่าไว้)

---

## กรณี 1: ค้นหาแบรนด์ที่มีอยู่แล้ว

รัน Read Brands script → ค้นหาชื่อแบรนด์ (case-insensitive) → ถ้าพบแสดงรหัสทันที → ถ้าไม่พบไปกรณี 2 โดยไม่ต้องถาม

```powershell
$xlPath = if ($env:PRODUCT_CODE_EXCEL) { $env:PRODUCT_CODE_EXCEL } else { "C:\Users\saran\OneDrive - Alisa intersupply CO.,LTD\Desktop\Product code\Product code and Brand by Claude.xlsx" }
$xl = New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false
$wb = $xl.Workbooks.Open($xlPath,0,$true)
$ws = $null; foreach ($s in $wb.Sheets) { if ($s.Name -eq "Brand abbr") { $ws=$s; break } }
$rows=$ws.UsedRange.Rows.Count; $brands=@()
for ($r=2;$r -le $rows;$r++) {
    $brands += [PSCustomObject]@{
        brand=$ws.Cells.Item($r,1).Text.Trim()
        abbr=$ws.Cells.Item($r,2).Text.Trim()
        product_group=$ws.Cells.Item($r,3).Text.Trim()
    }
}
$wb.Close($false); $xl.Quit(); [Runtime.InteropServices.Marshal]::ReleaseComObject($xl)|Out-Null
$brands | ConvertTo-Json
```

คำตอบเมื่อพบ:
```
✅ แบรนด์ **[Brand]** มีรหัสย่อแล้ว: **[Abbr]**
Product group: [Product group]
```

---

## กรณี 2: สร้างรหัสย่อใหม่ + บันทึกทันที

1. ดูรายการทั้งหมดจาก Step 1 เพื่อตรวจสอบไม่ซ้ำ
2. สร้างรหัสตามอัลกอริทึมด้านล่าง
3. **รัน Add Brand script ทันที — ไม่ต้องถาม**
4. Script จะเรียง A→Z อัตโนมัติ

### อัลกอริทึมสร้างรหัส (3 ตัวอักษร A-Z เท่านั้น ห้ามซ้ำ)

ลองตามลำดับจนกว่าจะไม่ซ้ำ:

| ลำดับ | วิธี | ตัวอย่าง SAMSUNG |
|-------|------|-----------------|
| 1 | 3 ตัวแรก | SAM |
| 2 | อักษรตัวแรกของแต่ละคำ (ชื่อ ≥ 3 คำ) | ALLEN BRADLEY LTD → ABL |
| 3 | 2 ตัวแรก + ตัวสุดท้าย | SAG |
| 4 | ตัวแรก + สระที่ 1 + พยัญชนะที่ 2 | SMG |
| 5 | ลองผสมตัวอักษรจากชื่อจนหาตัวที่ไม่ซ้ำ | - |

```powershell
$xlPath = if ($env:PRODUCT_CODE_EXCEL) { $env:PRODUCT_CODE_EXCEL } else { "C:\Users\saran\OneDrive - Alisa intersupply CO.,LTD\Desktop\Product code\Product code and Brand by Claude.xlsx" }
$brandName    = "[ชื่อแบรนด์เต็ม UPPERCASE]"
$abbr         = "[รหัสย่อ 3 ตัว UPPERCASE]"
$productGroup = "[ประเภทสินค้า ถ้ามี]"

$xl = New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false
$wb = $xl.Workbooks.Open($xlPath,0,$false)
$ws = $null; foreach ($s in $wb.Sheets) { if ($s.Name -eq "Brand abbr") { $ws=$s; break } }

# Check duplicate
$rows=$ws.UsedRange.Rows.Count
for ($r=2;$r -le $rows;$r++) {
    $existAbbr = $ws.Cells.Item($r,2).Text.Trim().ToUpper()
    $existBrand = $ws.Cells.Item($r,1).Text.Trim().ToUpper()
    if ($existAbbr -eq $abbr) { Write-Error "DUPLICATE: abbr $abbr already used by $existBrand"; $wb.Close($false); $xl.Quit(); exit 1 }
    if ($existBrand -eq $brandName) { Write-Error "DUPLICATE: brand $brandName already exists with abbr $existAbbr"; $wb.Close($false); $xl.Quit(); exit 1 }
}

# Add new row then sort A-Z
$newRow = $rows+1
$ws.Cells.Item($newRow,1).Value2=$brandName; $ws.Cells.Item($newRow,2).Value2=$abbr; $ws.Cells.Item($newRow,3).Value2=$productGroup
$dataRange=$ws.Range($ws.Cells.Item(2,1),$ws.Cells.Item($ws.UsedRange.Rows.Count,3))
$dataRange.Sort($ws.Cells.Item(1,1),1)|Out-Null
$wb.Save(); $wb.Close($false); $xl.Quit(); [Runtime.InteropServices.Marshal]::ReleaseComObject($xl)|Out-Null
Write-Output "SUCCESS: Added $brandName ($abbr)"
```

### คำตอบหลังบันทึกสำเร็จ

```
✅ บันทึกสำเร็จ!
แบรนด์: **[Brand]** → รหัสย่อ: **[Abbr]**
```

---

## ข้อควรระวัง

- รหัสต้อง **3 ตัวอักษร A-Z** เท่านั้น (ห้ามมีตัวเลขหรืออักขระพิเศษ)
- ถ้า Excel เปิดค้างอยู่ script จะ error → ปิด Excel ก่อนแล้วลองใหม่
- ชื่อแบรนด์ที่บันทึกจะเป็น **UPPERCASE** เสมอ
