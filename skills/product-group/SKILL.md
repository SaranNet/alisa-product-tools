---
name: product-group
description: Look up the correct Product Type abbreviation (Abbr. type) from Alisa Intersupply Excel — always pick ONE group. Use when asked about product type codes, "what type is [product]", "product group for [name]", or to list all groups.
---

# Product Group Lookup

ตัดสินใจเลือก Product Group ที่เหมาะสมที่สุด **1 กลุ่มเดียว** เสมอ — ไม่แสดงตัวเลือกให้ผู้ใช้เลือกเอง

## Excel File

```
C:\Users\saran\OneDrive - Alisa intersupply CO.,LTD\Desktop\Product code\Product code and Brand by Claude.xlsx
```
(หรือ `$env:PRODUCT_CODE_EXCEL` ถ้าตั้งค่าไว้)

---

## ขั้นตอน

1. รัน **Read Groups** script เพื่อดึง product groups ทั้งหมด
2. เลือก **1 กลุ่ม** ตามเกณฑ์ด้านล่าง
3. ตอบด้วย Abbr. type ทันที
4. รัน **Save Product** script เพื่อบันทึกลง "Product lookup" sheet

---

## Step 1 — Read Groups

```powershell
$xlPath = if ($env:PRODUCT_CODE_EXCEL) { $env:PRODUCT_CODE_EXCEL } else { "C:\Users\saran\OneDrive - Alisa intersupply CO.,LTD\Desktop\Product code\Product code and Brand by Claude.xlsx" }
$xl = New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false
$wb = $xl.Workbooks.Open($xlPath,0,$true)
$ws = $null; foreach ($s in $wb.Sheets) { if ($s.Name -eq "Product group") { $ws=$s; break } }
$rows = $ws.UsedRange.Rows.Count; $groups=@()
for ($r=2;$r -le $rows;$r++) {
    $groups += [PSCustomObject]@{
        abbr=$ws.Cells.Item($r,1).Text.Trim(); main_type=$ws.Cells.Item($r,2).Text.Trim()
        category=$ws.Cells.Item($r,3).Text.Trim(); product_list=$ws.Cells.Item($r,4).Text.Trim()
    }
}
$wb.Close($false); $xl.Quit(); [Runtime.InteropServices.Marshal]::ReleaseComObject($xl)|Out-Null
$groups | ConvertTo-Json
```

---

## Step 2 — เกณฑ์เลือก 1 กลุ่ม (เรียงลำดับ)

1. **ตรงชื่อ main_type เป๊ะ** → เลือกทันที
2. **อยู่ใน product_list** ของกลุ่มใด → เลือกกลุ่มนั้น
3. **Category ตรงกัน** → เลือกกลุ่ม specific ที่สุด
4. **ใกล้เคียงที่สุด** จากความรู้ไฟฟ้า/อุตสาหกรรม
5. ไม่มีกลุ่มใดเหมาะ → **OT** (Others)

| สินค้า | กลุ่ม | เหตุผล |
|--------|-------|--------|
| Heat shrink tube | CA | อยู่ใน product_list ของ Cable |
| Solenoid valve | OT | ไม่มีกลุ่มที่ตรง |
| Din rail | LV | อุปกรณ์เสริม panel ไฟฟ้า |
| Cable marker | IA | ระบุใน product_list |

---

## Step 3 — รูปแบบคำตอบ

กรณีปกติ:
```
✅ **[Abbr]**
[Main type] — [Category]
```

กรณี Others:
```
⬜ **OT**
Others — ไม่มีกลุ่มสินค้าที่ตรงกับ [ชื่อสินค้า]
```

---

## Step 4 — Save Product (รันเสมอหลังตอบ)

ดึง P/N และ Description จากคำถาม ถ้าไม่มี P/N ให้ใช้ description จากสินค้าที่ถาม

```powershell
$xlPath = if ($env:PRODUCT_CODE_EXCEL) { $env:PRODUCT_CODE_EXCEL } else { "C:\Users\saran\OneDrive - Alisa intersupply CO.,LTD\Desktop\Product code\Product code and Brand by Claude.xlsx" }
$pn          = "[P/N ถ้ามี — ว่างได้]"
$description = "[ชื่อสินค้าหรือ description ย่อๆ]"
$productType = "[Abbr. type ที่เลือก เช่น JB]"

$xl = New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false
$wb = $xl.Workbooks.Open($xlPath,0,$false)
$ws = $null
foreach ($s in $wb.Sheets) { if ($s.Name -eq "Product lookup") { $ws=$s; break } }
if (-not $ws) {
    $ws = $wb.Sheets.Add(); $ws.Name="Product lookup"
    $ws.Cells.Item(1,1).Value2="P/N"; $ws.Cells.Item(1,2).Value2="Description"; $ws.Cells.Item(1,3).Value2="Product Type"
}
# Check duplicate
$lastRow=1; $exists=$false
$usedRows=$ws.UsedRange.Rows.Count
for ($r=2;$r -le $usedRows;$r++) {
    if ($ws.Cells.Item($r,1).Text.Trim().ToUpper() -eq $pn.ToUpper()) { $exists=$true; break }
    $lastRow=$r
}
if ($exists) { Write-Output "SKIP: P/N $pn already exists" }
else {
    $newRow=$lastRow+1
    $ws.Cells.Item($newRow,1).Value2=$pn; $ws.Cells.Item($newRow,2).Value2=$description; $ws.Cells.Item($newRow,3).Value2=$productType
    $wb.Save(); Write-Output "SAVED: $pn | $description | $productType at row $newRow"
}
$wb.Close($false); $xl.Quit(); [Runtime.InteropServices.Marshal]::ReleaseComObject($xl)|Out-Null
```

---

## แนะนำเพิ่ม Product Group ใหม่

หลังตอบรหัสย่อ ประเมิน **เฉพาะกรณีที่ไม่ตรงชัดเจน** (กรณี 4 หรือ 5):

- สินค้าประเภทนี้มีหลายรายการในระบบ + ไม่มีกลุ่มที่เหมาะ → แนะนำเพิ่มกลุ่มใหม่
- เป็นสินค้าเบ็ดเตล็ดจริงๆ → ไม่ต้องแนะนำ

```
💡 แนะนำ: สินค้าประเภท [ชื่อ] อาจเหมาะกับการเพิ่ม Product Group ใหม่
ต้องการให้สร้าง Product Group ใหม่ไหม?
```
