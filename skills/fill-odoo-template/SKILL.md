---
name: fill-odoo-template
description: Fill Odoo product import template from [TYPE]PN[BRAND] format. Use when adding products to Odoo — parses product code like [JB]ENCA36N30BLP[HOF], looks up brand full name, fills all required columns, saves as a new timestamped file (original template stays blank).
---

# Fill Odoo Template

เขียนข้อมูล product ลงไฟล์ Odoo import template จากรหัส `[TYPE]PN[BRAND]`

## รูปแบบรหัส Product

```
[TYPE]PARTNUMBER[BRAND]
```

ตัวอย่าง: `[JB]ENCA36N30BLP[HOF]`

| ส่วน | ตัวอย่าง | ความหมาย |
|------|---------|-----------|
| `[JB]` | Product Type Abbr | รหัสย่อประเภทสินค้า |
| `ENCA36N30BLP` | Part Number | รหัสสินค้า |
| `[HOF]` | Brand Abbr | รหัสย่อแบรนด์ |

---

## Excel Files

```
Template : C:\Users\saran\OneDrive - Alisa intersupply CO.,LTD\Desktop\Product code\product_template to Odoo.xls
Brand DB : C:\Users\saran\OneDrive - Alisa intersupply CO.,LTD\Desktop\Product code\Product code and Brand by Claude.xlsx
```

---

## ขั้นตอน

1. **Parse** รหัส `[TYPE]PN[BRAND]` จากข้อมูลที่ผู้ใช้ให้มา
2. **รัน fill script** ด้านล่าง — สร้างไฟล์ใหม่ทุกครั้ง ไม่ทับ template
3. **แจ้งชื่อไฟล์** ที่สร้าง

---

## Parse รหัส [TYPE]PN[BRAND]

จากรหัสที่ผู้ใช้ให้มา ให้ดึงข้อมูลดังนี้:
- `Name` = ชื่อสินค้า (Column Name ใน Odoo) รูปแบบ: `[TYPE]PN[BRAND]`  
- `BrandAbbr` = รหัสย่อแบรนด์ เช่น `HOF`
- `PN` = Part number เช่น `ENCA36N30BLP`
- `Description` = คำอธิบายสินค้า (ถ้าผู้ใช้ระบุมา)
- `Units` = หน่วย เช่น `EA`, `m`, `pcs`

ถ้ามีหลาย product ให้รวมในรัน script เดียว (batch mode)

---

## Fill Script

```powershell
param(
    [string]$Products = "",   # JSON array สำหรับ batch mode
    [string]$Name = "",       # single mode
    [string]$BrandAbbr = "",
    [string]$PN = "",
    [string]$Description = "",
    [string]$Units = "",
    [string]$TemplateExcel = "C:\Users\saran\OneDrive - Alisa intersupply CO.,LTD\Desktop\Product code\product_template to Odoo.xls",
    [string]$BrandExcel    = "C:\Users\saran\OneDrive - Alisa intersupply CO.,LTD\Desktop\Product code\Product code and Brand by Claude.xlsx"
)

if (-not (Test-Path $TemplateExcel)) { Write-Error "Template not found: $TemplateExcel"; exit 1 }
if (-not (Test-Path $BrandExcel))    { Write-Error "Brand file not found: $BrandExcel"; exit 1 }

$productList = @()
if ($Products.Trim() -ne "") {
    try { $productList = ConvertFrom-Json $Products }
    catch { Write-Error "Invalid JSON: $($_.Exception.Message)"; exit 1 }
} elseif ($Name.Trim() -ne "") {
    $productList = @([PSCustomObject]@{
        Name=$Name.Trim(); BrandAbbr=$BrandAbbr.Trim().ToUpper()
        PN=$PN.Trim().ToUpper(); Description=$Description.Trim(); Units=$Units.Trim()
    })
} else { Write-Error "Provide -Products (JSON) or -Name"; exit 1 }

# Copy template — original stays blank
$dir=[IO.Path]::GetDirectoryName($TemplateExcel); $ext=[IO.Path]::GetExtension($TemplateExcel)
$ts=Get-Date -Format "yyyyMMdd_HHmmss"
$newName=("product_template to Odoo {0}{1}" -f $ts,$ext); $outPath=Join-Path $dir $newName
Copy-Item -Path $TemplateExcel -Destination $outPath -Force

function RGBtoXL($r,$g,$b){[int]($b*65536+$g*256+$r)}
$cHeader=RGBtoXL 31 31 31; $cHeaderTxt=RGBtoXL 255 255 255
$cOdd=RGBtoXL 255 255 255; $cEven=RGBtoXL 238 238 238
$cBorder=RGBtoXL 192 192 192; $cText=RGBtoXL 26 26 26; $font="Aptos"

function Set-Theme($s) {
    if (-not $s) { return }
    $tr=$s.UsedRange.Rows.Count; $tc=$s.UsedRange.Columns.Count
    if ($tr -lt 1 -or $tc -lt 1) { return }
    for ($c=1;$c -le $tc;$c++) {
        $h=$s.Cells.Item(1,$c)
        $h.Font.Name=$font;$h.Font.Size=12;$h.Font.Bold=$true
        $h.Font.Color=$cHeaderTxt;$h.Interior.Color=$cHeader
        $h.HorizontalAlignment=-4108;$h.VerticalAlignment=-4108
        $h.Borders.Weight=2;$h.Borders.Color=$cBorder
    }
    for ($r=2;$r -le $tr;$r++) {
        $bg=if(($r%2)-eq 0){$cEven}else{$cOdd}
        for ($c=1;$c -le $tc;$c++) {
            $cell=$s.Cells.Item($r,$c)
            $cell.Font.Name=$font;$cell.Font.Size=10;$cell.Font.Bold=$true
            $cell.Font.Color=$cText;$cell.Interior.Color=$bg
            $cell.VerticalAlignment=-4107;$cell.HorizontalAlignment=-4131
            $cell.Borders.Weight=2;$cell.Borders.Color=$cBorder
        }
    }
}

$xl=$null
try {
    $xl=New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false

    # Load brand map
    $wbB=$xl.Workbooks.Open($BrandExcel,0,$true)
    $wsB=$null; foreach ($s in $wbB.Sheets) { if ($s.Name.Trim() -eq "Brand abbr") { $wsB=$s; break } }
    if (-not $wsB) { $wbB.Close($false); $xl.Quit(); Write-Error "Sheet Brand abbr not found"; exit 1 }
    $brandMap=@{}; $bRows=$wsB.UsedRange.Rows.Count
    for ($r=2;$r -le $bRows;$r++) {
        $ab=$wsB.Cells.Item($r,2).Text.Trim().ToUpper()
        $fn=$wsB.Cells.Item($r,1).Text.Trim().ToUpper()
        if ($ab) { $brandMap[$ab]=$fn }
    }
    $wbB.Close($false)

    # Open output copy
    $wbT=$xl.Workbooks.Open($outPath,0,$false)
    $wsT=$null; foreach ($s in $wbT.Sheets) { if ($s.Name.Trim() -eq "Template") { $wsT=$s; break } }
    if (-not $wsT) { $wbT.Close($false); $xl.Quit(); Write-Error "Sheet Template not found"; exit 1 }

    # Find next row & ID
    $maxID=0; $lastDataRow=1; $scanRows=$wsT.UsedRange.Rows.Count
    for ($r=2;$r -le $scanRows;$r++) {
        $eid=$wsT.Cells.Item($r,1).Text.Trim()
        if ($eid -match "product_template_(\d+)") { $n=[int]$Matches[1]; if ($n -gt $maxID){$maxID=$n;$lastDataRow=$r} }
    }
    $nextID=$maxID+1; $nextRow=$lastDataRow+1; $results=@()

    foreach ($p in $productList) {
        $pName=$p.Name.Trim(); $pAbbr=$p.BrandAbbr.Trim().ToUpper()
        $pPN=$p.PN.Trim().ToUpper(); $pDesc=$p.Description.Trim(); $pUnits=$p.Units.Trim()
        if (-not $brandMap.ContainsKey($pAbbr)) { Write-Warning "SKIP: brand $pAbbr not found"; continue }
        $bFull=$brandMap[$pAbbr]
        $desc=("BRAND: {0}`nP/N: {1}`nDESCRIPTION: {2}" -f $bFull,$pPN,$pDesc)
        $wsT.Cells.Item($nextRow,1).Value2=("product_template_{0}" -f $nextID)
        $wsT.Cells.Item($nextRow,2).Value2=$pName; $wsT.Cells.Item($nextRow,3).Value2="Goods"
        $wsT.Cells.Item($nextRow,4).Value2=1; $wsT.Cells.Item($nextRow,5).Value2=""
        $wsT.Cells.Item($nextRow,6).Value2=0; $wsT.Cells.Item($nextRow,7).Value2=$pUnits
        $wsT.Cells.Item($nextRow,8).Value2=0; $wsT.Cells.Item($nextRow,9).Value2=0
        $wsT.Cells.Item($nextRow,10).Value2=$desc; $wsT.Cells.Item($nextRow,11).Value2=$desc
        $wsT.Cells.Item($nextRow,12).Value2=$desc; $wsT.Cells.Item($nextRow,13).Value2=$desc
        $wsT.Cells.Item($nextRow,14).Value2="manassaporn.s"
        for ($c=10;$c -le 13;$c++){$wsT.Cells.Item($nextRow,$c).WrapText=$true}
        $results+=[PSCustomObject]@{ID=$nextID;Name=$pName;Brand=$bFull;PN=$pPN;Units=$pUnits}
        $nextID++; $nextRow++
    }

    Set-Theme $wsT; $wbT.Save(); $wbT.Close($false); $xl.Quit()
    Write-Output ("SUCCESS: {0} product(s)" -f $results.Count)
    foreach ($r in $results) { Write-Output ("  [{0}] {1}  Brand={2}  PN={3}" -f $r.ID,$r.Name,$r.Brand,$r.PN) }
    Write-Output ("FILE: {0}" -f $newName)
} catch {
    if ($null -ne $xl) { try { $xl.Quit() } catch {} }
    Write-Error ("ERROR: {0}" -f $_.Exception.Message); exit 1
} finally {
    if ($null -ne $xl) { try { [Runtime.InteropServices.Marshal]::ReleaseComObject($xl)|Out-Null } catch {} }
}
```

---

## วิธีใช้

**Single product:**
```powershell
# รัน script ด้านบนโดยใส่ค่า -Name, -BrandAbbr, -PN, -Description, -Units
.\fill.ps1 -Name "[JB]ENCA36N30BLP[HOF]" -BrandAbbr "HOF" -PN "ENCA36N30BLP" -Description "Steel enclosure" -Units "EA"
```

**Batch (หลาย product):**
```powershell
.\fill.ps1 -Products '[
  {"Name":"[JB]ENCA36N30BLP[HOF]","BrandAbbr":"HOF","PN":"ENCA36N30BLP","Description":"Steel enclosure","Units":"EA"},
  {"Name":"[CT]5SHY42L6500[ABB]","BrandAbbr":"ABB","PN":"5SHY42L6500","Description":"Current transformer","Units":"EA"}
]'
```

---

## Column Mapping

| Col | ชื่อ Column | ค่า |
|-----|------------|-----|
| A | id | product_template_[N] |
| B | name | [TYPE]PN[BRAND] |
| C | type | Goods |
| D | tracking | 1 |
| E | categ_id | (ว่าง) |
| F | active | 0 |
| G | uom_id | หน่วย (EA, m, pcs) |
| H | uom_po_id | 0 |
| I | pos_category_id | 0 |
| J-M | description_* | BRAND: X / P/N: Y / DESCRIPTION: Z |
| N | responsible_id | manassaporn.s |
