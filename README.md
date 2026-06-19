# alisa-product-tools

Claude Code plugin สำหรับทีม **Alisa Intersupply** — จัดการข้อมูลสินค้าครบวงจร ตั้งแต่รหัสย่อ Product Type, Brand Abbreviation จนถึงการเขียน Odoo import template

## Skills

| Skill | คำสั่ง | หน้าที่ |
|-------|--------|---------|
| `product-group` | `/alisa-product-tools:product-group` | หารหัสย่อ Product Type เช่น Junction box → **JB** |
| `brand-abbr` | `/alisa-product-tools:brand-abbr` | หา/สร้างรหัสย่อแบรนด์ 3 ตัวอักษร |
| `product-import` | `/alisa-product-tools:product-import` | นำเข้าหลายแบรนด์พร้อมกัน |
| `product-report` | `/alisa-product-tools:product-report` | สรุปภาพรวมสินค้าและแบรนด์ |
| `fill-odoo-template` | `/alisa-product-tools:fill-odoo-template` | เขียน Odoo product template จากรหัส `[TYPE]PN[BRAND]` |

## ติดตั้ง

```bash
npx skills add SaranNet/alisa-product-tools
```

## ตัวอย่างการใช้งาน

**หารหัสย่อ Product Type:**
> "Junction box ใช้รหัสอะไร?" → **JB**
> "Cable gland รหัสคืออะไร?" → **CG**

**สร้าง/ค้นหา Brand Abbreviation:**
> "รหัสย่อของ ABB คืออะไร?"
> "สร้างรหัสย่อสำหรับแบรนด์ PHOENIX LIGHTING"

**เขียน Odoo Template (batch):**
> เพิ่ม product เหล่านี้:
> - [JB]ENCA36N30BLP[HOF] — Steel enclosure — EA
> - [CT]5SHY42L6500[ABB] — Current transformer — EA
> - [CA]NYY-G 3x4+1x2.5[TRI] — Power cable — m

## รูปแบบรหัส Product

`[TYPE]PARTNUMBER[BRAND]` เช่น `[CT]123T4[MRA]`

| ส่วน | ตัวอย่าง | ความหมาย |
|------|---------|-----------|
| `[CT]` | TYPE Abbr | รหัสย่อ Product Type |
| `123T4` | Part Number | รหัสสินค้า |
| `[MRA]` | Brand Abbr | รหัสย่อแบรนด์ |

## ข้อกำหนดระบบ

- Windows 10/11
- Microsoft Excel
- Claude Code CLI

## ไฟล์ที่ใช้

```
Product code and Brand by Claude.xlsx  — ฐานข้อมูล Product Type + Brand (อ่าน/เขียน)
product_template to Odoo.xls           — Odoo import template (อ่าน/เขียน)
```
