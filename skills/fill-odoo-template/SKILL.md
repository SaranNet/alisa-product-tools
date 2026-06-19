---
name: fill-odoo-template
description: à¸™à¸³à¹€à¸‚à¹‰à¸²à¸‚à¹‰à¸­à¸¡à¸¹à¸¥ product à¸¥à¸‡à¹„à¸Ÿà¸¥à¹Œ product_template to Odoo.xls â€” à¹ƒà¸Šà¹‰à¸—à¸¸à¸à¸„à¸£à¸±à¹‰à¸‡à¸—à¸µà¹ˆà¸•à¹‰à¸­à¸‡à¸à¸²à¸£à¹€à¸žà¸´à¹ˆà¸¡ product à¹€à¸‚à¹‰à¸² Odoo template à¸ˆà¸²à¸à¸£à¸«à¸±à¸ª [TYPE]PN[BRAND] à¹€à¸Šà¹ˆà¸™ [CT]123T4[MRA] à¸žà¸£à¹‰à¸­à¸¡ description à¹à¸¥à¸°à¸«à¸™à¹ˆà¸§à¸¢
---

# Fill Odoo Product Template

à¸£à¸±à¸šà¸‚à¹‰à¸­à¸¡à¸¹à¸¥ product à¸ˆà¸²à¸à¸œà¸¹à¹‰à¹ƒà¸Šà¹‰ parse à¸£à¸«à¸±à¸ª [TYPE]PN[BRAND] à¹à¸¥à¹‰à¸§à¹€à¸‚à¸µà¸¢à¸™à¸¥à¸‡à¹„à¸Ÿà¸¥à¹Œ `product_template to Odoo.xls` à¸­à¸±à¸•à¹‚à¸™à¸¡à¸±à¸•à¸´

## Excel File Paths

```
Template : C:\Users\saran\OneDrive - Alisa intersupply CO.,LTD\Desktop\Product code\product_template to Odoo.xls
Brand DB : C:\Users\saran\OneDrive - Alisa intersupply CO.,LTD\Desktop\Product code\Product code and Brand by Claude.xlsx
```

## à¸£à¸¹à¸›à¹à¸šà¸šà¸£à¸«à¸±à¸ª Product

`[TYPE]PARTNUMBER[BRAND]` à¹€à¸Šà¹ˆà¸™ `[CT]123T4[MRA]`

| à¸ªà¹ˆà¸§à¸™ | à¸•à¸±à¸§à¸­à¸¢à¹ˆà¸²à¸‡ | à¸§à¸´à¸˜à¸µ parse |
|------|---------|-----------|
| `[CT]` | TYPE Abbr | à¸£à¸°à¸«à¸§à¹ˆà¸²à¸‡ `[` à¹à¸£à¸ à¹à¸¥à¸° `]` à¹à¸£à¸ |
| `123T4` | Part Number | à¸£à¸°à¸«à¸§à¹ˆà¸²à¸‡ `]` à¹à¸£à¸ à¹à¸¥à¸° `[` à¸ªà¸¸à¸”à¸—à¹‰à¸²à¸¢ |
| `[MRA]` | Brand Abbr | à¸£à¸°à¸«à¸§à¹ˆà¸²à¸‡ `[` à¸ªà¸¸à¸”à¸—à¹‰à¸²à¸¢ à¹à¸¥à¸° `]` à¸ªà¸¸à¸”à¸—à¹‰à¸²à¸¢ |

## à¸‚à¸±à¹‰à¸™à¸•à¸­à¸™

1. **Parse** à¸£à¸«à¸±à¸ª product à¸—à¸¸à¸à¸•à¸±à¸§à¸ˆà¸²à¸à¸œà¸¹à¹‰à¹ƒà¸Šà¹‰
2. **à¸–à¸²à¸¡à¸–à¹‰à¸²à¸‚à¸²à¸”** Units à¹à¸¥à¸° Description (à¸–à¹‰à¸²à¹„à¸¡à¹ˆà¹„à¸”à¹‰à¸£à¸±à¸šà¸¡à¸²)
3. **à¸£à¸±à¸™ script à¸„à¸£à¸±à¹‰à¸‡à¹€à¸”à¸µà¸¢à¸§** â€” à¸ªà¹ˆà¸‡à¸—à¸¸à¸ product à¹€à¸›à¹‡à¸™ JSON array (`-Products`)
4. **à¹à¸ªà¸”à¸‡à¸ªà¸£à¸¸à¸›** à¹€à¸›à¹‡à¸™à¸•à¸²à¸£à¸²à¸‡

## à¸£à¸±à¸™ Script

### à¸«à¸¥à¸²à¸¢ product (à¹à¸™à¸°à¸™à¸³ â€” Excel à¹€à¸›à¸´à¸”à¸„à¸£à¸±à¹‰à¸‡à¹€à¸”à¸µà¸¢à¸§)

```powershell
& "C:\Users\saran\.claude\plugins\cache\alisa-product-tools\alisa-product-tools\1.0.0\skills\fill-odoo-template\scripts\fill_template.ps1" `
    -Products '[{"Name":"[CT]123T4[MRA]","BrandAbbr":"MRA","PN":"123T4","Description":"Current Transformer 4VA","Units":"EA"},{"Name":"[JB]ENCA36N30BLP[HOF]","BrandAbbr":"HOF","PN":"ENCA36N30BLP","Description":"Steel enclosure junction box","Units":"EA"}]'
```

### product à¹€à¸”à¸µà¸¢à¸§

```powershell
& "C:\Users\saran\.claude\plugins\cache\alisa-product-tools\alisa-product-tools\1.0.0\skills\fill-odoo-template\scripts\fill_template.ps1" `
    -Name        "[CT]123T4[MRA]" `
    -BrandAbbr   "MRA" `
    -PN          "123T4" `
    -Description "Current Transformer 4VA" `
    -Units       "EA"
```

Script à¸ˆà¸°:
- à¹‚à¸«à¸¥à¸”à¸—à¸¸à¸à¹à¸šà¸£à¸™à¸”à¹Œà¸ˆà¸²à¸ `Brand abbr` sheet à¹€à¸‚à¹‰à¸² memory à¸„à¸£à¸±à¹‰à¸‡à¹€à¸”à¸µà¸¢à¸§
- à¸«à¸² External ID à¸–à¸±à¸”à¹„à¸› (`product_template_N`) à¸­à¸±à¸•à¹‚à¸™à¸¡à¸±à¸•à¸´
- à¹€à¸‚à¸µà¸¢à¸™à¸—à¸¸à¸ product row à¹ƒà¸™à¸„à¸£à¸±à¹‰à¸‡à¹€à¸”à¸µà¸¢à¸§
- Apply white/gray monochrome theme à¸«à¸¥à¸±à¸‡à¹€à¸‚à¸µà¸¢à¸™à¹€à¸ªà¸£à¹‡à¸ˆà¸—à¸±à¹‰à¸‡à¸«à¸¡à¸”
- Save à¸„à¸£à¸±à¹‰à¸‡à¹€à¸”à¸µà¸¢à¸§

## Column à¸—à¸µà¹ˆà¹€à¸•à¸´à¸¡à¸­à¸±à¸•à¹‚à¸™à¸¡à¸±à¸•à¸´

| Column | à¸„à¹ˆà¸² |
|--------|-----|
| External ID | `product_template_N` (auto-increment) |
| Product Type | `Goods` |
| Track Inventory | `1` |
| Barcode | (à¸§à¹ˆà¸²à¸‡) |
| Sales Price | `0` |
| Cost | `0` |
| Weight | `0` |
| Sales/Purchase/Inventory/Delivery Description | `BRAND: [à¸Šà¸·à¹ˆà¸­à¹€à¸•à¹‡à¸¡]\nP/N: [PN]\nDESCRIPTION: [DESC]` |
| Responsible | `manassaporn.s` |

## à¸ªà¸´à¹ˆà¸‡à¸—à¸µà¹ˆà¸•à¹‰à¸­à¸‡à¸£à¸±à¸šà¸ˆà¸²à¸à¸œà¸¹à¹‰à¹ƒà¸Šà¹‰

| Field | à¸•à¹‰à¸­à¸‡à¸à¸²à¸£ | à¸•à¸±à¸§à¸­à¸¢à¹ˆà¸²à¸‡ |
|-------|---------|---------|
| Product code | à¸šà¸±à¸‡à¸„à¸±à¸š | `[CT]123T4[MRA]` |
| Description | à¸šà¸±à¸‡à¸„à¸±à¸š | `Current Transformer 4VA` |
| Units | à¸šà¸±à¸‡à¸„à¸±à¸š | `EA`, `m`, `Units`, `pcs`, `set` |

> à¸–à¹‰à¸²à¸œà¸¹à¹‰à¹ƒà¸Šà¹‰à¹„à¸¡à¹ˆà¹„à¸”à¹‰à¸£à¸°à¸šà¸¸ Units à¸«à¸£à¸·à¸­ Description à¹ƒà¸«à¹‰à¸–à¸²à¸¡à¸à¹ˆà¸­à¸™à¸£à¸±à¸™ script

## à¸£à¸¹à¸›à¹à¸šà¸šà¸„à¸³à¸•à¸­à¸š

```
âœ… à¹€à¸žà¸´à¹ˆà¸¡ product à¸ªà¸³à¹€à¸£à¹‡à¸ˆ

| External ID | Name | Brand | P/N | Units |
|-------------|------|-------|-----|-------|
| product_template_2 | [CT]123T4[MRA] | MITSUBISHI | 123T4 | EA |
```

à¸«à¸²à¸à¸¡à¸µà¸«à¸¥à¸²à¸¢ product à¹ƒà¸«à¹‰à¹à¸ªà¸”à¸‡à¸—à¸¸à¸à¸£à¸²à¸¢à¸à¸²à¸£à¹ƒà¸™à¸•à¸²à¸£à¸²à¸‡à¹€à¸”à¸µà¸¢à¸§
