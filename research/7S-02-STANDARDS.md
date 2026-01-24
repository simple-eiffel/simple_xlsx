# 7S-02-STANDARDS.md

**Date**: 2026-01-23

**BACKWASH** | Date: 2026-01-23

## Library: simple_xlsx

### Applicable Standards
1. **Office Open XML (OOXML)** - ECMA-376, ISO/IEC 29500
   - Defines the .xlsx file format structure
   - ZIP container with XML files inside

2. **ZIP Format** - PKWARE specification
   - Container format for XLSX files

### File Structure Compliance
```
[archive.xlsx]
  /[Content_Types].xml
  /_rels/.rels
  /xl/workbook.xml
  /xl/_rels/workbook.xml.rels
  /xl/worksheets/sheet1.xml, sheet2.xml, ...
  /xl/sharedStrings.xml
  /xl/styles.xml
```

### Required Content Types
- `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml`
- `application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml`
- `application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml`
- `application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml`

### Cell Types per OOXML
- `s` - Shared string index
- `b` - Boolean
- `n` - Number (default)
- `str` - Inline string
- `inlineStr` - Inline string (alternate)

### Date Handling
- Excel stores dates as serial numbers
- Epoch: 1899-12-30 (day 0)
- Fractional part represents time of day
