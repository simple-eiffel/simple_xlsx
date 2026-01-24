# S01-PROJECT-INVENTORY.md
**BACKWASH** | Date: 2026-01-23

## Library: simple_xlsx

### Source Files

| File | Path | Purpose |
|------|------|---------|
| simple_xlsx.e | src/ | Main facade class |
| xlsx_workbook.e | src/ | Workbook container |
| xlsx_sheet.e | src/ | Single worksheet |
| xlsx_row.e | src/ | Row of cells |
| xlsx_cell.e | src/ | Individual cell |
| xlsx_cell_type.e | src/ | Cell type constants |
| xlsx_style.e | src/ | Cell formatting |
| xlsx_reader.e | src/ | XLSX file parser |
| xlsx_writer.e | src/ | XLSX file generator |
| zip_archive.e | src/ | ZIP wrapper |

### Test Files

| File | Path | Purpose |
|------|------|---------|
| test_app.e | testing/ | Test application entry |
| lib_tests.e | testing/ | Test suite |

### Configuration Files

| File | Purpose |
|------|---------|
| simple_xlsx.ecf | Library ECF |
| simple_xlsx_tests.ecf | Test target ECF |

### Dependencies
- simple_archive (SIMPLE_ZIP)
- simple_xml (SIMPLE_XML, SIMPLE_XML_DOCUMENT)
- simple_file (SIMPLE_FILE)
- EiffelBase
- Gobo XML

### Directory Structure
```
simple_xlsx/
  src/
    simple_xlsx.e
    xlsx_workbook.e
    xlsx_sheet.e
    xlsx_row.e
    xlsx_cell.e
    xlsx_cell_type.e
    xlsx_style.e
    xlsx_reader.e
    xlsx_writer.e
    zip_archive.e
  testing/
    test_app.e
    lib_tests.e
  research/
  specs/
```
