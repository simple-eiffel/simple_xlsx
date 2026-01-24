# S07-SPEC-SUMMARY.md
**BACKWASH** | Date: 2026-01-23

## Library: simple_xlsx

### Executive Summary
simple_xlsx provides Eiffel applications with the ability to read and write Microsoft Excel .xlsx files. It implements the Office Open XML specification for spreadsheets, using the simple_* ecosystem for ZIP and XML operations.

### Key Capabilities
1. **Read XLSX** - Parse existing Excel files, extract cell data
2. **Write XLSX** - Create new Excel files programmatically
3. **Multi-Sheet** - Support multiple worksheets per workbook
4. **Cell Types** - String, number, boolean, date, formula (cached)
5. **Basic Styling** - Font, color, alignment
6. **CSV Conversion** - Import/export CSV files

### Architecture
```
SIMPLE_XLSX (Facade)
    |
    +-- XLSX_WORKBOOK
    |       +-- XLSX_SHEET[]
    |               +-- XLSX_ROW[]
    |                       +-- XLSX_CELL[]
    |
    +-- XLSX_READER ----> ZIP_ARCHIVE
    |                          |
    +-- XLSX_WRITER ----> SIMPLE_ZIP
```

### Class Count
- Total: 10 classes
- Facade: 1 (SIMPLE_XLSX)
- Domain: 6 (WORKBOOK, SHEET, ROW, CELL, STYLE, CELL_TYPE)
- Infrastructure: 3 (READER, WRITER, ZIP_ARCHIVE)

### Contract Coverage
- All public features have preconditions
- All modifying features have postconditions
- Class invariants on all domain classes
- Error state tracked via `has_error` / `last_error`

### Ecosystem Integration
- Depends on: simple_archive, simple_xml, simple_file
- No external non-Eiffel dependencies
- Pure Eiffel implementation with inline C for ZIP
