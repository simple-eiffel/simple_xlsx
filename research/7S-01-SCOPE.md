# 7S-01-SCOPE.md

**Date**: 2026-01-23

**BACKWASH** | Date: 2026-01-23

## Library: simple_xlsx

### Problem Domain
Excel spreadsheet (.xlsx) file reading and writing for Eiffel applications.

### Core Use Cases
1. Read existing .xlsx files and extract cell data
2. Create new Excel workbooks programmatically
3. Write tabular data to spreadsheet format
4. Import/export CSV to/from XLSX format
5. Manipulate sheets, rows, cells, and styles

### Target Users
- Eiffel developers needing Excel file I/O
- Applications requiring report generation
- Data processing pipelines

### Boundaries
- **In Scope**: XLSX read/write, multi-sheet workbooks, cell types, basic styling, CSV conversion, merged cells, frozen panes
- **Out of Scope**: XLS (legacy format), charts, pivot tables, macros, formulas evaluation, password protection

### Success Criteria
- Can open, read, modify, and save .xlsx files
- Preserves cell types (string, number, boolean, date)
- Handles shared strings for memory efficiency
- Supports multiple sheets per workbook
