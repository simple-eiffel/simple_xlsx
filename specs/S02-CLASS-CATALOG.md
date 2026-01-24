# S02-CLASS-CATALOG.md
**BACKWASH** | Date: 2026-01-23

## Library: simple_xlsx

### Class Hierarchy

```
ANY
  SIMPLE_XLSX           -- Main facade
  XLSX_WORKBOOK         -- Workbook container
  XLSX_SHEET            -- Single worksheet
  XLSX_ROW              -- Row of cells
  XLSX_CELL_TYPE        -- Cell type constants
    XLSX_CELL           -- Individual cell (inherits XLSX_CELL_TYPE)
  XLSX_STYLE            -- Cell formatting
  XLSX_READER           -- XLSX file parser
  XLSX_WRITER           -- XLSX file generator
  ZIP_ARCHIVE           -- ZIP wrapper
```

### Class Descriptions

| Class | Responsibility | Key Collaborators |
|-------|----------------|-------------------|
| SIMPLE_XLSX | High-level API facade | XLSX_READER, XLSX_WRITER |
| XLSX_WORKBOOK | Contains sheets, manages persistence | XLSX_SHEET, XLSX_READER, XLSX_WRITER |
| XLSX_SHEET | Contains rows/cells, sheet operations | XLSX_ROW, XLSX_CELL |
| XLSX_ROW | Sparse cell collection for one row | XLSX_CELL |
| XLSX_CELL | Single cell with value and type | XLSX_CELL_TYPE, XLSX_STYLE |
| XLSX_CELL_TYPE | Cell type enumeration (string, number, etc.) | - |
| XLSX_STYLE | Font, color, alignment properties | - |
| XLSX_READER | Parse .xlsx to workbook | ZIP_ARCHIVE, SIMPLE_XML |
| XLSX_WRITER | Generate .xlsx from workbook | ZIP_ARCHIVE |
| ZIP_ARCHIVE | ZIP read/write operations | SIMPLE_ZIP |

### Creation Procedures

| Class | Creators |
|-------|----------|
| SIMPLE_XLSX | make |
| XLSX_WORKBOOK | make, make_from_file |
| XLSX_SHEET | make |
| XLSX_ROW | make |
| XLSX_CELL | make, make_string, make_number, make_boolean, make_date, default_create |
| XLSX_STYLE | make |
| XLSX_READER | make |
| XLSX_WRITER | make |
| ZIP_ARCHIVE | make |
