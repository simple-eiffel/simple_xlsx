# S06-BOUNDARIES.md
**BACKWASH** | Date: 2026-01-23

## Library: simple_xlsx

### API Boundaries

#### Public API (SIMPLE_XLSX)
- `make` - Constructor
- `open` - Open existing file
- `create_workbook` - New workbook
- `save` - Save workbook
- `read_sheet_as_table` - Quick read
- `write_table` - Quick write
- `import_csv` / `export_csv` - CSV conversion
- `has_error` / `last_error` - Error handling

#### Domain Objects (Public)
- `XLSX_WORKBOOK` - Full access
- `XLSX_SHEET` - Full access
- `XLSX_ROW` - Full access
- `XLSX_CELL` - Full access
- `XLSX_STYLE` - Full access
- `XLSX_CELL_TYPE` - Constants only

#### Internal Classes (Implementation)
- `XLSX_READER` - Used by facade/workbook
- `XLSX_WRITER` - Used by facade/workbook
- `ZIP_ARCHIVE` - Used by reader/writer

### Export Policies

```eiffel
XLSX_READER:
  feature {SIMPLE_XLSX, XLSX_WORKBOOK}
    read

XLSX_WRITER:
  feature {SIMPLE_XLSX, XLSX_WORKBOOK}
    write

ZIP_ARCHIVE:
  feature {XLSX_READER, XLSX_WRITER}
    -- all features

XLSX_CELL:
  feature {XLSX_SHEET, XLSX_ROW}
    -- internal manipulation
```

### Integration Points

| External System | Integration Method |
|-----------------|-------------------|
| simple_archive | ZIP_ARCHIVE wraps SIMPLE_ZIP |
| simple_xml | XLSX_READER uses for parsing |
| simple_file | SIMPLE_XLSX uses for CSV |
| File System | Standard Eiffel file I/O |

### Error Propagation

```
File Error -> ZIP_ARCHIVE.last_error
            -> XLSX_READER.last_error
            -> XLSX_WORKBOOK.last_error
            -> SIMPLE_XLSX.last_error
```

Each layer captures and propagates errors upward.
