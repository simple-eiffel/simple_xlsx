# S04-FEATURE-SPECS.md
**BACKWASH** | Date: 2026-01-23

## Library: simple_xlsx

### SIMPLE_XLSX Features

| Feature | Signature | Description |
|---------|-----------|-------------|
| make | | Create facade |
| version | : STRING | Library version ("0.1.0") |
| open | (path: STRING_32): detachable XLSX_WORKBOOK | Open existing file |
| create_workbook | : XLSX_WORKBOOK | Create new workbook |
| save | (wb: XLSX_WORKBOOK; path: STRING_32): BOOLEAN | Save to file |
| read_sheet_as_table | (path: STRING_32; index: INTEGER): ARRAYED_LIST[ARRAYED_LIST[STRING_32]] | Quick read |
| write_table | (table: ARRAYED_LIST[...]; path: STRING_32): BOOLEAN | Quick write |
| import_csv | (csv_path, xlsx_path: STRING_32): BOOLEAN | CSV to XLSX |
| export_csv | (xlsx_path, csv_path: STRING_32): BOOLEAN | XLSX to CSV |
| has_error | : BOOLEAN | Error status |
| last_error | : STRING_32 | Error message |

### XLSX_WORKBOOK Features

| Feature | Signature | Description |
|---------|-----------|-------------|
| sheets | : ARRAYED_LIST[XLSX_SHEET] | All sheets |
| sheet_count | : INTEGER | Number of sheets |
| sheet | (index: INTEGER): detachable XLSX_SHEET | Get by index |
| sheet_by_name | (name: STRING_32): detachable XLSX_SHEET | Get by name |
| active_sheet | : detachable XLSX_SHEET | Currently active |
| add_sheet | (name: STRING_32): XLSX_SHEET | Add new sheet |
| remove_sheet | (index: INTEGER) | Remove sheet |
| save | (path: STRING_32): BOOLEAN | Save workbook |
| has_sheet_named | (name: STRING_32): BOOLEAN | Name exists? |

### XLSX_SHEET Features

| Feature | Signature | Description |
|---------|-----------|-------------|
| name | : STRING_32 | Sheet name |
| rows | : ARRAYED_LIST[XLSX_ROW] | All rows |
| row_count | : INTEGER | Number of rows |
| column_count | : INTEGER | Max column used |
| cell | (col, row: INTEGER): detachable XLSX_CELL | Get cell |
| cell_by_reference | (ref: STRING): detachable XLSX_CELL | Get by "A1" ref |
| put_string | (col, row: INTEGER; val: STRING_32) | Set string |
| put_number | (col, row: INTEGER; val: REAL_64) | Set number |
| put_boolean | (col, row: INTEGER; val: BOOLEAN) | Set boolean |
| put_date | (col, row: INTEGER; val: DATE_TIME) | Set date |
| freeze_panes | (col, row: INTEGER) | Freeze rows/cols |
| merge_cells | (start_col, start_row, end_col, end_row: INTEGER) | Merge range |

### XLSX_CELL Features

| Feature | Signature | Description |
|---------|-----------|-------------|
| column | : INTEGER | Column index (1-based) |
| row | : INTEGER | Row index (1-based) |
| cell_type | : INTEGER | Type constant |
| string_value | : detachable STRING_32 | String value |
| number_value | : REAL_64 | Number value |
| boolean_value | : BOOLEAN | Boolean value |
| date_value | : detachable DATE_TIME | Date value |
| formula | : detachable STRING_32 | Formula if any |
| is_empty, is_string, is_number, is_boolean, is_date, is_formula | : BOOLEAN | Type queries |
| set_string, set_number, set_boolean, set_date | | Value setters |
| cell_reference | : STRING_32 | "A1" format |
| to_string | : STRING_32 | String representation |
