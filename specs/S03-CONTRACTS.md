# S03-CONTRACTS.md
**BACKWASH** | Date: 2026-01-23

## Library: simple_xlsx

### SIMPLE_XLSX Contracts

```eiffel
make
  ensure
    no_error: not has_error

open (a_path: STRING_32): detachable XLSX_WORKBOOK
  require
    path_not_empty: not a_path.is_empty

create_workbook: XLSX_WORKBOOK
  ensure
    result_not_void: Result /= Void
    one_sheet: Result.sheet_count = 1

save (a_workbook: XLSX_WORKBOOK; a_path: STRING_32): BOOLEAN
  require
    workbook_not_void: a_workbook /= Void
    path_not_empty: not a_path.is_empty
```

### XLSX_WORKBOOK Contracts

```eiffel
make
  ensure
    one_sheet: sheets.count = 1
    active_sheet_valid: active_sheet_index = 1

sheet (a_index: INTEGER): detachable XLSX_SHEET
  require
    index_valid: a_index >= 1 and a_index <= sheet_count

add_sheet (a_name: STRING_32): XLSX_SHEET
  require
    name_not_empty: not a_name.is_empty
    name_unique: not has_sheet_named (a_name)
  ensure
    sheet_added: sheets.count = old sheets.count + 1
    sheet_exists: has_sheet_named (a_name)

remove_sheet (a_index: INTEGER)
  require
    index_valid: a_index >= 1 and a_index <= sheet_count
    not_only_sheet: sheet_count > 1
  ensure
    sheet_removed: sheets.count = old sheets.count - 1
```

### XLSX_SHEET Contracts

```eiffel
make (a_name: STRING_32)
  require
    name_not_empty: not a_name.is_empty
  ensure
    name_set: name.same_string (a_name)
    rows_empty: rows.is_empty

cell (a_column, a_row: INTEGER): detachable XLSX_CELL
  require
    column_positive: a_column >= 1
    row_positive: a_row >= 1

put_cell (a_cell: XLSX_CELL)
  require
    cell_not_void: a_cell /= Void
  ensure
    cell_exists: attached cell (a_cell.column, a_cell.row)
```

### XLSX_CELL Contracts

```eiffel
make (a_column, a_row: INTEGER)
  require
    column_positive: a_column >= 1
    row_positive: a_row >= 1
  ensure
    column_set: column = a_column
    row_set: row = a_row
    is_empty: is_empty

set_string (a_value: STRING_32)
  require
    value_not_void: a_value /= Void
  ensure
    is_string: is_string
```

### Class Invariants

```eiffel
XLSX_WORKBOOK:
  sheets_exist: sheets /= Void
  at_least_one_sheet: sheets.count >= 1
  active_sheet_valid: active_sheet_index >= 1 and active_sheet_index <= sheets.count

XLSX_CELL:
  column_positive: column >= 1
  row_positive: row >= 1
  valid_cell_type: is_valid_type (cell_type)

XLSX_SHEET:
  name_not_empty: not name.is_empty
  rows_exist: rows /= Void
```
