<p align="center">
  <img src="https://raw.githubusercontent.com/simple-eiffel/.github/main/profile/assets/logo.png" alt="simple_ library logo" width="400">
</p>

# simple_xlsx

**[Documentation](https://simple-eiffel.github.io/simple_xlsx/)** | **[GitHub](https://github.com/simple-eiffel/simple_xlsx)**

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Eiffel](https://img.shields.io/badge/Eiffel-25.02-blue.svg)](https://www.eiffel.org/)
[![Design by Contract](https://img.shields.io/badge/DbC-enforced-orange.svg)]()
[![Built with simple_codegen](https://img.shields.io/badge/Built_with-simple__codegen-blueviolet.svg)](https://github.com/simple-eiffel/simple_code)

Read and write Excel (.xlsx) files using Office Open XML format.

Part of the [Simple Eiffel](https://github.com/simple-eiffel) ecosystem.

## Status

**Development** - Initial release with ZIP support via minizip-ng

## Overview

SIMPLE_XLSX provides Eiffel applications with the ability to read and write Microsoft Excel files in the modern .xlsx format (Office Open XML). It uses SIMPLE_XML for XML parsing/building and SIMPLE_ZIP from simple_archive for ZIP compression.

## Features

- **Read Excel Workbooks** - Open existing .xlsx files and access data
- **Write Excel Workbooks** - Create new spreadsheets programmatically
- **Multiple Sheets** - Support for workbooks with multiple worksheets
- **Cell Access** - Get/set cells by A1 notation or row/column index
- **Data Types** - Support for text, numbers, dates, and formulas
- **Design by Contract** - Full preconditions, postconditions, invariants
- **Void Safe** - Fully void-safe implementation
- **SCOOP Compatible** - Ready for concurrent use

## Installation

1. Set the ecosystem environment variable (one-time setup for all simple_* libraries):
```bash
export SIMPLE_EIFFEL=D:\prod
```

2. Add to your ECF:
```xml
<library name="simple_xlsx" location="$SIMPLE_EIFFEL/simple_xlsx/simple_xlsx.ecf"/>
```

## Quick Start

### Reading an Excel File

```eiffel
local
    xlsx: SIMPLE_XLSX
    workbook: XLSX_WORKBOOK
    sheet: XLSX_SHEET
do
    create xlsx.make

    -- Open workbook
    workbook := xlsx.open ("data.xlsx")

    -- Get first sheet
    sheet := workbook.sheet (1)

    -- Read cell value
    print (sheet.cell ("A1").as_string)
end
```

### Creating an Excel File

```eiffel
local
    xlsx: SIMPLE_XLSX
    workbook: XLSX_WORKBOOK
    sheet: XLSX_SHEET
do
    create xlsx.make

    -- Create new workbook
    workbook := xlsx.create_workbook

    -- Add a sheet
    sheet := workbook.add_sheet ("Sales Data")

    -- Add headers and data
    sheet.set_cell ("A1", "Product")
    sheet.set_cell ("B1", "Quantity")
    sheet.set_cell ("A2", "Widget")
    sheet.set_cell ("B2", 100)

    -- Save
    workbook.save ("output.xlsx")
end
```

## API Reference

### SIMPLE_XLSX

| Feature | Description |
|---------|-------------|
| `make` | Create instance |
| `open (path)` | Open existing Excel file |
| `create_workbook` | Create new workbook |

### XLSX_WORKBOOK

| Feature | Description |
|---------|-------------|
| `sheet (index)` | Get sheet by index (1-based) |
| `sheet_by_name (name)` | Get sheet by name |
| `add_sheet (name)` | Add new worksheet |
| `sheet_count` | Number of sheets |
| `save (path)` | Save workbook to file |

### XLSX_SHEET

| Feature | Description |
|---------|-------------|
| `cell (ref)` | Get cell by A1 reference |
| `cell_at (row, col)` | Get cell by row/column |
| `set_cell (ref, value)` | Set cell value |
| `row_count` | Number of rows with data |
| `column_count` | Number of columns with data |

### XLSX_CELL

| Feature | Description |
|---------|-------------|
| `as_string` | Get value as string |
| `as_integer` | Get value as integer |
| `as_real` | Get value as real |
| `is_empty` | Check if cell is empty |
| `cell_type` | Get cell data type |

## Dependencies

- simple_file
- simple_xml
- simple_archive (provides ZIP support via minizip-ng 4.0.10)

## License

MIT License - Copyright (c) 2024-2025, Larry Rix
