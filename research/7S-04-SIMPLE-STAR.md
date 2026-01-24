# 7S-04-SIMPLE-STAR.md

**Date**: 2026-01-23

**BACKWASH** | Date: 2026-01-23

## Library: simple_xlsx

### Dependencies on simple_* Ecosystem

| Library | Purpose | Integration Point |
|---------|---------|-------------------|
| simple_archive | ZIP handling | ZIP_ARCHIVE wraps SIMPLE_ZIP |
| simple_xml | XML parsing | XLSX_READER uses SIMPLE_XML |
| simple_file | File operations | CSV import/export |

### Ecosystem Patterns Followed
1. **Facade Pattern** - SIMPLE_XLSX provides simplified API
2. **Quick Class** - One-liner operations for common tasks
3. **Error Handling** - `has_error` / `last_error` pattern
4. **Fluent API** - Method chaining where applicable

### Class Naming Convention
- `SIMPLE_XLSX` - Main facade
- `XLSX_*` - Domain classes (WORKBOOK, SHEET, ROW, CELL, etc.)
- No `SIMPLE_` prefix on internal classes

### Shared String Optimization
Like ecosystem pattern for shared resources:
- Collect unique strings during write
- Index-based lookup during read
- Reduces file size and memory usage
