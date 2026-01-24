# 7S-06-SIZING.md

**Date**: 2026-01-23

**BACKWASH** | Date: 2026-01-23

## Library: simple_xlsx

### Code Metrics

| Class | Lines | Features | Complexity |
|-------|-------|----------|------------|
| SIMPLE_XLSX | ~260 | 12 | Low (facade) |
| XLSX_WORKBOOK | ~220 | 18 | Medium |
| XLSX_SHEET | ~310 | 25 | Medium |
| XLSX_CELL | ~345 | 30 | Medium |
| XLSX_ROW | ~240 | 18 | Low |
| XLSX_READER | ~305 | 10 | High (parsing) |
| XLSX_WRITER | ~350 | 15 | High (generation) |
| XLSX_STYLE | ~195 | 20 | Low |
| XLSX_CELL_TYPE | ~60 | 8 | Low |
| ZIP_ARCHIVE | ~160 | 12 | Low (wrapper) |

### Total Estimated
- **Lines of Code**: ~2,445
- **Classes**: 10
- **Features**: ~168

### Memory Characteristics
- Shared strings: O(unique strings)
- Cell storage: O(non-empty cells) - sparse storage
- Row storage: O(rows with data) - sparse storage

### Performance Targets
- Read 10K cells: < 1 second
- Write 10K cells: < 2 seconds
- Memory: < 50MB for typical workbook
