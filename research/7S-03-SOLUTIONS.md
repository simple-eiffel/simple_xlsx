# 7S-03-SOLUTIONS.md

**Date**: 2026-01-23

**BACKWASH** | Date: 2026-01-23

## Library: simple_xlsx

### Alternative Solutions Evaluated

| Solution | Language | Pros | Cons |
|----------|----------|------|------|
| Apache POI | Java | Full-featured, mature | Java dependency |
| OpenXML SDK | C# | Microsoft official | .NET dependency |
| openpyxl | Python | Simple API | Python dependency |
| xlsx-js | JavaScript | Lightweight | Node.js dependency |
| libxlsxwriter | C | Fast, minimal | Manual memory management |

### Why Native Eiffel
1. **No FFI complexity** - Pure Eiffel + inline C for ZIP
2. **DBC integration** - Contracts for all operations
3. **Void safety** - Compile-time null checking
4. **Type safety** - Strong typing for cell values
5. **Ecosystem fit** - Integrates with simple_* libraries

### Architecture Decision
- Use simple_archive (simple_zip) for ZIP operations
- Use simple_xml for XML parsing/generation
- Layered design: Facade -> Domain Objects -> Reader/Writer

### Trade-offs Accepted
- No formula evaluation (read cached values only)
- Basic styling only (font, color, alignment)
- No macro support
- Limited to .xlsx (not .xls)
