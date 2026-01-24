# 7S-05-SECURITY.md

**Date**: 2026-01-23

**BACKWASH** | Date: 2026-01-23

## Library: simple_xlsx

### Security Considerations

#### Input Validation
1. **File Path Validation**
   - Check path not empty
   - Validate file exists before reading
   - Handle permission errors gracefully

2. **XML Parsing**
   - Delegate to simple_xml (uses Gobo XML)
   - No external entity expansion (XXE prevention)
   - Size limits on shared strings table

3. **ZIP Extraction**
   - Validate archive structure
   - Check for zip bombs (excessive compression ratio)
   - Limit extracted file sizes

#### Data Integrity
1. **Cell Type Preservation**
   - Maintain type through read/write cycle
   - No silent type coercion

2. **Unicode Handling**
   - Full UTF-8 support via STRING_32
   - Proper XML escaping on output

#### Attack Vectors Mitigated
- **Path Traversal**: No user-controlled paths in archive
- **XML Bombs**: Gobo parser limits
- **Memory Exhaustion**: Lazy loading of large sheets

### Not Addressed
- Password-protected workbooks (not supported)
- Macro security (macros not supported)
- Digital signatures (not implemented)
