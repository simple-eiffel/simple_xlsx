# S05-CONSTRAINTS.md
**BACKWASH** | Date: 2026-01-23

## Library: simple_xlsx

### Technical Constraints

1. **File Format**
   - Only .xlsx (Office Open XML) supported
   - No .xls (BIFF) support
   - No .xlsm (macro-enabled) support

2. **Content Limits**
   - Max sheets: Limited by memory
   - Max rows: 1,048,576 (Excel limit)
   - Max columns: 16,384 (Excel limit)
   - Max string length: 32,767 characters per cell

3. **Cell Types**
   - String, Number, Boolean, Date, Formula
   - No rich text within cells
   - Formulas stored but not evaluated

4. **Styling**
   - Font: name, size, bold, italic, underline
   - Colors: hex RGB format
   - Alignment: horizontal, vertical
   - No borders (not implemented)
   - No conditional formatting

### Dependency Constraints

1. **simple_archive Required**
   - Must be compiled and available
   - SIMPLE_ZIP class used for ZIP operations

2. **simple_xml Required**
   - SIMPLE_XML for XML parsing
   - SIMPLE_XML_DOCUMENT for document manipulation

3. **EiffelStudio Version**
   - Requires EiffelStudio 22.05 or later
   - Void-safe mode required

### Platform Constraints

1. **Windows**
   - Tested on Windows 10/11
   - Path separators handled automatically

2. **File System**
   - UTF-8 file paths supported
   - No network path support tested

### Performance Constraints

1. **Memory**
   - Entire workbook loaded into memory
   - No streaming/lazy loading for large files
   - Sparse storage reduces memory for sparse data

2. **Processing**
   - Single-threaded read/write
   - No SCOOP parallel processing yet
