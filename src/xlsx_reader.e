note
	description: "[
		Parse .xlsx files (ZIP containing XML).

		.xlsx is Office Open XML format:
		- [Content_Types].xml - content type definitions
		- xl/workbook.xml - workbook structure
		- xl/worksheets/sheet1.xml, sheet2.xml, ... - sheet data
		- xl/sharedStrings.xml - shared string table
		- xl/styles.xml - cell styles
	]"
	author: "Larry Rix"
	date: "$Date$"
	revision: "$Revision$"

class
	XLSX_READER

create
	make

feature {NONE} -- Initialization

	make
			-- Create reader.
		do
			create last_error.make_empty
			create shared_strings.make (100)
		ensure
			no_error: not has_error
		end

feature -- Access

	last_error: STRING_32
			-- Last error message

feature -- Status Report

	has_error: BOOLEAN
			-- Did last operation fail?
		do
			Result := not last_error.is_empty
		end

feature -- Query

	is_valid_xlsx (a_path: STRING_32): BOOLEAN
			-- Check if file is a valid .xlsx file.
		require
			path_not_empty: not a_path.is_empty
		local
			l_archive: ZIP_ARCHIVE
		do
			create l_archive.make
			if l_archive.open (a_path.to_string_8) then
				-- Check for required files
				Result := l_archive.has_entry ("[Content_Types].xml") and
				          l_archive.has_entry ("xl/workbook.xml")
				l_archive.close
			end
		end

feature -- Operation

	read (a_path: STRING_32): detachable XLSX_WORKBOOK
			-- Read workbook from .xlsx file.
		require
			path_not_empty: not a_path.is_empty
		local
			l_archive: ZIP_ARCHIVE
			l_workbook_xml, l_sheet_xml, l_strings_xml: detachable STRING_8
			l_sheet_paths: ARRAYED_LIST [STRING_8]
			l_sheet_names: ARRAYED_LIST [STRING_32]
			l_sheet: XLSX_SHEET
			i: INTEGER
		do
			last_error.wipe_out
			shared_strings.wipe_out

			create l_archive.make
			if not l_archive.open (a_path.to_string_8) then
				last_error := {STRING_32} "Cannot open file: " + a_path
			else
				-- Read shared strings first (needed for cell values)
				l_strings_xml := l_archive.read_entry_as_string ("xl/sharedStrings.xml")
				if attached l_strings_xml as ss then
					parse_shared_strings (ss)
				end

				-- Read workbook.xml to get sheet names
				l_workbook_xml := l_archive.read_entry_as_string ("xl/workbook.xml")
				if attached l_workbook_xml as wb then
					create l_sheet_names.make (3)
					create l_sheet_paths.make (3)
					parse_workbook_xml (wb, l_sheet_names, l_sheet_paths)

					-- Create workbook
					create Result.make

					-- Clear default sheet if we have sheets to load
					if not l_sheet_names.is_empty then
						Result.sheets.wipe_out
					end

					-- Read each sheet
					from i := 1 until i > l_sheet_names.count loop
						if l_sheet_paths.valid_index (i) then
							l_sheet_xml := l_archive.read_entry_as_string (l_sheet_paths.i_th (i))
							if attached l_sheet_xml as sx then
								create l_sheet.make (l_sheet_names.i_th (i))
								parse_sheet_xml (sx, l_sheet)
								Result.sheets.extend (l_sheet)
							end
						end
						i := i + 1
					end

					-- Ensure we have at least one sheet
					if Result.sheets.is_empty then
						Result.sheets.extend (create {XLSX_SHEET}.make ({STRING_32} "Sheet1"))
					end
					Result.set_active_sheet (1)
				else
					last_error := {STRING_32} "Cannot read workbook.xml"
				end

				l_archive.close
			end
		end

feature {NONE} -- Implementation

	shared_strings: ARRAYED_LIST [STRING_32]
			-- Shared string table

	parse_shared_strings (a_xml: STRING_8)
			-- Parse xl/sharedStrings.xml to populate shared_strings.
		local
			l_xml: SIMPLE_XML
			l_doc: SIMPLE_XML_DOCUMENT
			l_text: STRING_32
		do
			create l_xml.make
			l_doc := l_xml.parse (a_xml)
			if l_doc.is_valid then
				if attached l_doc.root as l_root then
					-- Find all <si> elements (string items)
					across l_root.elements ("si") as si loop
						-- Get text from <t> child element
						if attached si.element ("t") as t_node then
							l_text := t_node.text.to_string_32
						else
							create l_text.make_empty
						end
						shared_strings.extend (l_text)
					end
				end
			end
		end

	parse_workbook_xml (a_xml: STRING_8; a_names: ARRAYED_LIST [STRING_32]; a_paths: ARRAYED_LIST [STRING_8])
			-- Parse xl/workbook.xml to get sheet names and paths.
		local
			l_xml: SIMPLE_XML
			l_doc: SIMPLE_XML_DOCUMENT
			l_name: detachable STRING
			l_idx: INTEGER
		do
			create l_xml.make
			l_doc := l_xml.parse (a_xml)
			if l_doc.is_valid then
				if attached l_doc.root as l_root then
					-- Navigate to sheets element
					if attached l_root.element ("sheets") as l_sheets then
						across l_sheets.elements ("sheet") as sh loop
							l_name := sh.attr ("name")
							if attached l_name as nm and then not nm.is_empty then
								a_names.extend (nm.to_string_32)
								-- Sheet path is based on index (sheet1.xml, sheet2.xml, etc.)
								l_idx := a_names.count
								a_paths.extend ("xl/worksheets/sheet" + l_idx.out + ".xml")
							end
						end
					end
				end
			end
		end

	parse_sheet_xml (a_xml: STRING_8; a_sheet: XLSX_SHEET)
			-- Parse xl/worksheets/sheetN.xml to populate sheet data.
		local
			l_xml: SIMPLE_XML
			l_doc: SIMPLE_XML_DOCUMENT
			l_row: XLSX_ROW
			l_cell: XLSX_CELL
			l_ref, l_type, l_value: detachable STRING
			l_col, l_row_num: INTEGER
			l_string_index: INTEGER
		do
			create l_xml.make
			l_doc := l_xml.parse (a_xml)
			if l_doc.is_valid then
				if attached l_doc.root as l_root then
					-- Navigate to sheetData element
					if attached l_root.element ("sheetData") as l_sheet_data then
						across l_sheet_data.elements ("row") as row_node loop
							-- Get row number
							if attached row_node.attr ("r") as r_attr then
								l_row_num := r_attr.to_integer
							else
								l_row_num := 0
							end
							if l_row_num >= 1 then
								create l_row.make (l_row_num)

								-- Find all <c> (cell) elements in this row
								across row_node.elements ("c") as cell_node loop
									l_ref := cell_node.attr ("r")
									l_type := cell_node.attr ("t")

									-- Parse cell reference to get column
									if attached l_ref as ref then
										parse_cell_reference (ref.to_string_32)
										l_col := last_parsed_column
									else
										l_col := 0
									end

									if l_col >= 1 then
										create l_cell.make (l_col, l_row_num)

										-- Get value from <v> child element
										if attached cell_node.element ("v") as v_node then
											l_value := v_node.text
										else
											l_value := Void
										end

										-- Set cell value based on type
										if attached l_type as lt and then lt.same_string ("s") then
											-- Shared string
											if attached l_value as lv and then lv.is_integer then
												l_string_index := lv.to_integer
												if l_string_index >= 0 and l_string_index < shared_strings.count then
													l_cell.set_string (shared_strings.i_th (l_string_index + 1))
												end
											end
										elseif attached l_type as lt and then lt.same_string ("b") then
											-- Boolean
											l_cell.set_boolean (attached l_value as lv and then lv.same_string ("1"))
										elseif attached l_type as lt and then (lt.same_string ("str") or lt.same_string ("inlineStr")) then
											-- Inline string
											if attached l_value as lv then
												l_cell.set_string (lv.to_string_32)
											end
										else
											-- Number (default)
											if attached l_value as lv then
												if lv.is_real then
													l_cell.set_number (lv.to_real_64)
												elseif not lv.is_empty then
													l_cell.set_string (lv.to_string_32)
												end
											end
										end

										l_row.put_cell (l_cell)
									end
								end

								if l_row.cell_count > 0 then
									a_sheet.rows.extend (l_row)
								end
							end
						end
					end
				end
			end
		end

	last_parsed_column: INTEGER
			-- Column from last parse_cell_reference call

	parse_cell_reference (a_ref: STRING_32)
			-- Parse cell reference like "A1" to extract column number.
		local
			i: INTEGER
			c: CHARACTER_32
		do
			last_parsed_column := 0
			from i := 1 until i > a_ref.count loop
				c := a_ref.item (i).as_upper
				if c >= 'A' and c <= 'Z' then
					last_parsed_column := last_parsed_column * 26 + (c.code - ('A').code + 1)
				else
					i := a_ref.count -- Exit on first non-letter
				end
				i := i + 1
			end
		end

invariant
	last_error_exists: last_error /= Void
	shared_strings_exists: shared_strings /= Void

end
