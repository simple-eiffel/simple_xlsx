note
	description: "[
		Generate .xlsx files (ZIP containing XML).

		Creates Office Open XML format with required files:
		- [Content_Types].xml
		- _rels/.rels
		- xl/workbook.xml
		- xl/_rels/workbook.xml.rels
		- xl/worksheets/sheet1.xml, sheet2.xml, ...
		- xl/sharedStrings.xml
		- xl/styles.xml
	]"
	author: "Larry Rix"
	date: "$Date$"
	revision: "$Revision$"

class
	XLSX_WRITER

create
	make

feature {NONE} -- Initialization

	make
			-- Create writer.
		do
			create last_error.make_empty
			create shared_strings.make (100)
			create string_index_map.make (100)
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

feature -- Operation

	write (a_workbook: XLSX_WORKBOOK; a_path: STRING_32): BOOLEAN
			-- Write workbook to .xlsx file.
		require
			workbook_not_void: a_workbook /= Void
			path_not_empty: not a_path.is_empty
		local
			l_archive: ZIP_ARCHIVE
			i: INTEGER
		do
			last_error.wipe_out
			shared_strings.wipe_out
			string_index_map.wipe_out

			-- First pass: collect all shared strings
			collect_shared_strings (a_workbook)

			create l_archive.make
			if l_archive.create_new (a_path.to_string_8) then
				-- Add required files
				l_archive.add_entry_from_string ("[Content_Types].xml", generate_content_types (a_workbook))
				l_archive.add_entry_from_string ("_rels/.rels", generate_rels)
				l_archive.add_entry_from_string ("xl/workbook.xml", generate_workbook_xml (a_workbook))
				l_archive.add_entry_from_string ("xl/_rels/workbook.xml.rels", generate_workbook_rels (a_workbook))
				l_archive.add_entry_from_string ("xl/styles.xml", generate_styles_xml)

				-- Add shared strings if any
				if not shared_strings.is_empty then
					l_archive.add_entry_from_string ("xl/sharedStrings.xml", generate_shared_strings_xml)
				end

				-- Add each sheet
				from i := 1 until i > a_workbook.sheet_count loop
					if attached a_workbook.sheet (i) as l_sheet then
						l_archive.add_entry_from_string ("xl/worksheets/sheet" + i.out + ".xml", generate_sheet_xml (l_sheet))
					end
					i := i + 1
				end

				if l_archive.finalize then
					Result := True
				else
					last_error := l_archive.last_error
				end
			else
				last_error := {STRING_32} "Cannot create file: " + a_path
			end
		end

feature {NONE} -- Implementation

	shared_strings: ARRAYED_LIST [STRING_32]
			-- Shared string table

	string_index_map: HASH_TABLE [INTEGER, STRING_32]
			-- Map string value to index in shared_strings

	collect_shared_strings (a_workbook: XLSX_WORKBOOK)
			-- Collect all unique strings from workbook.
		local
			l_str: STRING_32
		do
			across a_workbook.sheets as sh loop
				across sh.rows as row loop
					across row.cells as cell loop
						if cell.is_string and attached cell.string_value as sv then
							l_str := sv
							if not string_index_map.has (l_str) then
								string_index_map.put (shared_strings.count, l_str)
								shared_strings.extend (l_str.twin)
							end
						end
					end
				end
			end
		end

	get_string_index (a_string: STRING_32): INTEGER
			-- Get index of string in shared strings table, or -1 if not found.
		do
			if string_index_map.has (a_string) then
				Result := string_index_map.item (a_string)
			else
				Result := -1
			end
		end

	generate_content_types (a_workbook: XLSX_WORKBOOK): STRING_8
			-- Generate [Content_Types].xml.
		local
			i: INTEGER
		do
			create Result.make (1000)
			Result.append ("<?xml version=%"1.0%" encoding=%"UTF-8%" standalone=%"yes%"?>%N")
			Result.append ("<Types xmlns=%"http://schemas.openxmlformats.org/package/2006/content-types%">%N")
			Result.append ("<Default Extension=%"rels%" ContentType=%"application/vnd.openxmlformats-package.relationships+xml%"/>%N")
			Result.append ("<Default Extension=%"xml%" ContentType=%"application/xml%"/>%N")
			Result.append ("<Override PartName=%"/xl/workbook.xml%" ContentType=%"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml%"/>%N")
			Result.append ("<Override PartName=%"/xl/styles.xml%" ContentType=%"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml%"/>%N")
			if not shared_strings.is_empty then
				Result.append ("<Override PartName=%"/xl/sharedStrings.xml%" ContentType=%"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml%"/>%N")
			end
			from i := 1 until i > a_workbook.sheet_count loop
				Result.append ("<Override PartName=%"/xl/worksheets/sheet" + i.out + ".xml%" ContentType=%"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml%"/>%N")
				i := i + 1
			end
			Result.append ("</Types>")
		end

	generate_rels: STRING_8
			-- Generate _rels/.rels.
		do
			create Result.make (500)
			Result.append ("<?xml version=%"1.0%" encoding=%"UTF-8%" standalone=%"yes%"?>%N")
			Result.append ("<Relationships xmlns=%"http://schemas.openxmlformats.org/package/2006/relationships%">%N")
			Result.append ("<Relationship Id=%"rId1%" Type=%"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument%" Target=%"xl/workbook.xml%"/>%N")
			Result.append ("</Relationships>")
		end

	generate_workbook_xml (a_workbook: XLSX_WORKBOOK): STRING_8
			-- Generate xl/workbook.xml.
		local
			i: INTEGER
		do
			create Result.make (1000)
			Result.append ("<?xml version=%"1.0%" encoding=%"UTF-8%" standalone=%"yes%"?>%N")
			Result.append ("<workbook xmlns=%"http://schemas.openxmlformats.org/spreadsheetml/2006/main%" ")
			Result.append ("xmlns:r=%"http://schemas.openxmlformats.org/officeDocument/2006/relationships%">%N")
			Result.append ("<sheets>%N")
			from i := 1 until i > a_workbook.sheet_count loop
				if attached a_workbook.sheet (i) as l_sheet then
					Result.append ("<sheet name=%"" + escape_xml (l_sheet.name).to_string_8 + "%" sheetId=%"" + i.out + "%" r:id=%"rId" + i.out + "%"/>%N")
				end
				i := i + 1
			end
			Result.append ("</sheets>%N")
			Result.append ("</workbook>")
		end

	generate_workbook_rels (a_workbook: XLSX_WORKBOOK): STRING_8
			-- Generate xl/_rels/workbook.xml.rels.
		local
			i, l_id: INTEGER
		do
			create Result.make (1000)
			Result.append ("<?xml version=%"1.0%" encoding=%"UTF-8%" standalone=%"yes%"?>%N")
			Result.append ("<Relationships xmlns=%"http://schemas.openxmlformats.org/package/2006/relationships%">%N")
			l_id := 1
			from i := 1 until i > a_workbook.sheet_count loop
				Result.append ("<Relationship Id=%"rId" + l_id.out + "%" Type=%"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet%" Target=%"worksheets/sheet" + i.out + ".xml%"/>%N")
				l_id := l_id + 1
				i := i + 1
			end
			Result.append ("<Relationship Id=%"rId" + l_id.out + "%" Type=%"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles%" Target=%"styles.xml%"/>%N")
			l_id := l_id + 1
			if not shared_strings.is_empty then
				Result.append ("<Relationship Id=%"rId" + l_id.out + "%" Type=%"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings%" Target=%"sharedStrings.xml%"/>%N")
			end
			Result.append ("</Relationships>")
		end

	generate_styles_xml: STRING_8
			-- Generate xl/styles.xml (minimal required styles).
		do
			create Result.make (500)
			Result.append ("<?xml version=%"1.0%" encoding=%"UTF-8%" standalone=%"yes%"?>%N")
			Result.append ("<styleSheet xmlns=%"http://schemas.openxmlformats.org/spreadsheetml/2006/main%">%N")
			Result.append ("<fonts count=%"1%"><font><sz val=%"11%"/><name val=%"Calibri%"/></font></fonts>%N")
			Result.append ("<fills count=%"2%"><fill><patternFill patternType=%"none%"/></fill><fill><patternFill patternType=%"gray125%"/></fill></fills>%N")
			Result.append ("<borders count=%"1%"><border/></borders>%N")
			Result.append ("<cellStyleXfs count=%"1%"><xf/></cellStyleXfs>%N")
			Result.append ("<cellXfs count=%"1%"><xf/></cellXfs>%N")
			Result.append ("</styleSheet>")
		end

	generate_shared_strings_xml: STRING_8
			-- Generate xl/sharedStrings.xml.
		do
			create Result.make (shared_strings.count * 50)
			Result.append ("<?xml version=%"1.0%" encoding=%"UTF-8%" standalone=%"yes%"?>%N")
			Result.append ("<sst xmlns=%"http://schemas.openxmlformats.org/spreadsheetml/2006/main%" count=%"" + shared_strings.count.out + "%" uniqueCount=%"" + shared_strings.count.out + "%">%N")
			across shared_strings as ss loop
				Result.append ("<si><t>" + escape_xml (ss).to_string_8 + "</t></si>%N")
			end
			Result.append ("</sst>")
		end

	generate_sheet_xml (a_sheet: XLSX_SHEET): STRING_8
			-- Generate xl/worksheets/sheetN.xml.
		local
			l_col_letter: STRING_32
		do
			create Result.make (a_sheet.row_count * 100)
			Result.append ("<?xml version=%"1.0%" encoding=%"UTF-8%" standalone=%"yes%"?>%N")
			Result.append ("<worksheet xmlns=%"http://schemas.openxmlformats.org/spreadsheetml/2006/main%">%N")
			Result.append ("<sheetData>%N")

			across a_sheet.rows as row loop
				Result.append ("<row r=%"" + row.row_number.out + "%">%N")
				across row.cells as cell loop
					l_col_letter := column_letter (cell.column)
					Result.append ("<c r=%"" + l_col_letter.to_string_8 + cell.row.out + "%"")

					if cell.is_string then
						-- Use shared string index
						if attached cell.string_value as sv then
							Result.append (" t=%"s%"><v>" + get_string_index (sv).out + "</v></c>%N")
						else
							Result.append ("><v></v></c>%N")
						end
					elseif cell.is_boolean then
						Result.append (" t=%"b%"><v>")
						if cell.boolean_value then
							Result.append ("1")
						else
							Result.append ("0")
						end
						Result.append ("</v></c>%N")
					elseif cell.is_number then
						Result.append ("><v>" + cell.number_value.out + "</v></c>%N")
					elseif cell.is_date then
						-- Excel stores dates as serial numbers
						if attached cell.date_value as dv then
							Result.append ("><v>" + date_to_serial (dv).out + "</v></c>%N")
						else
							Result.append ("/>%N")
						end
					else
						Result.append ("/>%N")
					end
				end
				Result.append ("</row>%N")
			end

			Result.append ("</sheetData>%N")

			-- Add merge cells if any
			if not a_sheet.merged_ranges.is_empty then
				Result.append ("<mergeCells count=%"" + a_sheet.merged_ranges.count.out + "%">%N")
				across a_sheet.merged_ranges as mr loop
					Result.append ("<mergeCell ref=%"" + column_letter (mr.start_col).to_string_8 + mr.start_row.out + ":" + column_letter (mr.end_col).to_string_8 + mr.end_row.out + "%"/>%N")
				end
				Result.append ("</mergeCells>%N")
			end

			Result.append ("</worksheet>")
		end

	column_letter (a_col: INTEGER): STRING_32
			-- Convert column number to letter (1=A, 26=Z, 27=AA).
		require
			positive: a_col >= 1
		local
			l_col, l_rem: INTEGER
		do
			create Result.make_empty
			l_col := a_col
			from until l_col = 0 loop
				l_rem := (l_col - 1) \\ 26
				Result.prepend_character ((('A').code + l_rem).to_character_32)
				l_col := (l_col - 1) // 26
			end
		end

	escape_xml (a_string: STRING_32): STRING_32
			-- Escape XML special characters.
		do
			create Result.make (a_string.count)
			across a_string as c loop
				inspect c
				when '<' then Result.append ({STRING_32} "&lt;")
				when '>' then Result.append ({STRING_32} "&gt;")
				when '&' then Result.append ({STRING_32} "&amp;")
				when '"' then Result.append ({STRING_32} "&quot;")
				when '%'' then Result.append ({STRING_32} "&apos;")
				else Result.extend (c)
				end
			end
		end

	date_to_serial (a_date: DATE_TIME): REAL_64
			-- Convert date to Excel serial number.
			-- Excel epoch is 1899-12-30 (day 0).
		local
			l_epoch: DATE_TIME
			l_duration: DATE_TIME_DURATION
		do
			create l_epoch.make (1899, 12, 30, 0, 0, 0)
			l_duration := a_date.relative_duration (l_epoch)
			Result := l_duration.day.to_double + (l_duration.hour * 3600 + l_duration.minute * 60 + l_duration.second).to_double / 86400.0
		end

invariant
	last_error_exists: last_error /= Void
	shared_strings_exists: shared_strings /= Void
	string_index_map_exists: string_index_map /= Void

end
