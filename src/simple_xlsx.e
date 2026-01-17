note
	description: "[
		SIMPLE_XLSX - Facade for Excel spreadsheet read/write.

		Quick methods for common operations:
		- Open existing .xlsx files
		- Create new workbooks
		- Read sheets as tables (list of lists)
		- Write tables to sheets
		- Import/export CSV
	]"
	author: "Larry Rix"
	date: "$Date$"
	revision: "$Revision$"

class
	SIMPLE_XLSX

create
	make

feature {NONE} -- Initialization

	make
			-- Create facade.
		do
			create reader.make
			create writer.make
			create last_error.make_empty
		ensure
			no_error: not has_error
		end

feature -- Access

	version: STRING = "0.1.0"
			-- Library version

	last_error: STRING_32
			-- Last error message

feature -- Status Report

	has_error: BOOLEAN
			-- Did last operation fail?
		do
			Result := not last_error.is_empty
		end

feature -- File Operations

	open (a_path: STRING_32): detachable XLSX_WORKBOOK
			-- Open existing .xlsx file.
		require
			path_not_empty: not a_path.is_empty
		do
			last_error.wipe_out
			Result := reader.read (a_path)
			if Result = Void then
				last_error := reader.last_error
			end
		end

	create_workbook: XLSX_WORKBOOK
			-- Create new empty workbook with one sheet.
		do
			last_error.wipe_out
			create Result.make
		ensure
			result_not_void: Result /= Void
			one_sheet: Result.sheet_count = 1
		end

	save (a_workbook: XLSX_WORKBOOK; a_path: STRING_32): BOOLEAN
			-- Save workbook to .xlsx file.
		require
			workbook_not_void: a_workbook /= Void
			path_not_empty: not a_path.is_empty
		do
			last_error.wipe_out
			Result := writer.write (a_workbook, a_path)
			if not Result then
				last_error := writer.last_error
			end
		end

feature -- Quick Read

	read_sheet_as_table (a_path: STRING_32; a_sheet_index: INTEGER): detachable ARRAYED_LIST [ARRAYED_LIST [STRING_32]]
			-- Read sheet as table of strings.
			-- Returns list of rows, each row is list of cell values.
		require
			path_not_empty: not a_path.is_empty
			index_positive: a_sheet_index >= 1
		local
			l_workbook: detachable XLSX_WORKBOOK
			l_sheet: detachable XLSX_SHEET
			l_row_data: ARRAYED_LIST [STRING_32]
			l_col: INTEGER
		do
			last_error.wipe_out
			l_workbook := open (a_path)
			if attached l_workbook as wb then
				if a_sheet_index <= wb.sheet_count then
					l_sheet := wb.sheet (a_sheet_index)
					if attached l_sheet as sh then
						create Result.make (sh.row_count)
						across sh.rows as row loop
							create l_row_data.make (row.max_column)
							from l_col := 1 until l_col > row.max_column loop
								if attached row.cell (l_col) as c then
									l_row_data.extend (c.to_string)
								else
									l_row_data.extend ({STRING_32} "")
								end
								l_col := l_col + 1
							end
							Result.extend (l_row_data)
						end
					end
				else
					last_error := {STRING_32} "Sheet index out of range"
				end
			end
		end

feature -- Quick Write

	write_table (a_table: ARRAYED_LIST [ARRAYED_LIST [STRING_32]]; a_path: STRING_32): BOOLEAN
			-- Write table of strings to new .xlsx file.
			-- Each inner list is a row.
		require
			table_not_void: a_table /= Void
			path_not_empty: not a_path.is_empty
		local
			l_workbook: XLSX_WORKBOOK
			l_sheet: detachable XLSX_SHEET
			l_row_num, l_col: INTEGER
		do
			last_error.wipe_out
			l_workbook := create_workbook
			l_sheet := l_workbook.active_sheet
			if attached l_sheet as sh then
				l_row_num := 1
				across a_table as row loop
					l_col := 1
					across row as cell loop
						sh.put_string (l_col, l_row_num, cell)
						l_col := l_col + 1
					end
					l_row_num := l_row_num + 1
				end
				Result := save (l_workbook, a_path)
			end
		end

feature -- CSV Operations

	import_csv (a_csv_path: STRING_32; a_xlsx_path: STRING_32): BOOLEAN
			-- Import CSV file to .xlsx.
		require
			csv_path_not_empty: not a_csv_path.is_empty
			xlsx_path_not_empty: not a_xlsx_path.is_empty
		local
			l_file: SIMPLE_FILE
			l_table: ARRAYED_LIST [ARRAYED_LIST [STRING_32]]
			l_row: ARRAYED_LIST [STRING_32]
			l_lines: LIST [STRING_32]
			l_cells: LIST [STRING_32]
		do
			last_error.wipe_out
			create l_file.make (a_csv_path)
			if l_file.exists then
				l_lines := l_file.read_text.split ('%N')
				create l_table.make (l_lines.count)
				across l_lines as line loop
					if not line.is_empty then
						l_cells := line.split (',')
						create l_row.make (l_cells.count)
						across l_cells as cell loop
							l_row.extend (cell.twin)
						end
						l_table.extend (l_row)
					end
				end
				Result := write_table (l_table, a_xlsx_path)
			else
				last_error := {STRING_32} "Cannot read CSV file: " + a_csv_path
			end
		end

	export_csv (a_xlsx_path: STRING_32; a_csv_path: STRING_32): BOOLEAN
			-- Export first sheet of .xlsx to CSV.
		require
			xlsx_path_not_empty: not a_xlsx_path.is_empty
			csv_path_not_empty: not a_csv_path.is_empty
		local
			l_table: detachable ARRAYED_LIST [ARRAYED_LIST [STRING_32]]
			l_file: SIMPLE_FILE
			l_content: STRING_32
			l_first_cell: BOOLEAN
		do
			last_error.wipe_out
			l_table := read_sheet_as_table (a_xlsx_path, 1)
			if attached l_table as tbl then
				create l_content.make (tbl.count * 100)
				across tbl as row loop
					l_first_cell := True
					across row as cell loop
						if not l_first_cell then
							l_content.append_character (',')
						end
						l_content.append (escape_csv (cell))
						l_first_cell := False
					end
					l_content.append_character ('%N')
				end
				create l_file.make (a_csv_path)
				Result := l_file.set_content (l_content)
				if not Result then
					last_error := {STRING_32} "Cannot write CSV file: " + a_csv_path
				end
			end
		end

feature {NONE} -- Implementation

	reader: XLSX_READER
			-- Reader instance

	writer: XLSX_WRITER
			-- Writer instance

	escape_csv (a_value: STRING_32): STRING_32
			-- Escape value for CSV output.
		do
			if a_value.has (',') or a_value.has ('"') or a_value.has ('%N') then
				create Result.make (a_value.count + 2)
				Result.append_character ('"')
				across a_value as c loop
					if c = '"' then
						Result.append ({STRING_32} "%"%"")
					else
						Result.extend (c)
					end
				end
				Result.append_character ('"')
			else
				Result := a_value.twin
			end
		end

invariant
	last_error_exists: last_error /= Void
	reader_exists: reader /= Void
	writer_exists: writer /= Void

end