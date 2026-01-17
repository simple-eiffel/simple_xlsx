note
	description: "[
		Excel workbook containing multiple sheets.

		Supports loading from and saving to .xlsx files.
		Sheet indices are 1-based (Excel style).
	]"
	author: "Larry Rix"
	date: "$Date$"
	revision: "$Revision$"

class
	XLSX_WORKBOOK

create
	make,
	make_from_file

feature {NONE} -- Initialization

	make
			-- Create empty workbook with one default sheet.
		do
			create sheets.make (3)
			create last_error.make_empty
			active_sheet_index := 0
			-- Add default sheet
			sheets.extend (create {XLSX_SHEET}.make ({STRING_32} "Sheet1"))
			active_sheet_index := 1
		ensure
			one_sheet: sheets.count = 1
			active_sheet_valid: active_sheet_index = 1
		end

	make_from_file (a_path: STRING_32)
			-- Load workbook from .xlsx file.
		require
			path_not_empty: not a_path.is_empty
		local
			l_reader: XLSX_READER
		do
			create sheets.make (3)
			create last_error.make_empty
			active_sheet_index := 0
			file_path := a_path.twin

			create l_reader.make
			if attached l_reader.read (a_path) as l_loaded then
				-- Copy sheets from loaded workbook
				sheets := l_loaded.sheets
				active_sheet_index := l_loaded.active_sheet_index
				is_loaded := True
			else
				last_error := l_reader.last_error
				-- Create default sheet on failure
				sheets.extend (create {XLSX_SHEET}.make ({STRING_32} "Sheet1"))
				active_sheet_index := 1
			end
		ensure
			has_sheets: sheets.count >= 1
		end

feature -- Access

	sheets: ARRAYED_LIST [XLSX_SHEET]
			-- All sheets in workbook

	sheet_count: INTEGER
			-- Number of sheets
		do
			Result := sheets.count
		end

	sheet (a_index: INTEGER): detachable XLSX_SHEET
			-- Get sheet by index (1-based).
		require
			index_valid: a_index >= 1 and a_index <= sheet_count
		do
			if sheets.valid_index (a_index) then
				Result := sheets.i_th (a_index)
			end
		end

	sheet_by_name (a_name: STRING_32): detachable XLSX_SHEET
			-- Get sheet by name, Void if not found.
		require
			name_not_empty: not a_name.is_empty
		do
			across sheets as ic loop
				if ic.name.same_string (a_name) then
					Result := ic
				end
			end
		end

	active_sheet: detachable XLSX_SHEET
			-- Currently active sheet.
		do
			if active_sheet_index >= 1 and active_sheet_index <= sheets.count then
				Result := sheets.i_th (active_sheet_index)
			end
		end

	active_sheet_index: INTEGER
			-- Index of active sheet (1-based)

	file_path: detachable STRING_32
			-- Path of loaded file (Void if created new)

	is_loaded: BOOLEAN
			-- Was workbook loaded from file?

	last_error: STRING_32
			-- Last error message

feature -- Element Change

	add_sheet (a_name: STRING_32): XLSX_SHEET
			-- Add new sheet and return it.
		require
			name_not_empty: not a_name.is_empty
			name_unique: not has_sheet_named (a_name)
		do
			create Result.make (a_name)
			sheets.extend (Result)
		ensure
			sheet_added: sheets.count = old sheets.count + 1
			sheet_exists: has_sheet_named (a_name)
		end

	remove_sheet (a_index: INTEGER)
			-- Remove sheet by index.
		require
			index_valid: a_index >= 1 and a_index <= sheet_count
			not_only_sheet: sheet_count > 1
		do
			sheets.go_i_th (a_index)
			sheets.remove
			-- Adjust active sheet index if needed
			if active_sheet_index > sheets.count then
				active_sheet_index := sheets.count
			end
		ensure
			sheet_removed: sheets.count = old sheets.count - 1
			active_valid: active_sheet_index >= 1 and active_sheet_index <= sheets.count
		end

	remove_sheet_by_name (a_name: STRING_32)
			-- Remove sheet by name.
		require
			name_not_empty: not a_name.is_empty
			sheet_exists: has_sheet_named (a_name)
			not_only_sheet: sheet_count > 1
		local
			i: INTEGER
		do
			from i := 1 until i > sheets.count loop
				if sheets.i_th (i).name.same_string (a_name) then
					remove_sheet (i)
					i := sheets.count + 1 -- Exit loop
				else
					i := i + 1
				end
			end
		ensure
			sheet_removed: not has_sheet_named (a_name)
		end

	set_active_sheet (a_index: INTEGER)
			-- Set active sheet by index.
		require
			index_valid: a_index >= 1 and a_index <= sheet_count
		do
			active_sheet_index := a_index
		ensure
			active_set: active_sheet_index = a_index
		end

feature -- Persistence

	save (a_path: STRING_32): BOOLEAN
			-- Save workbook to file. Returns True on success.
		require
			path_not_empty: not a_path.is_empty
		local
			l_writer: XLSX_WRITER
		do
			create l_writer.make
			Result := l_writer.write (Current, a_path)
			if Result then
				file_path := a_path.twin
				last_error.wipe_out
			else
				last_error := l_writer.last_error
			end
		end

feature -- Query

	has_sheet_named (a_name: STRING_32): BOOLEAN
			-- Does a sheet with `a_name' exist?
		require
			name_not_empty: not a_name.is_empty
		do
			Result := sheet_by_name (a_name) /= Void
		end

	has_error: BOOLEAN
			-- Did last operation fail?
		do
			Result := not last_error.is_empty
		end

invariant
	sheets_exist: sheets /= Void
	at_least_one_sheet: sheets.count >= 1
	active_sheet_valid: active_sheet_index >= 1 and active_sheet_index <= sheets.count
	last_error_exists: last_error /= Void

end
