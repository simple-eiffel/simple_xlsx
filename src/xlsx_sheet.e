note
	description: "[
		Single worksheet in an Excel workbook.

		Provides access to cells, rows, and sheet properties.
		All indices are 1-based (Excel style).
	]"
	author: "Larry Rix"
	date: "$Date$"
	revision: "$Revision$"

class
	XLSX_SHEET

create
	make

feature {NONE} -- Initialization

	make (a_name: STRING_32)
			-- Create sheet with `a_name'.
		require
			name_not_empty: not a_name.is_empty
		do
			name := a_name.twin
			create rows.make (100)
			create column_widths.make (26)
			create merged_ranges.make (0)
			frozen_column := 0
			frozen_row := 0
		ensure
			name_set: name.same_string (a_name)
			rows_empty: rows.is_empty
		end

feature -- Access

	name: STRING_32
			-- Sheet name

	rows: ARRAYED_LIST [XLSX_ROW]
			-- All rows in sheet (sparse - only non-empty rows)

	row_count: INTEGER
			-- Number of rows with data
		do
			Result := rows.count
		end

	column_count: INTEGER
			-- Maximum column used in any row
		do
			across rows as ic loop
				if ic.max_column > Result then
					Result := ic.max_column
				end
			end
		end

	row (a_row: INTEGER): detachable XLSX_ROW
			-- Get row by number (1-based), Void if not exists.
		require
			row_positive: a_row >= 1
		do
			across rows as ic loop
				if ic.row_number = a_row then
					Result := ic
				end
			end
		end

	cell (a_column, a_row: INTEGER): detachable XLSX_CELL
			-- Get cell at column and row.
		require
			column_positive: a_column >= 1
			row_positive: a_row >= 1
		do
			if attached row (a_row) as r then
				Result := r.cell (a_column)
			end
		end

	cell_by_reference (a_ref: STRING): detachable XLSX_CELL
			-- Get cell by reference (e.g., "A1", "B2", "AA100").
		require
			ref_not_empty: not a_ref.is_empty
		local
			l_col, l_row: INTEGER
		do
			parse_cell_reference (a_ref)
			l_col := last_parsed_column
			l_row := last_parsed_row
			if l_col >= 1 and l_row >= 1 then
				Result := cell (l_col, l_row)
			end
		end

	merged_ranges: ARRAYED_LIST [TUPLE [start_col, start_row, end_col, end_row: INTEGER]]
			-- List of merged cell ranges

feature -- Element Change

	set_name (a_name: STRING_32)
			-- Rename sheet.
		require
			name_not_empty: not a_name.is_empty
		do
			name := a_name.twin
		ensure
			name_set: name.same_string (a_name)
		end

	put_cell (a_cell: XLSX_CELL)
			-- Add or replace cell.
		require
			cell_not_void: a_cell /= Void
		local
			l_row: XLSX_ROW
		do
			if attached row (a_cell.row) as r then
				r.put_cell (a_cell)
			else
				create l_row.make (a_cell.row)
				l_row.put_cell (a_cell)
				rows.extend (l_row)
			end
		ensure
			cell_exists: attached cell (a_cell.column, a_cell.row)
		end

	put_string (a_column, a_row: INTEGER; a_value: STRING_32)
			-- Set string value at column, row.
		require
			column_positive: a_column >= 1
			row_positive: a_row >= 1
		local
			l_cell: XLSX_CELL
		do
			create l_cell.make_string (a_column, a_row, a_value)
			put_cell (l_cell)
		end

	put_number (a_column, a_row: INTEGER; a_value: REAL_64)
			-- Set number value at column, row.
		require
			column_positive: a_column >= 1
			row_positive: a_row >= 1
		local
			l_cell: XLSX_CELL
		do
			create l_cell.make_number (a_column, a_row, a_value)
			put_cell (l_cell)
		end

	put_boolean (a_column, a_row: INTEGER; a_value: BOOLEAN)
			-- Set boolean value at column, row.
		require
			column_positive: a_column >= 1
			row_positive: a_row >= 1
		local
			l_cell: XLSX_CELL
		do
			create l_cell.make_boolean (a_column, a_row, a_value)
			put_cell (l_cell)
		end

	put_date (a_column, a_row: INTEGER; a_value: DATE_TIME)
			-- Set date value at column, row.
		require
			column_positive: a_column >= 1
			row_positive: a_row >= 1
			value_not_void: a_value /= Void
		local
			l_cell: XLSX_CELL
		do
			create l_cell.make_date (a_column, a_row, a_value)
			put_cell (l_cell)
		end

feature -- Column Properties

	column_width (a_column: INTEGER): REAL_64
			-- Get column width (in character units), default 8.43.
		require
			column_positive: a_column >= 1
		do
			if column_widths.valid_index (a_column) then
				Result := column_widths.i_th (a_column)
			end
			if Result <= 0 then
				Result := 8.43 -- Default Excel column width
			end
		end

	set_column_width (a_column: INTEGER; a_width: REAL_64)
			-- Set column width in character units.
		require
			column_positive: a_column >= 1
			width_positive: a_width > 0
		do
			-- Extend array if needed
			from until column_widths.count >= a_column loop
				column_widths.extend (0.0)
			end
			column_widths.put_i_th (a_width, a_column)
		ensure
			width_set: column_width (a_column) = a_width
		end

feature -- Freeze Panes

	frozen_column: INTEGER
			-- Column up to which panes are frozen (0 = none)

	frozen_row: INTEGER
			-- Row up to which panes are frozen (0 = none)

	freeze_panes (a_column, a_row: INTEGER)
			-- Freeze rows and columns.
			-- Cells to the left of `a_column' and above `a_row' will be frozen.
		require
			column_non_negative: a_column >= 0
			row_non_negative: a_row >= 0
		do
			frozen_column := a_column
			frozen_row := a_row
		ensure
			column_frozen: frozen_column = a_column
			row_frozen: frozen_row = a_row
		end

feature -- Merged Cells

	merge_cells (a_start_col, a_start_row, a_end_col, a_end_row: INTEGER)
			-- Merge cell range.
		require
			start_col_valid: a_start_col >= 1
			start_row_valid: a_start_row >= 1
			end_col_valid: a_end_col >= a_start_col
			end_row_valid: a_end_row >= a_start_row
		do
			merged_ranges.extend ([a_start_col, a_start_row, a_end_col, a_end_row])
		ensure
			range_added: merged_ranges.count = old merged_ranges.count + 1
		end

feature -- Query

	max_row: INTEGER
			-- Highest row number with data, 0 if empty.
		do
			across rows as ic loop
				if ic.row_number > Result then
					Result := ic.row_number
				end
			end
		end

feature {NONE} -- Implementation

	column_widths: ARRAYED_LIST [REAL_64]
			-- Column widths indexed by column number

	last_parsed_column: INTEGER
			-- Column from last parse_cell_reference call

	last_parsed_row: INTEGER
			-- Row from last parse_cell_reference call

	parse_cell_reference (a_ref: STRING)
			-- Parse cell reference like "A1", "B2", "AA100".
			-- Sets last_parsed_column and last_parsed_row.
		local
			i: INTEGER
			c: CHARACTER_32
			l_letters, l_digits: STRING_32
		do
			create l_letters.make_empty
			create l_digits.make_empty
			last_parsed_column := 0
			last_parsed_row := 0

			from i := 1 until i > a_ref.count loop
				c := a_ref.item (i).as_upper
				if c >= 'A' and c <= 'Z' then
					l_letters.extend (c)
				elseif c >= '0' and c <= '9' then
					l_digits.extend (c)
				end
				i := i + 1
			end

			-- Convert letters to column number
			from i := 1 until i > l_letters.count loop
				last_parsed_column := last_parsed_column * 26 + (l_letters.item (i).code - ('A').code + 1)
				i := i + 1
			end

			-- Convert digits to row number
			if l_digits.is_integer then
				last_parsed_row := l_digits.to_integer
			end
		end

invariant
	name_not_empty: not name.is_empty
	rows_exist: rows /= Void
	merged_ranges_exist: merged_ranges /= Void

end
