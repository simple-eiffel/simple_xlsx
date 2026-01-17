note
	description: "[
		Row of cells in an Excel worksheet.

		Row numbers use 1-based indexing (Excel style).
	]"
	author: "Larry Rix"
	date: "$Date$"
	revision: "$Revision$"

class
	XLSX_ROW

create
	make

feature {NONE} -- Initialization

	make (a_row_number: INTEGER)
			-- Create row at `a_row_number'.
		require
			row_positive: a_row_number >= 1
		do
			row_number := a_row_number
			create cells.make (10)
			height := 15.0 -- Default row height in points
		ensure
			row_number_set: row_number = a_row_number
			cells_empty: cells.is_empty
		end

feature -- Access

	row_number: INTEGER
			-- Row number (1-based)

	cells: ARRAYED_LIST [XLSX_CELL]
			-- All cells in row (sparse - only non-empty cells)

	cell_count: INTEGER
			-- Number of cells with values
		do
			Result := cells.count
		end

	cell (a_column: INTEGER): detachable XLSX_CELL
			-- Get cell at column (1-based), Void if not set.
		require
			column_positive: a_column >= 1
		do
			across cells as ic loop
				if ic.column = a_column then
					Result := ic
				end
			end
		end

	cell_by_letter (a_column: STRING): detachable XLSX_CELL
			-- Get cell by column letter (A, B, ..., AA, AB, ...).
		require
			column_not_empty: not a_column.is_empty
		local
			l_col_num: INTEGER
		do
			l_col_num := column_number (a_column)
			if l_col_num >= 1 then
				Result := cell (l_col_num)
			end
		end

	height: REAL_64
			-- Row height in points

	is_hidden: BOOLEAN
			-- Is row hidden?

feature -- Element Change

	put_cell (a_cell: XLSX_CELL)
			-- Add or replace cell at its column position.
		require
			cell_not_void: a_cell /= Void
			same_row: a_cell.row = row_number
		local
			l_found: BOOLEAN
			l_idx: INTEGER
		do
			-- Find existing cell at same column
			from
				l_idx := 1
			until
				l_idx > cells.count or l_found
			loop
				if cells.i_th (l_idx).column = a_cell.column then
					cells.put_i_th (a_cell, l_idx)
					l_found := True
				end
				l_idx := l_idx + 1
			end
			-- Add new cell if not found
			if not l_found then
				cells.extend (a_cell)
			end
		ensure
			cell_exists: attached cell (a_cell.column)
		end

	put_string (a_column: INTEGER; a_value: STRING_32)
			-- Set string value at column.
		require
			column_positive: a_column >= 1
		local
			l_cell: XLSX_CELL
		do
			create l_cell.make_string (a_column, row_number, a_value)
			put_cell (l_cell)
		ensure
			cell_set: attached cell (a_column) as c implies c.is_string
		end

	put_number (a_column: INTEGER; a_value: REAL_64)
			-- Set number at column.
		require
			column_positive: a_column >= 1
		local
			l_cell: XLSX_CELL
		do
			create l_cell.make_number (a_column, row_number, a_value)
			put_cell (l_cell)
		ensure
			cell_set: attached cell (a_column) as c implies c.is_number
		end

	put_boolean (a_column: INTEGER; a_value: BOOLEAN)
			-- Set boolean at column.
		require
			column_positive: a_column >= 1
		local
			l_cell: XLSX_CELL
		do
			create l_cell.make_boolean (a_column, row_number, a_value)
			put_cell (l_cell)
		ensure
			cell_set: attached cell (a_column) as c implies c.is_boolean
		end

	put_date (a_column: INTEGER; a_value: DATE_TIME)
			-- Set date at column.
		require
			column_positive: a_column >= 1
			value_not_void: a_value /= Void
		local
			l_cell: XLSX_CELL
		do
			create l_cell.make_date (a_column, row_number, a_value)
			put_cell (l_cell)
		ensure
			cell_set: attached cell (a_column) as c implies c.is_date
		end

	remove_cell (a_column: INTEGER)
			-- Remove cell at column.
		require
			column_positive: a_column >= 1
		local
			l_idx: INTEGER
		do
			from
				l_idx := cells.count
			until
				l_idx < 1
			loop
				if cells.i_th (l_idx).column = a_column then
					cells.go_i_th (l_idx)
					cells.remove
				end
				l_idx := l_idx - 1
			end
		ensure
			cell_removed: cell (a_column) = Void
		end

	set_height (a_height: REAL_64)
			-- Set row height in points.
		require
			height_positive: a_height > 0
		do
			height := a_height
		ensure
			height_set: height = a_height
		end

	set_hidden (a_hidden: BOOLEAN)
			-- Set row visibility.
		do
			is_hidden := a_hidden
		ensure
			hidden_set: is_hidden = a_hidden
		end

feature -- Query

	max_column: INTEGER
			-- Highest column number with a cell, 0 if empty.
		do
			across cells as ic loop
				if ic.column > Result then
					Result := ic.column
				end
			end
		end

feature {NONE} -- Implementation

	column_number (a_letter: STRING): INTEGER
			-- Convert column letter to number (A=1, Z=26, AA=27).
		local
			i: INTEGER
			c: CHARACTER_32
		do
			from
				i := 1
			until
				i > a_letter.count
			loop
				c := a_letter.item (i).as_upper
				if c >= 'A' and c <= 'Z' then
					Result := Result * 26 + (c.code - ('A').code + 1)
				end
				i := i + 1
			end
		end

invariant
	row_number_positive: row_number >= 1
	height_positive: height > 0
	cells_exist: cells /= Void

end
