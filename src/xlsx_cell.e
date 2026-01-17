note
	description: "[
		Single cell with value and format in an Excel spreadsheet.

		Supports string, number, boolean, date, and formula cell types.
		Cell coordinates use 1-based indexing (Excel style).
	]"
	author: "Larry Rix"
	date: "$Date$"
	revision: "$Revision$"

class
	XLSX_CELL

inherit
	XLSX_CELL_TYPE
		redefine
			default_create
		end

create
	make,
	make_string,
	make_number,
	make_boolean,
	make_date,
	default_create

feature {NONE} -- Initialization

	default_create
			-- Create empty cell at A1.
		do
			column := 1
			row := 1
			cell_type := empty_type
			number_value := 0.0
		end

	make (a_column, a_row: INTEGER)
			-- Create empty cell at `a_column', `a_row'.
		require
			column_positive: a_column >= 1
			row_positive: a_row >= 1
		do
			column := a_column
			row := a_row
			cell_type := empty_type
			number_value := 0.0
		ensure
			column_set: column = a_column
			row_set: row = a_row
			is_empty: is_empty
		end

	make_string (a_column, a_row: INTEGER; a_value: STRING_32)
			-- Create string cell.
		require
			column_positive: a_column >= 1
			row_positive: a_row >= 1
			value_not_void: a_value /= Void
		do
			make (a_column, a_row)
			set_string (a_value)
		ensure
			column_set: column = a_column
			row_set: row = a_row
			is_string: is_string
			value_set: attached string_value as sv implies sv.same_string (a_value)
		end

	make_number (a_column, a_row: INTEGER; a_value: REAL_64)
			-- Create numeric cell.
		require
			column_positive: a_column >= 1
			row_positive: a_row >= 1
		do
			make (a_column, a_row)
			set_number (a_value)
		ensure
			column_set: column = a_column
			row_set: row = a_row
			is_number: is_number
			value_set: number_value = a_value
		end

	make_boolean (a_column, a_row: INTEGER; a_value: BOOLEAN)
			-- Create boolean cell.
		require
			column_positive: a_column >= 1
			row_positive: a_row >= 1
		do
			make (a_column, a_row)
			set_boolean (a_value)
		ensure
			column_set: column = a_column
			row_set: row = a_row
			is_boolean: is_boolean
			value_set: boolean_value = a_value
		end

	make_date (a_column, a_row: INTEGER; a_value: DATE_TIME)
			-- Create date cell.
		require
			column_positive: a_column >= 1
			row_positive: a_row >= 1
			value_not_void: a_value /= Void
		do
			make (a_column, a_row)
			set_date (a_value)
		ensure
			column_set: column = a_column
			row_set: row = a_row
			is_date: is_date
		end

feature -- Access

	column: INTEGER
			-- Column index (1-based)

	row: INTEGER
			-- Row index (1-based)

	cell_type: INTEGER
			-- Type of cell value

	string_value: detachable STRING_32
			-- String value (if string type)

	number_value: REAL_64
			-- Numeric value (if number type)

	boolean_value: BOOLEAN
			-- Boolean value (if boolean type)

	date_value: detachable DATE_TIME
			-- Date value (if date type)

	formula: detachable STRING_32
			-- Formula string (if formula type)

	style: detachable XLSX_STYLE
			-- Cell style/format

feature -- Status Report

	is_empty: BOOLEAN
			-- Is this an empty cell?
		do
			Result := cell_type = empty_type
		end

	is_string: BOOLEAN
			-- Is this a string cell?
		do
			Result := cell_type = string_type
		end

	is_number: BOOLEAN
			-- Is this a numeric cell?
		do
			Result := cell_type = number_type
		end

	is_boolean: BOOLEAN
			-- Is this a boolean cell?
		do
			Result := cell_type = boolean_type
		end

	is_date: BOOLEAN
			-- Is this a date cell?
		do
			Result := cell_type = date_type
		end

	is_formula: BOOLEAN
			-- Is this a formula cell?
		do
			Result := cell_type = formula_type
		end

feature -- Element Change

	set_string (a_value: STRING_32)
			-- Set string value.
		require
			value_not_void: a_value /= Void
		do
			cell_type := string_type
			string_value := a_value.twin
			date_value := Void
			formula := Void
		ensure
			is_string: is_string
			value_set: attached string_value as sv implies sv.same_string (a_value)
		end

	set_number (a_value: REAL_64)
			-- Set numeric value.
		do
			cell_type := number_type
			number_value := a_value
			string_value := Void
			date_value := Void
			formula := Void
		ensure
			is_number: is_number
			value_set: number_value = a_value
		end

	set_boolean (a_value: BOOLEAN)
			-- Set boolean value.
		do
			cell_type := boolean_type
			boolean_value := a_value
			string_value := Void
			date_value := Void
			formula := Void
		ensure
			is_boolean: is_boolean
			value_set: boolean_value = a_value
		end

	set_date (a_value: DATE_TIME)
			-- Set date value.
		require
			value_not_void: a_value /= Void
		do
			cell_type := date_type
			date_value := a_value.twin
			string_value := Void
			formula := Void
		ensure
			is_date: is_date
		end

	set_formula (a_formula, a_cached_value: STRING_32)
			-- Set formula with cached value.
		require
			formula_not_empty: not a_formula.is_empty
		do
			cell_type := formula_type
			formula := a_formula.twin
			string_value := a_cached_value.twin
			date_value := Void
		ensure
			is_formula: is_formula
			formula_set: attached formula as f implies f.same_string (a_formula)
		end

	clear
			-- Clear cell value (make empty).
		do
			cell_type := empty_type
			string_value := Void
			number_value := 0.0
			boolean_value := False
			date_value := Void
			formula := Void
		ensure
			is_empty: is_empty
		end

	set_style (a_style: XLSX_STYLE)
			-- Apply style to cell.
		require
			style_not_void: a_style /= Void
		do
			style := a_style
		ensure
			style_set: style = a_style
		end

feature -- Conversion

	cell_reference: STRING_32
			-- Excel-style cell reference (e.g., A1, B2, AA100).
		do
			Result := column_letter (column) + row.out
		end

	to_string: STRING_32
			-- String representation of cell value.
		do
			inspect cell_type
			when 0 then -- empty
				create Result.make_empty
			when 1 then -- string
				if attached string_value as sv then
					Result := sv.twin
				else
					create Result.make_empty
				end
			when 2 then -- number
				Result := number_value.out
			when 3 then -- boolean
				if boolean_value then
					Result := {STRING_32} "TRUE"
				else
					Result := {STRING_32} "FALSE"
				end
			when 4 then -- date
				if attached date_value as dv then
					Result := dv.out
				else
					create Result.make_empty
				end
			when 5 then -- formula
				-- Return cached value for formulas
				if attached string_value as sv then
					Result := sv.twin
				else
					Result := number_value.out
				end
			else
				create Result.make_empty
			end
		end

feature {NONE} -- Implementation

	column_letter (a_col: INTEGER): STRING_32
			-- Convert column number to Excel letter (1=A, 26=Z, 27=AA).
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

invariant
	column_positive: column >= 1
	row_positive: row >= 1
	valid_cell_type: is_valid_type (cell_type)

end
