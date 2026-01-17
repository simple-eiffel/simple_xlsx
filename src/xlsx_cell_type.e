note
	description: "[
		Enumeration of cell value types for Excel spreadsheet cells.
	]"
	author: "Larry Rix"
	date: "$Date$"
	revision: "$Revision$"

class
	XLSX_CELL_TYPE

feature -- Constants

	string_type: INTEGER = 1
			-- String cell type

	number_type: INTEGER = 2
			-- Numeric cell type (INTEGER or REAL)

	boolean_type: INTEGER = 3
			-- Boolean cell type

	date_type: INTEGER = 4
			-- Date/DateTime cell type

	formula_type: INTEGER = 5
			-- Formula cell type (stores cached value)

	empty_type: INTEGER = 0
			-- Empty cell type

feature -- Query

	is_valid_type (a_type: INTEGER): BOOLEAN
			-- Is `a_type' a valid cell type?
		do
			Result := a_type >= empty_type and a_type <= formula_type
		ensure
			definition: Result = (a_type >= empty_type and a_type <= formula_type)
		end

	type_name (a_type: INTEGER): STRING_32
			-- Human-readable name for cell type.
		require
			valid_type: is_valid_type (a_type)
		do
			inspect a_type
			when 0 then Result := {STRING_32} "empty"
			when 1 then Result := {STRING_32} "string"
			when 2 then Result := {STRING_32} "number"
			when 3 then Result := {STRING_32} "boolean"
			when 4 then Result := {STRING_32} "date"
			when 5 then Result := {STRING_32} "formula"
			else
				Result := {STRING_32} "unknown"
			end
		end

end
