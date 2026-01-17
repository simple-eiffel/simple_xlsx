note
	description: "[
		ZIP archive wrapper for XLSX file operations.
		Delegates to SIMPLE_ZIP from simple_archive library.
	]"
	author: "Larry Rix"

class
	ZIP_ARCHIVE

create
	make

feature {NONE} -- Initialization

	make
			-- Create archive handler.
		do
			create zip.make
			create entries.make (10)
			create last_error.make_empty
		end

feature -- Access

	last_error: STRING_32
			-- Last error message

	entries: ARRAYED_LIST [TUPLE [name: STRING_8; content: STRING_8]]
			-- In-memory entries for writing

feature -- Status Report

	has_error: BOOLEAN
			-- Did last operation fail?
		do
			Result := not last_error.is_empty
		end

feature -- Read Operations

	open (a_path: STRING_8): BOOLEAN
			-- Open ZIP file for reading.
		require
			path_not_empty: not a_path.is_empty
		do
			last_error.wipe_out
			current_archive_path := a_path
			if zip.is_valid_archive (a_path) then
				is_open := True
				Result := True
			else
				last_error := {STRING_32} "Cannot open ZIP file: " + a_path
				Result := False
			end
		end

	has_entry (a_name: STRING_8): BOOLEAN
			-- Does archive contain entry named `a_name'?
		require
			name_not_empty: not a_name.is_empty
		do
			if attached current_archive_path as l_path then
				Result := zip.archive_contains (l_path, a_name)
			end
		end

	read_entry_as_string (a_name: STRING_8): detachable STRING_8
			-- Read entry content as string.
		require
			name_not_empty: not a_name.is_empty
		do
			if attached current_archive_path as l_path then
				Result := zip.extract_entry (l_path, a_name)
				if zip.has_error then
					last_error := zip.last_error.to_string_32
				end
			end
		end

	close
			-- Close archive.
		do
			is_open := False
			current_archive_path := Void
		end

feature -- Write Operations

	create_new (a_path: STRING_8): BOOLEAN
			-- Create new ZIP file for writing.
		require
			path_not_empty: not a_path.is_empty
		do
			last_error.wipe_out
			entries.wipe_out
			archive_path := a_path
			is_writing := True
			Result := True
		end

	add_entry_from_string (a_name: STRING_8; a_content: STRING_8)
			-- Add entry with string content.
		require
			name_not_empty: not a_name.is_empty
			content_not_void: a_content /= Void
			is_writing: is_writing
		do
			entries.extend ([a_name, a_content])
		end

	finalize: BOOLEAN
			-- Finalize and write ZIP file.
		do
			if is_writing and attached archive_path as l_path then
				zip.begin_create (l_path)
				if not zip.has_error then
					across entries as l_entry loop
						zip.add_entry_from_string (l_entry.name, l_entry.content)
					end
					zip.end_create
					Result := not zip.has_error
					if zip.has_error then
						last_error := zip.last_error.to_string_32
					end
				else
					last_error := zip.last_error.to_string_32
				end
				is_writing := False
			end
		end

feature -- Status

	is_writing: BOOLEAN
			-- Are we in write mode?

	is_open: BOOLEAN
			-- Are we in read mode?

feature {NONE} -- Implementation

	zip: SIMPLE_ZIP
			-- Underlying ZIP handler

	archive_path: detachable STRING_8
			-- Path for writing

	current_archive_path: detachable STRING_8
			-- Path for reading

invariant
	last_error_exists: last_error /= Void
	entries_exist: entries /= Void
	zip_exists: zip /= Void

end
