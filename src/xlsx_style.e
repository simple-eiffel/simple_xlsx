note
	description: "[
		Cell formatting (font, color, border) for Excel spreadsheet cells.
	]"
	author: "Larry Rix"
	date: "$Date$"
	revision: "$Revision$"

class
	XLSX_STYLE

create
	make

feature {NONE} -- Initialization

	make
			-- Create default style.
		do
			create font_name.make_from_string ("Calibri")
			font_size := 11.0
			create font_color.make_from_string ("000000")
			create background_color.make_empty
			create number_format.make_empty
			create horizontal_alignment.make_from_string ("general")
			create vertical_alignment.make_from_string ("bottom")
		ensure
			not_bold: not is_bold
			not_italic: not is_italic
			not_underline: not is_underline
			default_font: font_name.same_string ({STRING_32} "Calibri")
			default_size: font_size = 11.0
		end

feature -- Font Attributes

	is_bold: BOOLEAN
			-- Is font bold?

	is_italic: BOOLEAN
			-- Is font italic?

	is_underline: BOOLEAN
			-- Is text underlined?

	font_name: STRING_32
			-- Font family name

	font_size: REAL_64
			-- Font size in points

	font_color: STRING_32
			-- Font color as hex (RRGGBB)

feature -- Cell Attributes

	background_color: STRING_32
			-- Background/fill color as hex (RRGGBB)

	number_format: STRING_32
			-- Number format string (e.g., "0.00", "yyyy-mm-dd")

	horizontal_alignment: STRING_32
			-- Horizontal alignment: general, left, center, right, fill, justify

	vertical_alignment: STRING_32
			-- Vertical alignment: top, center, bottom, justify

feature -- Font Setters

	set_bold (a_bold: BOOLEAN)
			-- Set bold state.
		do
			is_bold := a_bold
		ensure
			bold_set: is_bold = a_bold
		end

	set_italic (a_italic: BOOLEAN)
			-- Set italic state.
		do
			is_italic := a_italic
		ensure
			italic_set: is_italic = a_italic
		end

	set_underline (a_underline: BOOLEAN)
			-- Set underline state.
		do
			is_underline := a_underline
		ensure
			underline_set: is_underline = a_underline
		end

	set_font (a_name: STRING_32; a_size: REAL_64)
			-- Set font name and size.
		require
			name_not_empty: not a_name.is_empty
			size_positive: a_size > 0
		do
			font_name := a_name.twin
			font_size := a_size
		ensure
			name_set: font_name.same_string (a_name)
			size_set: font_size = a_size
		end

	set_font_color (a_color: STRING_32)
			-- Set font color (hex RRGGBB format).
		require
			valid_color: a_color.count = 6
		do
			font_color := a_color.twin
		ensure
			color_set: font_color.same_string (a_color)
		end

feature -- Cell Setters

	set_background_color (a_color: STRING_32)
			-- Set background color (hex RRGGBB format).
		require
			valid_color: a_color.is_empty or a_color.count = 6
		do
			background_color := a_color.twin
		ensure
			color_set: background_color.same_string (a_color)
		end

	set_number_format (a_format: STRING_32)
			-- Set number format string.
		do
			number_format := a_format.twin
		ensure
			format_set: number_format.same_string (a_format)
		end

	set_alignment (a_horizontal, a_vertical: STRING_32)
			-- Set cell alignment.
		require
			valid_horizontal: is_valid_horizontal (a_horizontal)
			valid_vertical: is_valid_vertical (a_vertical)
		do
			horizontal_alignment := a_horizontal.twin
			vertical_alignment := a_vertical.twin
		ensure
			horizontal_set: horizontal_alignment.same_string (a_horizontal)
			vertical_set: vertical_alignment.same_string (a_vertical)
		end

feature -- Query

	has_background: BOOLEAN
			-- Does this style have a background color?
		do
			Result := not background_color.is_empty
		end

	has_number_format: BOOLEAN
			-- Does this style have a number format?
		do
			Result := not number_format.is_empty
		end

feature -- Validation

	is_valid_horizontal (a_align: STRING_32): BOOLEAN
			-- Is `a_align' a valid horizontal alignment?
		do
			Result := a_align.same_string ({STRING_32} "general") or
			          a_align.same_string ({STRING_32} "left") or
			          a_align.same_string ({STRING_32} "center") or
			          a_align.same_string ({STRING_32} "right") or
			          a_align.same_string ({STRING_32} "fill") or
			          a_align.same_string ({STRING_32} "justify")
		end

	is_valid_vertical (a_align: STRING_32): BOOLEAN
			-- Is `a_align' a valid vertical alignment?
		do
			Result := a_align.same_string ({STRING_32} "top") or
			          a_align.same_string ({STRING_32} "center") or
			          a_align.same_string ({STRING_32} "bottom") or
			          a_align.same_string ({STRING_32} "justify")
		end

invariant
	font_name_not_empty: not font_name.is_empty
	font_size_positive: font_size > 0
	font_color_valid: font_color.count = 6
	horizontal_alignment_not_empty: not horizontal_alignment.is_empty
	vertical_alignment_not_empty: not vertical_alignment.is_empty

end
