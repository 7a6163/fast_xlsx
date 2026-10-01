# frozen_string_literal: true

require_relative "fast_xlsx/version"

# Fast .xlsx writer backed by rust_xlsxwriter.
#
# @example
#   wb = FastXlsx::Workbook.new
#   ws = wb.add_worksheet("Report")
#   ws << ["id", "name", "created_at"]
#   ws.concat(records.map { |r| [r.id, r.name, r.created_at] })
#   wb.save("report.xlsx")
#
# Rows and columns are 0-based. Errors: TypeError for an argument of the wrong
# type, RangeError for a row or column outside the sheet, ArgumentError for a
# value that isn't allowed, {FastXlsx::Error} for what the workbook can't do.
module FastXlsx
  # Raised for what the workbook can't do: duplicate or invalid names, writing
  # to rows already on disk, overlapping merges, text over Excel's limits, ...
  class Error < StandardError; end

  # A cell value written as an Excel formula.
  #
  # @!attribute [r] expression
  #   @return [String] the formula, without the leading "="
  # @example
  #   ws << [1, 2, FastXlsx::Formula.new("SUM(A1:B1)")]
  Formula = Data.define(:expression) do
    def initialize(expression:)
      super(expression: expression.to_s)
    end
  end

  # A cell value written as a hyperlink.
  #
  # @!attribute [r] url
  #   @return [String] http(s)://, mailto:, file:// or internal:Sheet2!A1
  # @!attribute [r] text
  #   @return [String, nil] shown in the cell instead of the URL
  # @example
  #   FastXlsx::URL.new("https://example.com")
  #   FastXlsx::URL.new("https://example.com/report/42", text: "Q3 report")
  # Subclassed (not a Data.define block) so .new can call super, which lets it
  # accept URL.new(url, text: ...) as well as the usual Data forms.
  class URL < Data.define(:url, :text) # rubocop:disable Style/DataInheritance
    # @overload new(url, text = nil)
    # @overload new(url:, text: nil)
    def self.new(*args, **kwargs)
      raise ArgumentError, "wrong number of arguments (given #{args.size}, expected 0..2)" if args.size > 2

      kwargs[:url] = args[0] unless args.empty?
      kwargs[:text] = args[1] if args.size > 1
      super(**kwargs)
    end

    # Also reached by URL[...] and #with, so values are normalized here.
    def initialize(url:, text: nil)
      super(url: url.to_s, text: text&.to_s)
    end
  end

  # Text with a format per segment.
  #
  # @example
  #   bold = FastXlsx::Format.new(bold: true)
  #   ws << [FastXlsx::RichString.new(["Total: ", bold], "1,234")]
  class RichString
    # @return [Array<Array(String, Format)>] [text, format or nil] per segment
    attr_reader :segments

    # @param parts [Array<String, Array(String, Format)>] each a String (default
    #   font) or [String, Format]
    # @raise [ArgumentError] when no segment is given
    def initialize(*parts)
      raise ArgumentError, "RichString needs at least one segment" if parts.empty?

      @segments = parts.map { |part| part.is_a?(Array) ? [part[0].to_s, part[1]] : [part.to_s, nil] }.freeze
    end
  end

  # The native extension looks up FastXlsx::Error, so load it after Error is defined.
  # Precompiled gems ship one binary per Ruby version (fast_xlsx/3.4/fast_xlsx);
  # a gem compiled on install has a single fast_xlsx/fast_xlsx.
  begin
    require "fast_xlsx/#{RUBY_VERSION[/\d+\.\d+/]}/fast_xlsx"
  rescue LoadError
    require "fast_xlsx/fast_xlsx"
  end

  # Owns the worksheets; serialize with {#to_xlsx} or {#save}.
  #
  # @!method to_xlsx
  #   The workbook as an .xlsx file. Runs without Ruby's global lock, so
  #   other threads keep running meanwhile.
  #   @return [String] binary
  #   @raise [FastXlsx::Error] e.g. duplicate defined names
  #
  # @!method define_name(name, formula)
  #   A defined name: "Rate" for the whole workbook, "Sheet1!Sales" for one
  #   sheet. Duplicates and unknown sheets are reported when saving.
  #   @param name [String]
  #   @param formula [String] e.g. "=0.96" or "=Sheet1!$A$1:$A$9"
  #   @return [self]
  #   @raise [FastXlsx::Error] for a name Excel doesn't allow
  class Workbook
    # Accepted memory: modes.
    MEMORY_MODES = %i[standard constant low].freeze

    # @param memory [Symbol] :standard keeps every cell in memory until saving.
    #   :constant and :low write each finished row to disk, so each worksheet
    #   must be filled top to bottom: :constant stores strings inline (memory
    #   stays flat), :low keeps Excel's shared string table (memory grows with
    #   the number of unique strings).
    # @raise [ArgumentError] for an unknown mode
    def self.new(memory: :standard)
      unless MEMORY_MODES.include?(memory)
        raise ArgumentError, "unknown memory mode #{memory.inspect} (expected one of #{MEMORY_MODES.join(", ")})"
      end

      _new(memory == :constant, memory == :low)
    end

    # @param name [String, nil] nil takes the first free "SheetN"
    # @return [Worksheet]
    # @raise [FastXlsx::Error] for an invalid name or one already used
    #   (names ignore case)
    def add_worksheet(name = nil)
      _add_worksheet(name).tap { |ws| worksheets << ws }
    end

    # The same Worksheet objects add_worksheet returned, so their append
    # position is shared.
    # @return [Array<Worksheet>]
    def worksheets
      @worksheets ||= []
    end

    # @param name [String]
    # @return [Worksheet, nil]
    def worksheet(name)
      worksheets.find { |ws| ws.name == name }
    end

    # Writes the .xlsx file. Runs without Ruby's global lock, like {#to_xlsx}.
    # @param path [String, Pathname, #to_path]
    # @return [nil]
    def save(path)
      _save(File.path(path))
    end

    # Document properties shown in Excel's File > Info. Later calls add to
    # earlier ones.
    # @param fields [Hash{Symbol => String}] title:, subject:, author:,
    #   manager:, company:, category:, keywords:, comments:, status:
    # @return [self]
    # @raise [ArgumentError] for an unknown field
    def properties(**fields)
      merged = (@properties || {}).merge(fields)
      _properties(merged) # validates before anything is remembered
      @properties = merged
      self
    end
  end

  # Cell writer for one sheet; create with {Workbook#add_worksheet}.
  #
  # A cell is (row, col) or a reference like "B2". A range is four 0-based
  # numbers (first_row, first_col, last_row, last_col), a reference like
  # "A1:D10", or rows and columns as Integers or Ranges, e.g. (0..9, 0..3).
  #
  # @!method write(*cell, value, format = nil)
  #   Writes one cell.
  #   @overload write(row, col, value, format = nil)
  #   @overload write(ref, value, format = nil)
  #   @param value [Numeric, String, Time, Date, DateTime, true, false, nil,
  #     Formula, URL, RichString, #to_s] dates get yyyy-mm-dd (hh:mm:ss) unless
  #     the format has a num_format
  #   @param format [Format, nil]
  #   @return [self]
  #
  # @!method <<(values)
  #   Appends a row after the last row written.
  #   @param values [Array] cell values, as for {#write}
  #   @return [self]
  #
  # @!method concat(rows)
  #   Appends several rows.
  #   @param rows [Array<Array>]
  #   @return [self]
  #
  # @!method next_row
  #   @return [Integer] the row {#<<} writes next
  #
  # @!method name
  #   @return [String]
  #
  # @!method row_height(row, height)
  #   @param height [Numeric] points, 0..409
  #   @return [self]
  #   @raise [ArgumentError] for a height outside 0..409
  #
  # @!method page_breaks(rows)
  #   Starts a printed page before each of these rows.
  #   @param rows [Array<Integer>]
  #   @return [self]
  #
  # @!method vertical_page_breaks(cols)
  #   Starts a printed page before each of these columns.
  #   @param cols [Array<Integer>]
  #   @return [self]
  #
  # @!method activate
  #   Makes Excel open on this sheet (un-hiding it if hidden).
  #   @return [self]
  #
  # @!method hide
  #   @return [self]
  #   @raise [FastXlsx::Error] for the sheet Excel opens on (the first one
  #     unless another is activated)
  #
  # @!method zoom(percent)
  #   @param percent [Integer, #to_int] 10..400
  #   @return [self]
  #
  # @!method tab_color(color)
  #   @param color [String, Integer] "#RRGGBB" or 0xRRGGBB
  #   @return [self]
  #
  # @!method default_row_height(height)
  #   Height of rows not given one with {#row_height}. Call it before
  #   {#row_height}, {#row_format}, {#hide_rows} and {#group_rows}: rows given
  #   those keep the earlier default.
  #   @param height [Numeric] points, above 0 and up to 409
  #   @return [self]
  #   @raise [FastXlsx::Error] when called after those
  #
  # @!method hide_gridlines
  #   Hides the gridlines on screen (see {#page_setup} for printing).
  #   @return [self]
  class Worksheet
    # write(row, col, value, format = nil) or write("B2", value, format = nil)
    # is native: the (row, col) form runs once per cell, so it skips a Ruby
    # wrapper. Other forms come here.
    def _write_ref(*args)
      row, col, (value, format) = CellRange.cell(args, 1..2)
      _write(row, col, value, format)
      self
    end
    private :_write_ref

    # Appends a row after the last row written.
    # @param values [Array] cell values, as for {#write}
    # @param format [Format, Array<Format, nil>, nil] one for every cell, or
    #   one per cell
    # @return [self]
    def append(values, format: nil)
      _append(values, format)
    end

    # @param columns [Integer, Range<Integer>]
    # @param width [Numeric] characters, 0..255
    # @return [self]
    def column_width(columns, width)
      range = CellRange.bounds(columns)
      _column_width(*range, width)
      @fixed_widths ||= {}
      @fixed_widths.delete(range) # re-insert so autofit replays calls in order
      @fixed_widths[range] = width
      self
    end

    # Sizes columns to the data written so far (in :constant / :low mode, the
    # rows still in memory). Widths set with {#column_width} are kept.
    # @return [self]
    def autofit
      _autofit
      @fixed_widths&.each { |bounds, width| _column_width(*bounds, width) }
      self
    end

    # Filter buttons on the range's first row.
    # @param range a cell range (see {Worksheet})
    # @return [self]
    def autofilter(*range)
      _autofilter(*CellRange.split(range).first)
    end

    # Merges the range and writes value (any cell type) into its first cell.
    # @overload merge_range(*range, value, format = nil)
    # @return [self]
    # @raise [FastXlsx::Error] when it overlaps an earlier merge
    def merge_range(*args)
      range, (value, format) = CellRange.split(args, 1..2)
      _merge_range(*range, value, format)
    end

    # Highlights cells in the range by rule; see the README for each type's
    # options.
    # @param range a cell range (see {Worksheet})
    # @param type [Symbol] :cell, :text, :formula, :data_bar or :color_scale
    # @option options [Symbol] :criteria e.g. :>, :between, :contains
    # @option options [Numeric, String, Array] :value
    # @option options [Format] :format
    # @option options [Integer] :colors 2 or 3 (:color_scale)
    # @return [self]
    def conditional_format(*range, type:, **)
      _conditional_format(*CellRange.split(range).first, { type: type, ** })
    end

    # Restricts what can be entered in the range; see the README.
    # @param range a cell range (see {Worksheet})
    # @param type [Symbol] :list, :whole_number, :decimal or :text_length
    # @option options [Symbol] :criteria e.g. :>, :between
    # @option options [Numeric, String, Array] :value
    # @option options [String] :input_title, :input_message, :error_title, :error_message
    # @return [self]
    def data_validation(*range, type:, **)
      _data_validation(*CellRange.split(range).first, { type: type, ** })
    end

    # Adds a comment (an Excel "note") to a cell.
    # @overload write_comment(*cell, text, author: nil)
    # @return [self]
    def write_comment(*args, author: nil)
      row, col, (text, *) = CellRange.cell(args, 1..1)
      _write_comment(row, col, text, author)
    end

    # Inserts a PNG, JPEG, GIF or BMP image with its top-left corner in the cell.
    # @overload insert_image(*cell, source, **options)
    #   @param source [String, #read] a file path or an IO
    #   @option options [Numeric] :scale
    #   @option options [Numeric] :width, :height pixels (the other keeps the ratio)
    #   @option options [Integer] :x_offset, :y_offset pixels
    #   @option options [String] :alt_text
    # @return [self]
    # @raise [FastXlsx::Error] for data that isn't a supported image
    def insert_image(*args, **)
      row, col, (source, *) = CellRange.cell(args, 1..1)
      bytes = source.respond_to?(:read) ? source.read : File.binread(source)
      _insert_image(row, col, bytes, { ** })
    end

    # Inserts a chart with its top-left corner in the cell.
    # @param cell (row, col) or "B2"
    # @param type [Symbol] :area, :bar, :column, :line, :pie, :doughnut,
    #   :radar, :scatter, and the stacked variants
    # @param series [Array<Hash>] { values:, categories:, name: } with ranges
    #   such as "Sheet1!$B$2:$B$13"
    # @option options [String] :title, :x_axis, :y_axis
    # @option options [Integer] :width, :height pixels
    # @return [self]
    def insert_chart(*cell, type:, series:, **)
      row, col, = CellRange.cell(cell)
      _insert_chart(row, col, { type: type, series: series, ** })
    end

    # Turns the range (header row included, total row too when total_row: true)
    # into an Excel table. Rows appended later under the header fill it.
    # @param range a cell range (see {Worksheet})
    # @option options [Array<String, Hash>] :columns header Strings or
    #   { header:, total:, total_label:, format: }
    # @option options [Symbol] :style e.g. :medium2
    # @option options [String] :name
    # @option options [Boolean] :total_row, :banded_rows, :autofilter
    # @return [self]
    def add_table(*range, **)
      _add_table(*CellRange.split(range).first, { ** })
    end

    # Printed page header, using Excel codes such as "&CPage &P of &N".
    # @param margin [Numeric, nil] inches
    # @return [self]
    def page_header(text, margin: nil)
      _page_header(text)
      margin ? margins(header: margin) : self
    end

    # Printed page footer; see {#page_header}.
    # @return [self]
    def page_footer(text, margin: nil)
      _page_footer(text)
      margin ? margins(footer: margin) : self
    end

    # Printing options.
    # @option options [Boolean] :landscape
    # @option options [Symbol, Integer] :paper :letter, :legal, :tabloid, :a3,
    #   :a4, :a5 or Excel's paper number (0: the printer's default)
    # @option options [Integer] :fit_width, :fit_height pages (0 or left out:
    #   as many as needed)
    # @param repeat_rows [Integer, Range, nil] printed on every page
    # @param repeat_columns [Integer, Range, nil] printed on every page
    # @param print_area a cell range: "A1:D100" or [first_row, first_col, last_row, last_col]
    # @option options [Boolean] :gridlines print the gridlines
    # @return [self]
    def page_setup(repeat_rows: nil, repeat_columns: nil, print_area: nil, **options)
      options[:repeat_rows] = CellRange.bounds(repeat_rows) if repeat_rows
      options[:repeat_columns] = CellRange.bounds(repeat_columns) if repeat_columns
      options[:print_area] = CellRange.split(print_area.is_a?(Array) ? print_area : [print_area]).first if print_area
      _page_setup(options)
      self
    end

    # Print margins in inches; margins not given keep their current value.
    # @return [self]
    # @raise [ArgumentError] for a negative or NaN margin
    def margins(left: nil, right: nil, top: nil, bottom: nil, header: nil, footer: nil)
      _margins(left, right, top, bottom, header, footer)
      self
    end

    # Outline group with an expand/collapse button. Grouping rows already
    # grouped nests them (up to 7 levels).
    # @param rows [Integer, Range<Integer>]
    # @param collapsed [Boolean] hidden until expanded
    # @return [self]
    # @raise [FastXlsx::Error] in :constant / :low memory mode, which can't
    #   write outline levels for rows
    def group_rows(rows, collapsed: false)
      _group_rows(*CellRange.bounds(rows), collapsed)
    end

    # Outline group of columns; see {#group_rows}. Works in every memory mode.
    # @param columns [Integer, Range<Integer>]
    # @return [self]
    def group_columns(columns, collapsed: false)
      _group_columns(*CellRange.bounds(columns), collapsed)
    end

    # Keeps the rows above and the columns left of the cell visible while
    # scrolling: (1, 0) or "A2" freezes the first row.
    # @return [self]
    def freeze_panes(*cell)
      _freeze_panes(*CellRange.cell(cell).first(2))
      self
    end

    # Locks the sheet against editing. Cells whose format has locked: false
    # stay editable. Calling it again replaces the password and actions.
    # @param password [String, nil] stops editing in Excel; it doesn't
    #   encrypt the file
    # @param allow [Symbol, Array<Symbol>] actions users may still take:
    #   :format_cells, :format_columns, :format_rows, :insert_columns,
    #   :insert_rows, :insert_links, :delete_columns, :delete_rows, :sort,
    #   :use_autofilter, :use_pivot_tables, :edit_scenarios, :edit_objects
    # @return [self]
    def protect(password: nil, allow: [])
      _protect(password, Array(allow))
    end

    # A range users can still edit on a protected sheet ({#protect}),
    # optionally with its own password. Give each range its own name.
    # @overload unprotect_range(*range, name: nil, password: nil)
    #   @param name [String, nil] shown in Excel's "Allow Edit Ranges"
    # @return [self]
    def unprotect_range(*range, name: nil, password: nil)
      _unprotect_range(*CellRange.split(range).first, name, password)
    end

    # The cells selected when the file opens.
    # @param range a cell range (see {Worksheet})
    # @return [self]
    def selection(*range)
      _selection(*CellRange.split(range).first)
    end

    # The cell scrolled to the top left when the file opens.
    # @param cell (row, col) or "B2"
    # @return [self]
    def top_left_cell(*cell)
      _top_left_cell(*CellRange.cell(cell).first(2))
    end

    # Turns off one of Excel's warnings (green triangles) in the range, e.g.
    # for codes stored as text. One kind per range (a rust_xlsxwriter limit);
    # overlapping ranges are accepted, but Excel then uses only one of them.
    # @overload ignore_error(*range, error)
    #   @param error [Symbol] :number_stored_as_text, :formula_error,
    #     :formula_differs, :formula_refers_to_empty_cells,
    #     :formula_omits_cells, :data_validation_error,
    #     :unlocked_cells_with_formula, :inconsistent_column_formula
    # @return [self]
    # @raise [FastXlsx::Error] when the range already has one
    def ignore_error(*args)
      range, (error, *) = CellRange.split(args, 1..1)
      _ignore_error(*range, error)
    end

    # Each row in the range is kept (about 1 KB) until saving.
    # @param rows [Integer, Range<Integer>]
    # @return [self]
    # @raise [FastXlsx::Error] in :constant / :low mode, for rows already on disk
    def hide_rows(rows)
      _hide_rows(*CellRange.bounds(rows))
    end

    # @param columns [Integer, Range<Integer>]
    # @return [self]
    def hide_columns(columns)
      _hide_columns(*CellRange.bounds(columns))
    end

    # Default format for cells in these rows that are written without one.
    # Use it or {#column_format} for a cell, not both: where both apply, Excel
    # may combine them or show the row's. Each row is kept (about 1 KB) until
    # saving, so style a whole sheet with {#column_format}.
    # @param rows [Integer, Range<Integer>]
    # @param format [Format]
    # @return [self]
    # @raise [FastXlsx::Error] in :constant / :low mode, for rows already on disk
    def row_format(rows, format)
      _row_format(*CellRange.bounds(rows), format)
      self
    end

    # Default format for cells in these columns that are written without one.
    # @param columns [Integer, Range<Integer>]
    # @param format [Format]
    # @return [self]
    def column_format(columns, format)
      _column_format(*CellRange.bounds(columns), format)
      self
    end
  end

  # Cell ranges in the styles Worksheet methods accept: four 0-based numbers,
  # an Excel reference ("A1:D10", "B2"), or rows and columns as Integers or
  # Ranges.
  # @api private
  module CellRange
    REF = /\A\$?([A-Za-z]{1,3})\$?([1-9]\d*)\z/ # ASCII only: /i also matches the Kelvin sign
    FORMS = 'a range is (first_row, first_col, last_row, last_col), "A1:D10" or (rows, cols)'
    CELL_FORMS = 'a cell is (row, col) or "B2"'

    module_function

    # Splits a range off the front of args and checks how many args follow.
    # Returns [[first_row, first_col, last_row, last_col], the args after it].
    def split(args, following = 0..0)
      range, rest = parse(args)
      check_following(rest, following, "cell range", FORMS)
      [range, rest]
    end

    def parse(args)
      case args
      in [Numeric, Numeric, Numeric, Numeric, *rest] then [args.first(4), rest]
      in [String => ref, *rest] then [excel(ref), rest]
      in [Integer | Range, Integer | Range, *rest] # (no "=> name" here: YARD can't parse it)
        [rows_and_cols(args[0], args[1]), rest]
      else
        raise ArgumentError, "expected a cell range, got #{args.inspect}; #{FORMS}"
      end
    end

    # (row, col) or a single-cell reference ("B2") off the front of args, as
    # [row, col, the args after it].
    def cell(args, following = 0..0)
      row, col, *rest = args.first.is_a?(String) ? [*single_cell(args.first), *args.drop(1)] : args
      raise ArgumentError, "expected a cell, got #{args.inspect}; #{CELL_FORMS}" if col.nil?

      check_following(rest, following, "cell", CELL_FORMS)
      [row, col, rest]
    end

    def single_cell(ref)
      first_row, first_col, last_row, last_col = excel(ref)
      return [first_row, first_col] if first_row == last_row && first_col == last_col

      raise ArgumentError, "expected a single cell like \"B2\", got #{ref.inspect}"
    end

    def check_following(rest, following, what, forms)
      return if following.cover?(rest.size)

      raise ArgumentError, "wrong number of arguments after the #{what} " \
                           "(given #{rest.size}, expected #{following.minmax.uniq.join("..")}); #{forms}"
    end

    def rows_and_cols(rows, cols)
      first_row, last_row = bounds(rows)
      first_col, last_col = bounds(cols)
      [first_row, first_col, last_row, last_col]
    end

    # "A1:D10", "$A$1:$D$10" or a single "B2".
    def excel(ref)
      cells = cell_matches(ref)
      rows = cells.map { |m| m[2].to_i - 1 }
      cols = cells.map { |m| column(m[1]) }
      [rows.min, cols.min, rows.max, cols.max]
    end

    def cell_matches(ref)
      cells = ref.split(":", -1).map { |cell| REF.match(cell) }
      return cells if (1..2).cover?(cells.size) && cells.all?

      raise ArgumentError, "invalid cell range #{ref.inspect}: use e.g. \"A1:D10\" or \"B2\""
    end

    # "A" => 0, "AA" => 26.
    def column(letters)
      letters.upcase.each_char.reduce(0) { |n, c| (n * 26) + c.ord - 64 } - 1
    end

    # [first, last] of an index or a Range.
    def bounds(indexes)
      return [indexes, indexes] if indexes.is_a?(Integer)
      unless indexes.is_a?(Range) && indexes.begin.is_a?(Integer) && indexes.end.is_a?(Integer)
        raise ArgumentError, "expected an Integer or a Range of Integers, got #{indexes.inspect}"
      end

      first, last = indexes.minmax
      raise ArgumentError, "empty range #{indexes.inspect}" unless first

      [first, last]
    end
  end
  private_constant :CellRange

  # A cell style; pass it to {Worksheet#write}, {Worksheet#append} and the
  # other methods that take a format. See the README for every option.
  #
  # @example
  #   header = FastXlsx::Format.new(bold: true, bg_color: "#DDEBF7", border_bottom: :thin)
  #   money  = FastXlsx::Format.new(num_format: "#,##0.00")
  class Format
    # @param options [Hash] bold:, italic:, underline:, strikeout:,
    #   font_script:, font_size:, font_name:, font_color:, bg_color:,
    #   num_format:, align:, valign:, text_wrap:, rotation:, indent:, shrink:,
    #   border:, border_left:, border_right:, border_top:, border_bottom:,
    #   border_color:, locked:, hidden:
    # @raise [ArgumentError] for an unknown option or an invalid value
    def self.new(**options)
      # Apply border: first so border_left: etc. override it whatever the order.
      options = { border: options[:border], **options.except(:border) } if options.key?(:border)
      _new(options)
    end
  end
end
