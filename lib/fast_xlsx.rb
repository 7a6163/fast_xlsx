# frozen_string_literal: true

require_relative "fast_xlsx/version"

# Fast .xlsx writer backed by rust_xlsxwriter.
module FastXlsx
  class Error < StandardError; end

  # Cell value written as an Excel formula, e.g. Formula.new("SUM(A1:A9)").
  Formula = Data.define(:expression) do
    def initialize(expression:)
      super(expression: expression.to_s)
    end
  end

  # Cell value written as a hyperlink, e.g. URL.new("https://example.com").
  # text: shown in the cell instead of the URL itself.
  # Subclassed (not a Data.define block) so .new can call super, which lets it
  # accept URL.new(url, text: ...) as well as the usual Data forms.
  class URL < Data.define(:url, :text) # rubocop:disable Style/DataInheritance
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

  # Text with a format per segment, e.g. RichString.new(["Total: ", bold], "1,234").
  # Each segment is a String (default font) or [String, Format].
  class RichString
    attr_reader :segments

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

  # Owns the worksheets; serialize with #to_xlsx or #save.
  class Workbook
    MEMORY_MODES = %i[standard constant low].freeze

    # memory: :standard keeps every cell in memory until saving. :constant and
    # :low write each finished row to disk, so each worksheet must be filled
    # top to bottom. :constant stores strings inline (memory stays flat); :low
    # keeps Excel's shared string table (memory grows with the number of
    # unique strings, output is standard).
    def self.new(memory: :standard)
      unless MEMORY_MODES.include?(memory)
        raise ArgumentError, "unknown memory mode #{memory.inspect} (expected one of #{MEMORY_MODES.join(", ")})"
      end

      _new(memory == :constant, memory == :low)
    end

    def add_worksheet(name = nil)
      _add_worksheet(name).tap { |ws| worksheets << ws }
    end

    # The same Worksheet objects add_worksheet returned, so their append
    # position is shared.
    def worksheets
      @worksheets ||= []
    end

    def worksheet(name)
      worksheets.find { |ws| ws.name == name }
    end

    # path: a String, Pathname or anything responding to #to_path.
    def save(path)
      _save(File.path(path))
    end

    # Document properties shown in Excel's File > Info: title:, subject:,
    # author:, manager:, company:, category:, keywords:, comments:, status:.
    # Later calls add to earlier ones.
    def properties(**fields)
      merged = (@properties || {}).merge(fields)
      _properties(merged) # validates before anything is remembered
      @properties = merged
      self
    end
  end

  # Cell writer for one sheet; create with Workbook#add_worksheet.
  class Worksheet
    def write(row, col, value, format = nil)
      _write(row, col, value, format)
      self
    end

    def append(values, format: nil)
      _append(values, format)
    end

    # columns: a 0-based column index or a Range of them. width is in characters.
    def column_width(columns, width)
      range = CellRange.bounds(columns)
      _column_width(*range, width)
      @fixed_widths ||= {}
      @fixed_widths.delete(range) # re-insert so autofit replays calls in order
      @fixed_widths[range] = width
      self
    end

    # Sizes columns to the data written so far. Widths set with
    # column_width are kept.
    def autofit
      _autofit
      @fixed_widths&.each { |bounds, width| _column_width(*bounds, width) }
      self
    end

    # The methods below take a cell range in any of these styles:
    #   (first_row, first_col, last_row, last_col)  four 0-based numbers
    #   ("A1:D10") or ("B2")                        an Excel reference
    #   (rows, cols)                                Integers or Ranges, e.g. (0..9, 0..3)

    # Filter buttons on the range's first row.
    def autofilter(*range)
      _autofilter(*CellRange.split(range).first)
    end

    # Merges the range and writes value (any cell type) into its first cell.
    def merge_range(*args)
      range, (value, format) = CellRange.split(args, 1..2)
      _merge_range(*range, value, format)
    end

    # Highlights cells in the range by rule. type: :cell, :text, :formula,
    # :data_bar or :color_scale; see the README for each type's options.
    def conditional_format(*range, type:, **)
      _conditional_format(*CellRange.split(range).first, { type: type, ** })
    end

    # Restricts what can be entered in the range. type: :list, :whole_number,
    # :decimal or :text_length; see the README for the options.
    def data_validation(*range, type:, **)
      _data_validation(*CellRange.split(range).first, { type: type, ** })
    end

    # Adds a comment (Excel "note") to a cell.
    def write_comment(row, col, text, author: nil)
      _write_comment(row, col, text, author)
    end

    # Inserts a PNG, JPEG, GIF or BMP image with its top-left corner in the
    # cell. source is a file path or an IO (anything responding to #read).
    # Options: scale: or width:/height: (pixels), x_offset:, y_offset: (pixels), alt_text:.
    def insert_image(row, col, source, **)
      bytes = source.respond_to?(:read) ? source.read : File.binread(source)
      _insert_image(row, col, bytes, { ** })
    end

    # Inserts a chart with its top-left corner in the cell. series is an Array
    # of { values:, categories:, name: } with Excel ranges such as
    # "Sheet1!$B$2:$B$13". Options: title:, x_axis:, y_axis:, width:, height:.
    def insert_chart(row, col, type:, series:, **)
      _insert_chart(row, col, { type: type, series: series, ** })
    end

    # Turns the range (header row included, total row too when total_row: true)
    # into an Excel table. columns: header Strings or { header:, total:,
    # total_label:, format: }; other options: style:, name:, total_row:,
    # banded_rows:, autofilter:.
    def add_table(*range, **)
      _add_table(*CellRange.split(range).first, { ** })
    end

    # Printed page header/footer using Excel codes such as "&CPage &P of &N".
    # margin: is in inches.
    def page_header(text, margin: nil)
      _page_header(text)
      margin ? margins(header: margin) : self
    end

    def page_footer(text, margin: nil)
      _page_footer(text)
      margin ? margins(footer: margin) : self
    end

    # Printing: landscape:, paper: (:letter, :legal, :tabloid, :a3, :a4, :a5 or
    # Excel's paper number), fit_width:/fit_height: (pages; 0 or left out =
    # as many as needed), repeat_rows:/repeat_columns: (an index or Range,
    # printed on every page), print_area: (any cell range), gridlines:.
    def page_setup(repeat_rows: nil, repeat_columns: nil, print_area: nil, **options)
      options[:repeat_rows] = CellRange.bounds(repeat_rows) if repeat_rows
      options[:repeat_columns] = CellRange.bounds(repeat_columns) if repeat_columns
      options[:print_area] = CellRange.split(print_area.is_a?(Array) ? print_area : [print_area]).first if print_area
      _page_setup(options)
      self
    end

    # Print margins in inches; margins not given keep their current value.
    def margins(left: nil, right: nil, top: nil, bottom: nil, header: nil, footer: nil)
      _margins(*[left, right, top, bottom, header, footer].map { |m| m || -1.0 })
      self
    end

    # Outline group with an expand/collapse button. rows: a 0-based row index
    # or a Range; grouping rows already grouped nests them (up to 7 levels).
    # collapsed: true hides them until expanded.
    def group_rows(rows, collapsed: false)
      _group_rows(*CellRange.bounds(rows), collapsed)
    end

    def group_columns(columns, collapsed: false)
      _group_columns(*CellRange.bounds(columns), collapsed)
    end

    # Locks the sheet against editing. Cells whose format has locked: false
    # stay editable. allow: actions users may still take, any of :format_cells,
    # :format_columns, :format_rows, :insert_columns, :insert_rows,
    # :insert_links, :delete_columns, :delete_rows, :sort, :use_autofilter,
    # :use_pivot_tables, :edit_scenarios, :edit_objects.
    def protect(password: nil, allow: [])
      _protect(password, Array(allow))
    end

    # Default format for cells in these columns that are written without one.
    def column_format(columns, format)
      _column_format(*CellRange.bounds(columns), format)
      self
    end
  end

  # Cell ranges in the styles Worksheet methods accept: four 0-based numbers,
  # an Excel reference ("A1:D10", "B2"), or rows and columns as Integers or
  # Ranges.
  module CellRange
    REF = /\A\$?([A-Za-z]{1,3})\$?([1-9]\d*)\z/ # ASCII only: /i also matches the Kelvin sign
    FORMS = 'a range is (first_row, first_col, last_row, last_col), "A1:D10" or (rows, cols)'

    module_function

    # Splits a range off the front of args and checks how many args follow.
    # Returns [[first_row, first_col, last_row, last_col], the args after it].
    def split(args, following = 0..0)
      range, rest = parse(args)
      unless following.cover?(rest.size)
        expected = following.minmax.uniq.join("..")
        raise ArgumentError, "wrong number of arguments after the cell range (given #{rest.size}, " \
                             "expected #{expected}); #{FORMS}"
      end

      [range, rest]
    end

    def parse(args)
      case args
      in [Numeric, Numeric, Numeric, Numeric, *rest] then [args.first(4), rest]
      in [String => ref, *rest] then [excel(ref), rest]
      in [Integer | Range => rows, Integer | Range => cols, *rest]
        [rows_and_cols(rows, cols), rest]
      else
        raise ArgumentError, "expected a cell range, got #{args.inspect}; #{FORMS}"
      end
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

  # Cell style, e.g. Format.new(bold: true). Pass to Worksheet#write.
  class Format
    def self.new(**options)
      # Apply border: first so border_left: etc. override it whatever the order.
      options = { border: options[:border], **options.except(:border) } if options.key?(:border)
      _new(options)
    end
  end
end
