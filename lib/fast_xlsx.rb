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
  require "fast_xlsx/fast_xlsx"

  # Owns the worksheets; serialize with #to_xlsx or #save.
  class Workbook
    # constant_memory: / low_memory: write each finished row to disk, so each
    # worksheet must be filled top to bottom. constant_memory stores strings
    # inline (memory stays flat); low_memory keeps Excel's shared string table
    # (memory grows with the number of unique strings, output is standard).
    def self.new(constant_memory: false, low_memory: false)
      raise ArgumentError, "use constant_memory: or low_memory:, not both" if constant_memory && low_memory

      _new(constant_memory, low_memory)
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

    # Document properties shown in Excel's File > Info: title:, subject:,
    # author:, manager:, company:, category:, keywords:, comments:, status:.
    # Later calls add to earlier ones.
    def set_properties(**fields)
      merged = (@properties || {}).merge(fields)
      _set_properties(merged) # validates before anything is remembered
      @properties = merged
      self
    end
  end

  # Cell writer for one sheet; create with Workbook#add_worksheet.
  class Worksheet
    def write(row, col, value, format = nil)
      _write(row, col, value, format)
    end

    def append(values, format: nil)
      _append(values, format)
    end

    # columns: a 0-based column index or a Range of them. width is in characters.
    def set_column_width(columns, width)
      bounds = column_bounds(columns)
      _set_column_width(*bounds, width)
      @fixed_widths ||= {}
      @fixed_widths.delete(bounds) # re-insert so autofit replays calls in order
      @fixed_widths[bounds] = width
      self
    end

    # Sizes columns to the data written so far. Widths set with
    # set_column_width are kept.
    def autofit
      _autofit
      @fixed_widths&.each { |bounds, width| _set_column_width(*bounds, width) }
      self
    end

    # Merges the range and writes value (any cell type) into its first cell.
    def merge_range(first_row, first_col, last_row, last_col, value, format = nil)
      _merge_range(first_row, first_col, last_row, last_col, value, format)
    end

    # Highlights cells in the range by rule. type: :cell, :text, :formula,
    # :data_bar or :color_scale; see the README for each type's options.
    def conditional_format(first_row, first_col, last_row, last_col, type:, **)
      _conditional_format(first_row, first_col, last_row, last_col, { type: type, ** })
    end

    # Restricts what can be entered in the range. type: :list, :whole_number,
    # :decimal or :text_length; see the README for the options.
    def data_validation(first_row, first_col, last_row, last_col, type:, **)
      _data_validation(first_row, first_col, last_row, last_col, { type: type, ** })
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
    def add_table(first_row, first_col, last_row, last_col, **)
      _add_table(first_row, first_col, last_row, last_col, { ** })
    end

    # Printed page header/footer using Excel codes such as "&CPage &P of &N".
    # margin: is in inches.
    def set_header(text, margin: nil)
      _set_header(text)
      margin ? set_margins(header: margin) : self
    end

    def set_footer(text, margin: nil)
      _set_footer(text)
      margin ? set_margins(footer: margin) : self
    end

    # Print margins in inches; margins not given keep their current value.
    def set_margins(left: nil, right: nil, top: nil, bottom: nil, header: nil, footer: nil)
      _set_margins(*[left, right, top, bottom, header, footer].map { |m| m || -1.0 })
      self
    end

    # Default format for cells in these columns that are written without one.
    def set_column_format(columns, format)
      _set_column_format(*column_bounds(columns), format)
      self
    end

    private

    def column_bounds(columns)
      columns.is_a?(Integer) ? [columns, columns] : columns.minmax
    end
  end

  # Cell style, e.g. Format.new(bold: true). Pass to Worksheet#write.
  class Format
    def self.new(**options)
      # Apply border: first so border_left: etc. override it whatever the order.
      options = { border: options[:border], **options.except(:border) } if options.key?(:border)
      _new(options)
    end
  end
end
