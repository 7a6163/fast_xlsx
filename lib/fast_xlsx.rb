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
  URL = Data.define(:url) do
    def initialize(url:)
      super(url: url.to_s)
    end
  end

  # The native extension looks up FastXlsx::Error, so load it after Error is defined.
  require "fast_xlsx/fast_xlsx"

  # Owns the worksheets; serialize with #to_xlsx or #save.
  class Workbook
    # constant_memory: rows are flushed to disk as they are written, so each
    # worksheet must be filled top to bottom.
    def self.new(constant_memory: false)
      _new(constant_memory)
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
      _set_column_width(*column_bounds(columns), width)
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
