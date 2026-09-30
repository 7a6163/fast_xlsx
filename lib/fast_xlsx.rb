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
      _add_worksheet(name)
    end
  end
end
