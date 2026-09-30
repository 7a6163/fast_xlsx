# frozen_string_literal: true

$LOAD_PATH.unshift File.expand_path("../lib", __dir__)
require "fast_xlsx"

require "minitest/autorun"
require "roo"
require "tempfile"

# Reads generated workbooks back with roo so tests assert on cell values, not XML.
module XlsxHelpers
  def open_xlsx(workbook)
    file = Tempfile.new(["fast_xlsx", ".xlsx"])
    file.binmode
    file.write(workbook.to_xlsx)
    file.close
    Roo::Excelx.new(file.path)
  end

  # Rows of the given sheet (index or name), nil for empty cells.
  def rows(workbook, sheet = 0)
    xlsx = open_xlsx(workbook)
    xlsx.default_sheet = sheet.is_a?(Integer) ? xlsx.sheets[sheet] : sheet
    return [] unless xlsx.last_row

    (1..xlsx.last_row).map { |r| (1..xlsx.last_column).map { |c| xlsx.cell(r, c) } }
  end
end
