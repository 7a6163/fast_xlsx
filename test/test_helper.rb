# frozen_string_literal: true

$LOAD_PATH.unshift File.expand_path("../lib", __dir__)
require "fast_xlsx"

require "minitest/autorun"
require "stringio"
require "zip"

module XlsxHelpers
  def xlsx_part(bytes, name)
    Zip::File.open_buffer(StringIO.new(bytes)).read(name)
  end

  def sheet_xml(bytes, index = 1)
    xlsx_part(bytes, "xl/worksheets/sheet#{index}.xml")
  end
end
