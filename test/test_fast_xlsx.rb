# frozen_string_literal: true

require "test_helper"
require "tmpdir"

class TestFastXlsx < Minitest::Test
  include XlsxHelpers

  def test_that_it_has_a_version_number
    refute_nil ::FastXlsx::VERSION
  end

  def test_writes_each_type
    wb = FastXlsx::Workbook.new
    ws = wb.add_worksheet("Data")
    ws << [1, 2.5, "hi", true, nil, :sym, Time.utc(2000, 1, 1, 12)]
    xml = sheet_xml(wb.to_xlsx)

    assert_includes xml, '<c r="A1"><v>1</v></c>'
    assert_includes xml, '<c r="B1"><v>2.5</v></c>'
    assert_match %r{<c r="C1" t="s"><v>0</v></c>}, xml
    assert_includes xml, '<c r="D1" t="b"><v>1</v></c>'
    refute_includes xml, 'r="E1"'
    assert_match %r{<c r="F1" t="s"><v>1</v></c>}, xml
    assert_includes xml, '<c r="G1"><v>36526.5</v></c>'
    assert_includes xlsx_part(wb.to_xlsx, "xl/sharedStrings.xml"), "<t>sym</t>"
  end

  def test_append_continues_after_last_written_row
    wb = FastXlsx::Workbook.new
    ws = wb.add_worksheet
    ws.write(4, 0, "x")
    ws << ["a"] << ["b"]
    rows = [["c"], ["d"]]
    ws.concat(rows)

    assert_equal 9, ws.next_row
    assert_includes sheet_xml(wb.to_xlsx), 'r="A9"'
  end

  def test_multiple_worksheets
    wb = FastXlsx::Workbook.new
    wb.add_worksheet("one") << [1]
    wb.add_worksheet("two") << [2]
    bytes = wb.to_xlsx

    assert_includes sheet_xml(bytes, 1), "<v>1</v>"
    assert_includes sheet_xml(bytes, 2), "<v>2</v>"
  end

  def test_constant_memory
    wb = FastXlsx::Workbook.new(constant_memory: true)
    ws = wb.add_worksheet
    ws.concat(Array.new(1000) { |i| [i, "row #{i}"] })

    assert_includes sheet_xml(wb.to_xlsx), 'r="B1000"'
  end

  def test_invalid_sheet_name_raises
    wb = FastXlsx::Workbook.new
    assert_raises(FastXlsx::Error) { wb.add_worksheet("bad[name]") }
  end

  def test_save_to_path
    Dir.mktmpdir do |dir|
      path = File.join(dir, "out.xlsx")
      wb = FastXlsx::Workbook.new
      wb.add_worksheet << ["saved"]
      wb.save(path)

      assert_equal "PK", File.binread(path, 2)
    end
  end
end
