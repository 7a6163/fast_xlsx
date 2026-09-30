# frozen_string_literal: true

require "test_helper"
require "date"
require "tmpdir"

class TestFastXlsx < Minitest::Test
  include XlsxHelpers

  def test_that_it_has_a_version_number
    refute_nil ::FastXlsx::VERSION
  end

  def test_numbers_are_written_as_numbers
    wb = FastXlsx::Workbook.new
    wb.add_worksheet << [1, 2.5, 10**12]

    assert_equal [[1, 2.5, 10**12]], rows(wb)
  end

  def test_strings_and_other_objects_are_written_as_text
    wb = FastXlsx::Workbook.new
    wb.add_worksheet << ["hi", :sym]

    assert_equal [%w[hi sym]], rows(wb)
  end

  def test_booleans_are_written_as_booleans
    wb = FastXlsx::Workbook.new
    wb.add_worksheet << [true, false]

    assert_equal [[true, false]], rows(wb)
  end

  def test_nil_leaves_the_cell_empty
    wb = FastXlsx::Workbook.new
    wb.add_worksheet << ["a", nil, "c"]

    assert_equal [["a", nil, "c"]], rows(wb)
  end

  def test_time_is_written_as_excel_serial_number
    wb = FastXlsx::Workbook.new
    wb.add_worksheet << [Time.utc(2000, 1, 1, 12)]

    # 2000-01-01 is serial 36526 in the 1900 date system; noon adds 0.5.
    assert_equal [[36_526.5]], rows(wb)
  end

  def test_date_is_written_as_excel_serial_number
    wb = FastXlsx::Workbook.new
    wb.add_worksheet << [Date.new(2000, 1, 1)]

    assert_equal [[36_526]], rows(wb)
  end

  def test_datetime_uses_its_own_wall_clock_time
    wb = FastXlsx::Workbook.new
    wb.add_worksheet << [DateTime.new(2000, 1, 1, 18, 0, 0, "+08:00")]

    # 18:00 local = 0.75 of a day, regardless of the +08:00 offset.
    assert_equal [[36_526.75]], rows(wb)
  end

  def test_formula_is_written_as_formula
    wb = FastXlsx::Workbook.new
    wb.add_worksheet << [1, 2, FastXlsx::Formula.new("SUM(A1:B1)")]

    assert_equal "SUM(A1:B1)", open_xlsx(wb).formula(1, 3)
  end

  def test_url_is_written_as_hyperlink
    wb = FastXlsx::Workbook.new
    wb.add_worksheet << [FastXlsx::URL.new("https://example.com/a")]

    xlsx = open_xlsx(wb)
    assert_equal "https://example.com/a", xlsx.hyperlink(1, 1)
    assert_equal "https://example.com/a", xlsx.cell(1, 1)
  end

  def test_write_applies_bold_format
    wb = FastXlsx::Workbook.new
    ws = wb.add_worksheet
    ws.write(0, 0, "bold", FastXlsx::Format.new(bold: true))
    ws.write(0, 1, "plain")

    xlsx = open_xlsx(wb)
    assert_predicate xlsx.font(1, 1), :bold?
    refute_predicate xlsx.font(1, 2), :bold?
  end

  def test_write_applies_italic_format
    wb = FastXlsx::Workbook.new
    wb.add_worksheet.write(0, 0, "x", FastXlsx::Format.new(italic: true))

    assert_predicate open_xlsx(wb).font(1, 1), :italic?
  end

  def test_write_applies_underline_format
    wb = FastXlsx::Workbook.new
    wb.add_worksheet.write(0, 0, "x", FastXlsx::Format.new(underline: true))

    assert_predicate open_xlsx(wb).font(1, 1), :underline?
  end

  def test_num_format_turns_serial_number_into_a_date
    wb = FastXlsx::Workbook.new
    wb.add_worksheet.write(0, 0, Date.new(2024, 2, 29), FastXlsx::Format.new(num_format: "yyyy-mm-dd"))

    xlsx = open_xlsx(wb)
    assert_equal "yyyy-mm-dd", xlsx.excelx_format(1, 1)
    assert_equal Date.new(2024, 2, 29), xlsx.cell(1, 1)
  end

  def test_append_applies_format_to_every_cell_and_advances
    wb = FastXlsx::Workbook.new
    ws = wb.add_worksheet
    ws.append(%w[id name], format: FastXlsx::Format.new(bold: true))
    ws << [1, "a"]

    xlsx = open_xlsx(wb)
    assert_equal [%w[id name], [1, "a"]], rows(wb)
    assert(xlsx.font(1, 1).bold? && xlsx.font(1, 2).bold?)
    refute_predicate xlsx.font(2, 1), :bold?
  end

  def test_unknown_format_option_raises
    error = assert_raises(ArgumentError) { FastXlsx::Format.new(bolt: true) }
    assert_includes error.message, "bolt"
  end

  def test_append_continues_after_last_written_row
    wb = FastXlsx::Workbook.new
    ws = wb.add_worksheet
    ws.write(1, 0, "x")
    ws << ["a"] << ["b"]
    more = [["c"]]
    ws.concat(more)

    assert_equal [[nil], ["x"], ["a"], ["b"], ["c"]], rows(wb)
    assert_equal 5, ws.next_row
  end

  def test_write_places_value_at_row_and_column
    wb = FastXlsx::Workbook.new
    wb.add_worksheet.write(0, 2, "here")

    assert_equal "here", open_xlsx(wb).cell(1, 3)
  end

  def test_multiple_worksheets_keep_their_own_rows
    wb = FastXlsx::Workbook.new
    wb.add_worksheet("one") << [1]
    wb.add_worksheet("two") << [2]

    assert_equal %w[one two], open_xlsx(wb).sheets
    assert_equal [[1]], rows(wb, "one")
    assert_equal [[2]], rows(wb, "two")
  end

  def test_constant_memory_writes_all_rows
    wb = FastXlsx::Workbook.new(constant_memory: true)
    wb.add_worksheet.concat(Array.new(1000) { |i| [i, "row #{i}"] })

    written = rows(wb)
    assert_equal 1000, written.size
    assert_equal [999, "row 999"], written.last
  end

  def test_constant_memory_rejects_writes_to_flushed_rows
    wb = FastXlsx::Workbook.new(constant_memory: true)
    ws = wb.add_worksheet
    ws << ["a"] << ["b"]

    assert_raises(FastXlsx::Error) { ws.write(0, 0, "late") }
  end

  def test_constant_memory_allows_writes_to_the_current_row
    wb = FastXlsx::Workbook.new(constant_memory: true)
    ws = wb.add_worksheet
    ws << ["a"]
    ws.write(0, 1, "b")

    assert_equal [%w[a b]], rows(wb)
  end

  def test_invalid_sheet_name_raises
    wb = FastXlsx::Workbook.new
    assert_raises(FastXlsx::Error) { wb.add_worksheet("bad[name]") }
  end

  def test_save_writes_a_readable_file
    Dir.mktmpdir do |dir|
      path = File.join(dir, "out.xlsx")
      wb = FastXlsx::Workbook.new
      wb.add_worksheet << ["saved"]
      wb.save(path)

      assert_equal "saved", Roo::Excelx.new(path).cell(1, 1)
    end
  end
end
