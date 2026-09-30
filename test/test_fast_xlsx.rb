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
