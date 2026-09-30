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

  def test_set_column_width_for_single_column_and_range
    wb = FastXlsx::Workbook.new
    ws = wb.add_worksheet
    ws.set_column_width(0, 20)
    ws.set_column_width(2..3, 5)
    ws << %w[a b c d]

    widths = column_widths(wb)
    assert_in_delta 20, widths[1], 1
    assert_nil widths[2]
    assert_in_delta 5, widths[3], 1
    assert_in_delta 5, widths[4], 1
  end

  def test_autofit_widens_columns_to_their_content
    wb = FastXlsx::Workbook.new
    ws = wb.add_worksheet
    ws << ["a much longer piece of text than the default width", "x"]
    ws.autofit

    widths = column_widths(wb)
    assert_operator widths[1], :>, 30
    assert_operator widths[2], :<, widths[1]
  end

  def test_autofilter_covers_the_given_range
    wb = FastXlsx::Workbook.new
    ws = wb.add_worksheet
    ws.concat([%w[id name score], [1, "a", 9], [2, "b", 7]])
    ws.autofilter(0, 0, 2, 2)

    assert_match(/<autoFilter ref="A1:C3"/, sheet_xml(wb))
  end

  def test_font_size_and_name
    wb = FastXlsx::Workbook.new
    wb.add_worksheet.write(0, 0, "x", FastXlsx::Format.new(font_size: 14, font_name: "Arial"))

    style = cell_style(wb, "A1")
    assert_equal 14.0, style[:font_size]
    assert_equal "Arial", style[:font_name]
  end

  def test_font_and_background_colors_from_hex_string_or_integer
    wb = FastXlsx::Workbook.new
    wb.add_worksheet.write(0, 0, "x", FastXlsx::Format.new(font_color: "#FF0000", bg_color: 0x00FF00))

    style = cell_style(wb, "A1")
    assert_equal "FFFF0000", style[:font_color]
    assert_equal "FF00FF00", style[:bg_color]
  end

  def test_invalid_color_raises
    error = assert_raises(ArgumentError) { FastXlsx::Format.new(font_color: "red") }
    assert_includes error.message, "red"
    error = assert_raises(ArgumentError) { FastXlsx::Format.new(bg_color: "#12345") }
    assert_includes error.message, "#12345"
  end

  def test_horizontal_and_vertical_alignment
    wb = FastXlsx::Workbook.new
    ws = wb.add_worksheet
    ws.write(0, 0, "x", FastXlsx::Format.new(align: :center, valign: :top))
    ws.write(0, 1, "y", FastXlsx::Format.new(align: :right, valign: :center))

    assert_equal %w[center top], cell_style(wb, "A1").values_at(:align, :valign)
    assert_equal %w[right center], cell_style(wb, "B1").values_at(:align, :valign)
  end

  def test_invalid_alignment_raises
    error = assert_raises(ArgumentError) { FastXlsx::Format.new(align: :middle) }
    assert_includes error.message, "middle"
  end

  def test_text_wrap
    wb = FastXlsx::Workbook.new
    wb.add_worksheet.write(0, 0, "a\nb", FastXlsx::Format.new(text_wrap: true))

    assert cell_style(wb, "A1")[:text_wrap]
  end

  def test_border_on_all_sides
    wb = FastXlsx::Workbook.new
    wb.add_worksheet.write(0, 0, "x", FastXlsx::Format.new(border: :thin))

    assert_equal({ left: "thin", right: "thin", top: "thin", bottom: "thin" }, cell_style(wb, "A1")[:border])
  end

  def test_border_per_side
    wb = FastXlsx::Workbook.new
    wb.add_worksheet.write(0, 0, "x", FastXlsx::Format.new(border_bottom: :double, border_left: :dashed))

    assert_equal({ left: "dashed", right: nil, top: nil, bottom: "double" }, cell_style(wb, "A1")[:border])
  end

  def test_border_side_wins_over_border_regardless_of_order
    wb = FastXlsx::Workbook.new
    wb.add_worksheet.write(0, 0, "x", FastXlsx::Format.new(border_left: :thick, border: :thin))

    assert_equal({ left: "thick", right: "thin", top: "thin", bottom: "thin" }, cell_style(wb, "A1")[:border])
  end

  def test_invalid_border_raises
    error = assert_raises(ArgumentError) { FastXlsx::Format.new(border: :bold) }
    assert_includes error.message, "bold"
  end

  def test_append_accepts_one_format_per_cell
    wb = FastXlsx::Workbook.new
    bold = FastXlsx::Format.new(bold: true)
    wb.add_worksheet.append(["a", 1, "c"], format: [bold, nil])

    xlsx = open_xlsx(wb)
    assert_equal([true, false, false], (1..3).map { |c| xlsx.font(1, c).bold? })
    assert_equal [["a", 1, "c"]], rows(wb)
  end

  def test_column_format_applies_to_cells_written_without_a_format
    wb = FastXlsx::Workbook.new
    ws = wb.add_worksheet
    ws.set_column_format(1..2, FastXlsx::Format.new(num_format: "#,##0.00"))
    ws << ["a", 1234.5, 2]

    xlsx = open_xlsx(wb)
    assert_equal "General", xlsx.excelx_format(1, 1)
    assert_equal "#,##0.00", xlsx.excelx_format(1, 2)
    assert_equal "#,##0.00", xlsx.excelx_format(1, 3)
  end

  def test_freeze_panes_freezes_rows_above_and_columns_left
    wb = FastXlsx::Workbook.new
    wb.add_worksheet.freeze_panes(1, 2)

    pane = sheet_xml(wb)[/<pane [^>]*>/]
    assert_match(/xSplit="2"/, pane)
    assert_match(/ySplit="1"/, pane)
    assert_match(/state="frozen"/, pane)
  end

  def test_merge_range_merges_cells_and_writes_the_value
    wb = FastXlsx::Workbook.new
    wb.add_worksheet.merge_range(0, 0, 0, 2, 1234, FastXlsx::Format.new(bold: true))

    assert_match(/<mergeCell ref="A1:C1"/, sheet_xml(wb))
    assert_equal 1234, open_xlsx(wb).cell(1, 1)
    assert_predicate open_xlsx(wb).font(1, 1), :bold?
  end

  def test_set_row_height
    wb = FastXlsx::Workbook.new
    ws = wb.add_worksheet
    ws.set_row_height(0, 30)
    ws << ["tall"]

    assert_match(/<row r="1"[^>]* ht="30" customHeight="1"/, sheet_xml(wb))
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
