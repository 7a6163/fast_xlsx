# frozen_string_literal: true

require "test_helper"
require "date"
require "fileutils"
require "pathname"
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

  # Data read from legacy CSVs is often in a non-UTF-8 encoding.
  def test_strings_in_other_encodings_are_converted_to_utf8
    wb = FastXlsx::Workbook.new
    wb.add_worksheet << ["caf\xE9".dup.force_encoding("ISO-8859-1"), "中文".encode("Big5"),
                         "日本".encode("Shift_JIS"), "plain".b]

    assert_equal [%w[café 中文 日本 plain]], rows(wb)
  end

  def test_binary_strings_with_non_ascii_bytes_raise
    ws = FastXlsx::Workbook.new.add_worksheet
    assert_raises(EncodingError) { ws << ["\xFF\xFE".b] }
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

  def test_time_uses_its_own_wall_clock_time
    wb = FastXlsx::Workbook.new
    wb.add_worksheet << [Time.new(2000, 1, 1, 18, 0, 0, "+08:00"), Time.new(2000, 1, 1, 6, 0, 0, "-05:00")]

    # 18:00 and 06:00 local, whatever their offsets from UTC.
    assert_equal [[36_526.75, 36_526.25]], rows(wb)
  end

  def test_date_is_written_as_excel_serial_number
    wb = FastXlsx::Workbook.new
    wb.add_worksheet << [Date.new(2000, 1, 1)]

    assert_equal [[36_526]], rows(wb)
  end

  # Excel's 1900 date system includes a non-existent 1900-02-29 (serial 60),
  # so real dates before 1900-03-01 are one serial lower.
  def test_dates_before_1900_03_01_match_excel_serials
    wb = FastXlsx::Workbook.new
    wb.add_worksheet << [Date.new(1900, 1, 1), Date.new(1900, 2, 28), Date.new(1900, 3, 1),
                         Time.utc(1900, 1, 1, 12), Time.utc(1900, 3, 1, 12)]

    assert_equal [[1, 59, 61, 1.5, 61.5]], rows(wb)
  end

  def test_dates_before_1900_raise
    ws = FastXlsx::Workbook.new.add_worksheet
    error = assert_raises(ArgumentError) { ws << [Date.new(1899, 12, 31)] }
    assert_includes error.message, "1900"
    assert_raises(ArgumentError) { ws << [Time.utc(1850, 6, 1)] }
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

  def test_url_can_display_other_text
    wb = FastXlsx::Workbook.new
    wb.add_worksheet << [FastXlsx::URL.new("https://example.com/report/42", text: "Q3 report")]

    xlsx = open_xlsx(wb)
    assert_equal "https://example.com/report/42", xlsx.hyperlink(1, 1)
    assert_equal "Q3 report", xlsx.cell(1, 1)
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

  def test_column_width_for_single_column_and_range
    wb = FastXlsx::Workbook.new
    ws = wb.add_worksheet
    ws.column_width(0, 20)
    ws.column_width(2..3, 5)
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

  def test_autofit_keeps_widths_set_explicitly
    wb = FastXlsx::Workbook.new
    ws = wb.add_worksheet
    ws << ["x" * 80, "y" * 80, "z" * 80]
    ws.column_width(0, 12)
    ws.column_width(2..2, 20)
    ws.autofit

    widths = column_widths(wb)
    assert_in_delta 12, widths[1], 1
    assert_operator widths[2], :>, 60
    assert_in_delta 20, widths[3], 1
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
    ws.column_format(1..2, FastXlsx::Format.new(num_format: "#,##0.00"))
    ws << ["a", 1234.5, 2]

    xlsx = open_xlsx(wb)
    assert_equal "General", xlsx.excelx_format(1, 1)
    assert_equal "#,##0.00", xlsx.excelx_format(1, 2)
    assert_equal "#,##0.00", xlsx.excelx_format(1, 3)
  end

  def sheet_protection(workbook)
    Nokogiri::XML(sheet_xml(workbook)).remove_namespaces!.at("sheetProtection")&.to_h
  end

  def test_protect_locks_the_sheet
    wb = FastXlsx::Workbook.new
    ws = wb.add_worksheet

    assert_same ws, ws.protect
    protection = sheet_protection(wb)
    assert_equal "1", protection["sheet"]
    assert_nil protection["password"]
    assert_nil protection["sort"] # not allowed
  end

  # In sheetProtection, "0" means the action is allowed.
  def test_protect_with_a_password_and_allowed_actions
    wb = FastXlsx::Workbook.new
    wb.add_worksheet.protect(password: "password", allow: %i[sort format_cells])

    protection = sheet_protection(wb)
    assert_equal "83AF", protection["password"] # Excel's hash of "password"
    assert_equal %w[0 0], protection.values_at("sort", "formatCells")
    assert_nil protection["insertRows"]
  end

  def test_protect_again_replaces_the_password_and_actions
    wb = FastXlsx::Workbook.new
    ws = wb.add_worksheet
    ws.protect(password: "password", allow: %i[sort])
    ws.protect

    protection = sheet_protection(wb)
    assert_nil protection["password"]
    assert_nil protection["sort"]
  end

  def test_empty_or_reversed_ranges_raise_argument_error
    ws = FastXlsx::Workbook.new.add_worksheet
    [5..2, 2...2].each do |range|
      error = assert_raises(ArgumentError) { ws.group_rows(range) }
      assert_match(/empty range/, error.message)
      assert_raises(ArgumentError) { ws.column_width(range, 10) }
    end
  end

  def test_protect_rejects_unknown_actions
    ws = FastXlsx::Workbook.new.add_worksheet
    error = assert_raises(ArgumentError) { ws.protect(allow: %i[sort fly]) }
    assert_match(/:fly/, error.message)
  end

  def test_format_locked_false_and_hidden_for_protected_sheets
    wb = FastXlsx::Workbook.new
    ws = wb.add_worksheet
    ws.write(0, 0, 1, FastXlsx::Format.new(locked: false))
    ws.write(0, 1, FastXlsx::Formula.new("1+1"), FastXlsx::Format.new(hidden: true))
    ws.write(0, 2, 2, FastXlsx::Format.new(locked: true, hidden: false))

    assert_equal [false, false], cell_style(wb, "A1").values_at(:locked, :hidden)
    assert_equal [true, true], cell_style(wb, "B1").values_at(:locked, :hidden)
    assert_equal [true, false], cell_style(wb, "C1").values_at(:locked, :hidden)
  end

  # Row number => [outlineLevel, hidden] for rows in sheet 1.
  def row_outlines(workbook)
    Nokogiri::XML(sheet_xml(workbook)).remove_namespaces!.css("sheetData row")
            .to_h { |row| [row["r"].to_i, [row["outlineLevel"], row["hidden"]]] }
  end

  def test_group_rows_nests_and_collapses
    wb = FastXlsx::Workbook.new
    ws = wb.add_worksheet
    ws.concat([[1], [2], [3], [4]])

    assert_same ws, ws.group_rows(1..2)
    ws.group_rows(2, collapsed: true)

    assert_equal({ 1 => [nil, nil], 2 => ["1", nil], 3 => %w[2 1], 4 => [nil, nil] }, row_outlines(wb))
  end

  def test_group_columns_collapses
    wb = FastXlsx::Workbook.new
    ws = wb.add_worksheet
    ws << [1, 2, 3, 4]

    assert_same ws, ws.group_columns(1..2, collapsed: true)
    cols = Nokogiri::XML(sheet_xml(wb)).remove_namespaces!.css("cols col")
    assert_includes cols.map { |c| c.to_h.values_at("min", "max", "outlineLevel", "hidden") }, %w[2 3 1 1]
  end

  # rust_xlsxwriter writes no outline levels for rows in these modes.
  def test_group_rows_raises_in_constant_and_low_memory_mode
    %i[constant low].each do |memory|
      ws = FastXlsx::Workbook.new(memory: memory).add_worksheet
      error = assert_raises(FastXlsx::Error) { ws.group_rows(1..2) }
      assert_match(/:standard/, error.message)
    end
  end

  def test_group_columns_works_in_constant_memory_mode
    wb = FastXlsx::Workbook.new(memory: :constant)
    ws = wb.add_worksheet
    ws.group_columns(1)
    ws << [0, 1]

    assert_equal "1", Nokogiri::XML(sheet_xml(wb)).remove_namespaces!.at("cols col[min='2']")["outlineLevel"]
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

  def test_row_height
    wb = FastXlsx::Workbook.new
    ws = wb.add_worksheet
    ws.row_height(0, 30)
    ws << ["tall"]

    assert_match(/<row r="1"[^>]* ht="30" customHeight="1"/, sheet_xml(wb))
  end

  def test_header_and_footer_with_margin
    wb = FastXlsx::Workbook.new
    ws = wb.add_worksheet
    ws.page_header("&CPage &P of &N")
    ws.page_footer("&L&A", margin: 0.2)

    xml = Nokogiri::XML(sheet_xml(wb)).remove_namespaces!
    assert_equal ["&CPage &P of &N", "&L&A"], [xml.at("oddHeader").text, xml.at("oddFooter").text]
    assert_equal "0.2", xml.at("pageMargins")["footer"]
  end

  def test_margins_keeps_defaults_for_the_rest
    wb = FastXlsx::Workbook.new
    wb.add_worksheet.margins(left: 0.5, top: 1)

    margins = Nokogiri::XML(sheet_xml(wb)).remove_namespaces!.at("pageMargins")
    assert_equal %w[0.5 0.7 1 0.75], %w[left right top bottom].map { |side| margins[side] }
  end

  def test_page_breaks
    wb = FastXlsx::Workbook.new
    ws = wb.add_worksheet
    ws.page_breaks([20, 40])
    ws.vertical_page_breaks([5])

    xml = Nokogiri::XML(sheet_xml(wb)).remove_namespaces!
    assert_equal %w[20 40], xml.css("rowBreaks brk").map { |b| b["id"] }
    assert_equal %w[5], xml.css("colBreaks brk").map { |b| b["id"] }
  end

  def test_worksheet_by_name_returns_the_same_worksheet
    wb = FastXlsx::Workbook.new
    data = wb.add_worksheet("Data")
    data << ["a"]
    wb.worksheet("Data") << ["b"]

    assert_same data, wb.worksheet("Data")
    assert_nil wb.worksheet("missing")
    assert_equal [["a"], ["b"]], rows(wb)
  end

  def test_worksheets_lists_sheets_with_their_names
    wb = FastXlsx::Workbook.new
    wb.add_worksheet
    wb.add_worksheet("Summary")

    assert_equal %w[Sheet1 Summary], wb.worksheets.map(&:name)
  end

  def test_strikeout_and_font_script
    wb = FastXlsx::Workbook.new
    ws = wb.add_worksheet
    ws.write(0, 0, "x", FastXlsx::Format.new(strikeout: true, font_script: :superscript))
    ws.write(0, 1, "y", FastXlsx::Format.new(font_script: :subscript))

    assert_equal [true, "superscript"], cell_style(wb, "A1").values_at(:strikeout, :script)
    assert_equal [false, "subscript"], cell_style(wb, "B1").values_at(:strikeout, :script)
  end

  def test_underline_styles
    wb = FastXlsx::Workbook.new
    ws = wb.add_worksheet
    ws.write(0, 0, "a", FastXlsx::Format.new(underline: :double))
    ws.write(0, 1, "b", FastXlsx::Format.new(underline: :single_accounting))
    ws.write(0, 2, "c", FastXlsx::Format.new(underline: true))

    assert_equal(%w[double singleAccounting single], %w[A1 B1 C1].map { |ref| cell_style(wb, ref)[:underline] })
  end

  def test_rotation_indent_and_shrink
    wb = FastXlsx::Workbook.new
    wb.add_worksheet.write(0, 0, "x", FastXlsx::Format.new(rotation: 45, indent: 2, shrink: true))

    assert_equal [45, 2, true], cell_style(wb, "A1").values_at(:rotation, :indent, :shrink)
  end

  def test_rotation_out_of_range_raises
    error = assert_raises(ArgumentError) { FastXlsx::Format.new(rotation: 120) }
    assert_includes error.message, "120"
  end

  def test_border_color
    wb = FastXlsx::Workbook.new
    wb.add_worksheet.write(0, 0, "x", FastXlsx::Format.new(border: :thin, border_color: "#FF0000"))

    assert_equal %w[FFFF0000] * 4, cell_style(wb, "A1")[:border_color].values
  end

  def test_conditional_format_cell_rule_with_format
    wb = FastXlsx::Workbook.new
    red = FastXlsx::Format.new(font_color: "#FF0000")
    wb.add_worksheet.conditional_format(0, 1, 9, 1, type: :cell, criteria: :<, value: 0, format: red)

    rule = conditional_formats(wb).first
    assert_equal ["B1:B10", "cellIs", "lessThan", ["0"], "FFFF0000"],
                 rule.values_at(:sqref, :type, :operator, :formulas, :font_color)
  end

  def test_conditional_format_cell_between_takes_two_values
    wb = FastXlsx::Workbook.new
    wb.add_worksheet.conditional_format(0, 0, 0, 0, type: :cell, criteria: :between, value: [1, 10],
                                                    format: FastXlsx::Format.new(bold: true))

    assert_equal ["between", %w[1 10]], conditional_formats(wb).first.values_at(:operator, :formulas)
  end

  def test_conditional_format_text_contains
    wb = FastXlsx::Workbook.new
    wb.add_worksheet.conditional_format(0, 0, 9, 0, type: :text, criteria: :contains, value: "error",
                                                    format: FastXlsx::Format.new(bold: true))

    assert_equal %w[containsText containsText error],
                 conditional_formats(wb).first.values_at(:type, :operator, :text)
  end

  def test_conditional_format_other_text_criteria
    {
      not_contains: %w[notContainsText notContains],
      begins_with: %w[beginsWith beginsWith],
      ends_with: %w[endsWith endsWith]
    }.each do |criteria, (type, operator)|
      wb = FastXlsx::Workbook.new
      wb.add_worksheet.conditional_format(0, 0, 9, 0, type: :text, criteria: criteria, value: "x",
                                                      format: FastXlsx::Format.new(bold: true))

      assert_equal [type, operator, "x"], conditional_formats(wb).first.values_at(:type, :operator, :text), criteria
    end
  end

  def test_conditional_format_formula
    wb = FastXlsx::Workbook.new
    wb.add_worksheet.conditional_format(0, 0, 9, 3, type: :formula, value: "=$D1>100",
                                                    format: FastXlsx::Format.new(bold: true))

    assert_equal ["A1:D10", "expression", ["$D1>100"]],
                 conditional_formats(wb).first.values_at(:sqref, :type, :formulas)
  end

  def test_conditional_format_data_bar_and_color_scales
    wb = FastXlsx::Workbook.new
    ws = wb.add_worksheet
    ws.conditional_format(0, 0, 9, 0, type: :data_bar)
    ws.conditional_format(0, 1, 9, 1, type: :color_scale)
    ws.conditional_format(0, 2, 9, 2, type: :color_scale, colors: 2)

    rules = conditional_formats(wb)
    assert_equal(%w[dataBar colorScale colorScale], rules.map { |r| r[:type] })
    assert_equal([3, 2], rules.drop(1).map { |r| r[:stops] })
  end

  def test_conditional_format_rejects_unknown_type_and_criteria
    ws = FastXlsx::Workbook.new.add_worksheet
    error = assert_raises(ArgumentError) { ws.conditional_format(0, 0, 0, 0, type: :sparkle) }
    assert_includes error.message, "sparkle"
    error = assert_raises(ArgumentError) { ws.conditional_format(0, 0, 0, 0, type: :cell, criteria: :~, value: 1) }
    assert_includes error.message, "~"
  end

  def test_data_validation_list_of_strings
    wb = FastXlsx::Workbook.new
    wb.add_worksheet.data_validation(1, 2, 9, 2, type: :list, value: %w[Open Closed])

    assert_equal ["C2:C10", "list", '"Open,Closed"'], data_validations(wb).first.values_at(:sqref, :type, :formula1)
  end

  def test_data_validation_list_from_a_range
    wb = FastXlsx::Workbook.new
    wb.add_worksheet.data_validation(0, 0, 0, 0, type: :list, value: "=$Z$1:$Z$3")

    assert_equal %w[list $Z$1:$Z$3], data_validations(wb).first.values_at(:type, :formula1)
  end

  def test_data_validation_list_of_numbers_and_mixed_values
    wb = FastXlsx::Workbook.new
    ws = wb.add_worksheet
    ws.data_validation(0, 0, 0, 0, type: :list, value: [1, 2, 3])
    ws.data_validation(0, 1, 0, 1, type: :list, value: [1.5, "N/A", :other])

    assert_equal ['"1,2,3"', '"1.5,N/A,other"'], data_validations(wb).map { |dv| dv[:formula1] }
  end

  def test_data_validation_number_rules
    wb = FastXlsx::Workbook.new
    ws = wb.add_worksheet
    ws.data_validation(0, 0, 0, 0, type: :whole_number, criteria: :between, value: [1, 10])
    ws.data_validation(0, 1, 0, 1, type: :decimal, criteria: :>=, value: 0.5)
    ws.data_validation(0, 2, 0, 2, type: :text_length, criteria: :<=, value: 50)

    assert_equal([
                   ["whole", nil, "1", "10"],
                   ["decimal", "greaterThanOrEqual", "0.5", nil],
                   ["textLength", "lessThanOrEqual", "50", nil]
                 ], data_validations(wb).map { |dv| dv.values_at(:type, :operator, :formula1, :formula2) })
  end

  def test_data_validation_messages
    wb = FastXlsx::Workbook.new
    wb.add_worksheet.data_validation(0, 0, 0, 0, type: :list, value: %w[Y N],
                                                 input_title: "Pick", input_message: "Y or N",
                                                 error_title: "Oops", error_message: "Only Y or N")

    assert_equal ["Pick", "Y or N", "Oops", "Only Y or N"],
                 data_validations(wb).first.values_at(:input_title, :input_message, :error_title, :error_message)
  end

  def test_data_validation_rejects_unknown_type
    ws = FastXlsx::Workbook.new.add_worksheet
    error = assert_raises(ArgumentError) { ws.data_validation(0, 0, 0, 0, type: :email) }
    assert_includes error.message, "email"
  end

  def test_write_comment_with_and_without_author
    wb = FastXlsx::Workbook.new
    ws = wb.add_worksheet
    ws.write_comment(0, 0, "Checked by finance", author: "Zac")
    ws.write_comment(2, 1, "Estimate")

    notes = comments(wb)
    assert_equal ["Zac", "Checked by finance"], notes["A1"]
    assert_equal "Estimate", notes["B3"].last
  end

  def test_insert_image_from_path
    Dir.mktmpdir do |dir|
      path = File.join(dir, "logo.png")
      File.binwrite(path, PNG_1X1)
      wb = FastXlsx::Workbook.new
      wb.add_worksheet.insert_image(2, 1, path)

      anchors, media = images(wb)
      assert_equal 1, media
      assert_equal [1, 2, 9525, 9525], anchors.first.values_at(:col, :row, :cx, :cy)
    end
  end

  def test_insert_image_from_io_with_scale_offset_and_alt_text
    wb = FastXlsx::Workbook.new
    wb.add_worksheet.insert_image(0, 0, StringIO.new(PNG_1X1), scale: 2, x_offset: 10, y_offset: 5,
                                                               alt_text: "Company logo")

    anchor = images(wb).first.first
    # 1 px = 9525 EMU at 96 DPI.
    assert_equal [19_050, 19_050, 95_250, 47_625, "Company logo"],
                 anchor.values_at(:cx, :cy, :col_off, :row_off, :alt_text)
  end

  def test_insert_image_with_pixel_width_and_height
    wb = FastXlsx::Workbook.new
    ws = wb.add_worksheet
    ws.insert_image(0, 0, StringIO.new(PNG_4X2), width: 40, height: 30)
    ws.insert_image(5, 0, StringIO.new(PNG_4X2), width: 20) # keeps the 2:1 aspect ratio
    ws.insert_image(10, 0, StringIO.new(PNG_4X2), height: 6)

    sizes = images(wb).first.map { |a| [a[:cx] / 9525, a[:cy] / 9525] }
    assert_equal [[40, 30], [20, 10], [12, 6]], sizes
  end

  def test_insert_image_rejects_scale_with_width_or_height
    ws = FastXlsx::Workbook.new.add_worksheet
    assert_raises(ArgumentError) { ws.insert_image(0, 0, StringIO.new(PNG_1X1), scale: 2, width: 40) }
  end

  def test_insert_image_rejects_non_image_data
    ws = FastXlsx::Workbook.new.add_worksheet
    assert_raises(FastXlsx::Error) { ws.insert_image(0, 0, StringIO.new("not an image")) }
  end

  SALES = [%w[Month Sales], ["Jan", 10], ["Feb", 25], ["Mar", 18]].freeze
  SALES_SERIES = { name: "Sales", categories: "Sheet1!$A$2:$A$4", values: "Sheet1!$B$2:$B$4" }.freeze

  def test_insert_chart_places_a_column_chart_with_its_series
    wb = FastXlsx::Workbook.new
    ws = wb.add_worksheet.concat(SALES)
    ws.insert_chart(1, 3, type: :column, series: [SALES_SERIES], width: 600, height: 360)

    c = chart(wb)
    assert_equal %w[barChart col], c.values_at(:plot, :bar_dir)
    assert_equal [["Sales", "Sheet1!$A$2:$A$4", "Sheet1!$B$2:$B$4"]], c[:series]
    assert_equal [3, 1], c[:anchor]
    assert_equal [600, 360], c[:size]
  end

  def test_insert_chart_titles
    wb = FastXlsx::Workbook.new
    wb.add_worksheet.concat(SALES).insert_chart(0, 3, type: :line, series: [SALES_SERIES],
                                                      title: "Monthly sales", x_axis: "Month", y_axis: "Amount")

    c = chart(wb)
    assert_equal ["lineChart", "Monthly sales", %w[Month Amount]], c.values_at(:plot, :title, :axis_titles)
  end

  def test_insert_chart_types
    {
      bar: %w[barChart bar clustered], column_stacked: %w[barChart col stacked],
      pie: ["pieChart", nil, nil], area: ["areaChart", nil, "standard"]
    }.each do |type, expected|
      wb = FastXlsx::Workbook.new
      wb.add_worksheet.concat(SALES).insert_chart(0, 3, type: type, series: [SALES_SERIES])

      assert_equal expected, chart(wb).values_at(:plot, :bar_dir, :grouping), type
    end
  end

  def test_insert_chart_rejects_unknown_type_and_missing_series
    ws = FastXlsx::Workbook.new.add_worksheet
    error = assert_raises(ArgumentError) { ws.insert_chart(0, 0, type: :bubble_tea, series: [SALES_SERIES]) }
    assert_includes error.message, "bubble_tea"
    assert_raises(ArgumentError) { ws.insert_chart(0, 0, type: :column, series: []) }
  end

  def test_autofit_restores_widths_in_the_order_they_were_set
    wb = FastXlsx::Workbook.new
    ws = wb.add_worksheet
    ws << (["x" * 80] * 4)
    ws.column_width(0..3, 10)
    ws.column_width(1, 30)
    ws.column_width(0..3, 12) # the last call wins for column B too
    ws.autofit

    assert_in_delta 12, column_widths(wb)[2], 1
  end

  def test_url_constructor_forms
    assert_equal FastXlsx::URL.new("https://a"), FastXlsx::URL.new(url: "https://a")
    assert_equal FastXlsx::URL.new("https://a"), FastXlsx::URL["https://a"]
    assert_equal "5", FastXlsx::URL.new("https://a").with(text: 5).text
    assert_raises(ArgumentError) { FastXlsx::URL.new("https://a", "text", "extra") }
  end

  def test_properties_adds_to_earlier_calls
    wb = FastXlsx::Workbook.new
    wb.properties(title: "Q3")
    wb.properties(author: "Zac")
    wb.add_worksheet

    core = Nokogiri::XML(Zip::File.open_buffer(StringIO.new(wb.to_xlsx)).read("docProps/core.xml")).remove_namespaces!
    assert_equal %w[Q3 Zac], [core.at("title")&.text, core.at("creator")&.text]
  end

  def test_header_and_footer_over_255_characters_raise
    ws = FastXlsx::Workbook.new.add_worksheet
    assert_raises(ArgumentError) { ws.page_header("x" * 256) }
    assert_raises(ArgumentError) { ws.page_footer("x" * 256) }
    ws.page_header("&[Page]#{"x" * 253}") # &[Page] counts as &P, so this is 255
  end

  def test_rich_string_mixes_formats_in_one_cell
    wb = FastXlsx::Workbook.new
    bold = FastXlsx::Format.new(bold: true)
    italic = FastXlsx::Format.new(italic: true)
    wb.add_worksheet << ["plain", FastXlsx::RichString.new(["Total: ", bold], "1,234 ", ["(est.)", italic])]

    assert_equal "plain", rows(wb).first.first # roo renders the rich cell as HTML, so check the runs below
    assert_equal [["Total: ", true, false], ["1,234 ", false, false], ["(est.)", false, true]], rich_runs(wb)
  end

  def test_rich_string_takes_a_cell_format
    wb = FastXlsx::Workbook.new
    wb.add_worksheet.write(0, 0, FastXlsx::RichString.new(["a", FastXlsx::Format.new(bold: true)], "b"),
                           FastXlsx::Format.new(align: :center))

    assert_equal "center", cell_style(wb, "A1")[:align]
  end

  def test_rich_string_rejects_a_non_format_segment
    ws = FastXlsx::Workbook.new.add_worksheet
    assert_raises(TypeError) { ws << [FastXlsx::RichString.new(%w[text bold])] }
    assert_raises(ArgumentError) { FastXlsx::RichString.new }
  end

  def test_add_table_with_headers_and_style
    wb = FastXlsx::Workbook.new
    ws = wb.add_worksheet
    ws.write(1, 0, "North")
    ws.add_table(0, 0, 3, 1, columns: %w[Region Sales], style: :medium2)

    t = table(wb)
    assert_equal ["A1:B4", "TableStyleMedium2", true], t.values_at(:ref, :style, :autofilter)
    assert_equal [["Region", nil, nil], ["Sales", nil, nil]], t[:columns]
    assert_equal [%w[Region Sales], ["North", nil]], rows(wb).first(2)
  end

  def test_add_table_total_row
    wb = FastXlsx::Workbook.new
    ws = wb.add_worksheet
    ws.add_table(0, 0, 4, 1, total_row: true,
                             columns: [{ header: "Region", total_label: "Total" }, { header: "Sales", total: :sum }])

    t = table(wb)
    assert_equal ["A1:B5", 1], t.values_at(:ref, :totals)
    assert_equal [["Region", nil, "Total"], %w[Sales sum] + [nil]], t[:columns]
  end

  def test_add_table_name_and_options
    wb = FastXlsx::Workbook.new
    wb.add_worksheet.add_table(0, 0, 2, 0, columns: %w[A], name: "Sales", banded_rows: false, autofilter: false)

    assert_equal ["Sales", "0", false], table(wb).values_at(:name, :banded_rows, :autofilter)
  end

  def test_add_table_then_append_fills_the_table
    %i[standard constant low].each do |memory|
      wb = FastXlsx::Workbook.new(memory: memory)
      ws = wb.add_worksheet
      ws.add_table(0, 0, 2, 1, columns: %w[Region Sales])
      ws << ["North", 10] << ["South", 20]

      assert_equal [%w[Region Sales], ["North", 10], ["South", 20]], rows(wb), "memory: #{memory}"
    end
  end

  def test_table_column_format_applies_to_a_merged_value
    money = FastXlsx::Format.new(num_format: "#,##0.00")
    wb = FastXlsx::Workbook.new
    ws = wb.add_worksheet
    ws.add_table(0, 0, 5, 1, columns: ["Region", { header: "Sales", format: money }])
    ws.merge_range(1, 1, 2, 1, 1234.5)

    assert_equal "#,##0.00", open_xlsx(wb).excelx_format(2, 2)
  end

  def test_table_column_format_applies_to_rows_appended_later
    money = FastXlsx::Format.new(num_format: "#,##0.00")
    bold = FastXlsx::Format.new(bold: true)
    %i[standard constant].each do |memory|
      wb = FastXlsx::Workbook.new(memory: memory)
      ws = wb.add_worksheet
      ws.add_table(0, 0, 3, 1, columns: ["Region", { header: "Sales", format: money }])
      ws.concat([["North", 1.5], ["South", 2.5]])
      ws.append(["East", 3.5], format: bold) # an explicit format wins
      ws << ["outside", 4.5]                 # row 5 is below the table

      xlsx = open_xlsx(wb)
      assert_equal ["#,##0.00", "#,##0.00"], [xlsx.excelx_format(2, 2), xlsx.excelx_format(3, 2)], "memory: #{memory}"
      assert_equal "General", xlsx.excelx_format(4, 2), "memory: #{memory}"
      assert_equal "General", xlsx.excelx_format(5, 2), "memory: #{memory}"
    end
  end

  def test_table_column_format_skips_the_header_row
    money = FastXlsx::Format.new(num_format: "#,##0.00")
    wb = FastXlsx::Workbook.new
    ws = wb.add_worksheet
    ws.add_table(0, 0, 3, 1, columns: ["Region", { header: "Sales", format: money }])
    ws.write(0, 1, "Sales (USD)")

    assert_equal "General", open_xlsx(wb).excelx_format(1, 2)
  end

  def test_add_table_after_its_header_row_was_flushed_raises
    wb = FastXlsx::Workbook.new(memory: :constant)
    ws = wb.add_worksheet
    ws << ["x"] << ["y"] << ["z"]
    assert_raises(FastXlsx::Error) { ws.add_table(0, 0, 2, 0, columns: %w[A]) }
  end

  def test_add_table_validates_columns_style_and_options
    ws = FastXlsx::Workbook.new.add_worksheet
    error = assert_raises(ArgumentError) { ws.add_table(0, 0, 3, 2, columns: %w[A B]) }
    assert_includes error.message, "3"
    error = assert_raises(ArgumentError) { ws.add_table(0, 0, 3, 0, columns: %w[A], style: :medium99) }
    assert_includes error.message, "medium99"
    error = assert_raises(ArgumentError) { ws.add_table(0, 0, 3, 0, columns: [{ header: "A", totl: :sum }]) }
    assert_includes error.message, "totl"
  end

  def test_option_typos_raise_instead_of_being_ignored
    ws = FastXlsx::Workbook.new.add_worksheet
    png = -> { StringIO.new(PNG_1X1) }
    {
      "formt" => -> { ws.conditional_format(0, 0, 0, 0, type: :data_bar, formt: nil) },
      "error_msg" => -> { ws.data_validation(0, 0, 0, 0, type: :list, value: %w[a], error_msg: "x") },
      "scael" => -> { ws.insert_image(0, 0, png.call, scael: 2) },
      "titel" => -> { ws.insert_chart(0, 0, type: :line, series: [SALES_SERIES], titel: "x") },
      "valeus" => -> { ws.insert_chart(0, 0, type: :line, series: [{ values: "Sheet1!$B$2:$B$4", valeus: "x" }]) }
    }.each do |typo, call|
      error = assert_raises(ArgumentError, typo, &call)
      assert_includes error.message, typo
    end
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

  # Cell values are converted with Ruby calls (to_s, jd, ...). If the workbook
  # lock were held during those calls, Ruby code touching the same workbook
  # would deadlock the process, so this runs in a child process with a timeout.
  def test_ruby_called_during_a_write_can_use_the_same_workbook
    script = <<~'RUBY'
      require "fast_xlsx"
      ws = FastXlsx::Workbook.new.add_worksheet
      label = Object.new
      label.define_singleton_method(:to_s) { "sheet #{ws.name}" }
      ws << [label]
      ws.write(1, 0, label)
      ws.merge_range(2, 0, 2, 1, label)
      print "ok"
    RUBY
    reader, writer = IO.pipe
    pid = Process.spawn(RbConfig.ruby, "-I", File.expand_path("../lib", __dir__), "-e", script, out: writer)
    writer.close
    deadline = Time.now + 15
    sleep 0.1 until (done = Process.wait2(pid, Process::WNOHANG)) || Time.now > deadline
    unless done
      Process.kill(:KILL, pid)
      Process.wait(pid)
      flunk "deadlocked: writing a value whose to_s reads the worksheet did not finish in 15s"
    end
    assert_equal "ok", reader.read
  end

  def test_a_row_that_fails_midway_writes_nothing
    wb = FastXlsx::Workbook.new
    ws = wb.add_worksheet
    bad = Object.new
    def bad.to_s = raise(ArgumentError, "boom")
    assert_raises(ArgumentError) { ws << [4, 5, bad] }
    ws << [9]

    assert_equal [[9]], rows(wb)
  end

  # Values only rust_xlsxwriter rejects, or that fail late: the row must still
  # come out empty, not half written.
  def test_a_row_that_fails_to_write_leaves_no_cells
    {
      "string over 32,767 chars" => "x" * 40_000,
      "invalid UTF-8" => "caf\xE9".dup.force_encoding("UTF-8"),
      "URL over 2,083 chars" => FastXlsx::URL.new("https://example.com/#{"a" * 3000}")
    }.each do |label, bad|
      wb = FastXlsx::Workbook.new
      ws = wb.add_worksheet
      assert_raises(StandardError, label) { ws << ["a", "b", bad] }
      ws << ["z"]

      assert_equal [["z"]], rows(wb), label
    end
  end

  # The limit counts characters, not bytes: "é" is 2 bytes.
  def test_strings_up_to_excels_limit_are_written
    longest = "é#{"x" * 32_766}" # 32,767 characters, 32,768 bytes
    multibyte = "é" * 20_000
    wb = FastXlsx::Workbook.new
    wb.add_worksheet << [longest, multibyte]

    assert_equal [[longest, multibyte]], rows(wb)
  end

  def test_a_row_with_more_columns_than_excel_allows_leaves_no_cells
    wb = FastXlsx::Workbook.new
    ws = wb.add_worksheet
    assert_raises(ArgumentError) { ws << Array.new(16_385, 1) }
    ws << ["z"]

    assert_equal [["z"]], rows(wb)
  end

  def test_a_row_that_fails_to_write_leaves_no_hyperlink
    wb = FastXlsx::Workbook.new
    ws = wb.add_worksheet
    assert_raises(FastXlsx::Error) { ws << [FastXlsx::URL.new("https://example.com"), "x" * 40_000] }

    refute_match(/hyperlink/, sheet_xml(wb))
  end

  def test_an_invalid_merge_range_keeps_the_cell_already_there
    wb = FastXlsx::Workbook.new
    ws = wb.add_worksheet
    ws.write(0, 0, "keep")
    assert_raises(FastXlsx::Error) { ws.merge_range(0, 0, 0, 0, "x") }

    assert_equal [["keep"]], rows(wb)
  end

  # A merge that fails must not push earlier rows to disk behind the
  # flushed-row check, or later writes to them would be dropped silently.
  def test_an_invalid_merge_range_does_not_flush_earlier_rows
    wb = FastXlsx::Workbook.new(memory: :constant)
    ws = wb.add_worksheet
    ws.write(0, 0, "a")
    assert_raises(FastXlsx::Error) { ws.merge_range(5, 0, 4, 0, "x") }
    ws.write(0, 1, "b")

    assert_equal [%w[a b]], rows(wb)
  end

  def test_merge_range_with_a_bad_value_leaves_no_merge
    %i[standard constant].each do |memory|
      wb = FastXlsx::Workbook.new(memory: memory)
      ws = wb.add_worksheet
      assert_raises(StandardError) { ws.merge_range(0, 0, 0, 2, "x" * 40_000) }
      ws.write(0, 0, "ok")

      refute_match(/mergeCell/, sheet_xml(wb), "memory: #{memory}")
      assert_equal [["ok"]], rows(wb), "memory: #{memory}"
    end
  end

  # Iterating a row must not dup it: dup'ing an Array longer than 3 elements
  # turns the caller's array into a shared root, one extra live object per row.
  def test_writing_rows_does_not_retain_ruby_objects
    rows = Array.new(10_000) { |i| [i, "s", i * 2, i * 3, i * 4] }
    {
      "<<" => ->(ws) { rows.each { |r| ws << r } },
      "concat" => ->(ws) { ws.concat(rows) }
    }.each do |name, write|
      ws = FastXlsx::Workbook.new(memory: :constant).add_worksheet
      GC.start
      before = GC.stat(:heap_live_slots)
      write.call(ws)
      GC.start

      assert_operator GC.stat(:heap_live_slots) - before, :<, 1_000, name
    end
  end

  def test_low_memory_uses_the_shared_string_table
    wb = FastXlsx::Workbook.new(memory: :low)
    wb.add_worksheet.concat(Array.new(1000) { |i| [i, %w[North South][i % 2]] })

    xml = sheet_xml(wb)
    assert_match(/<c r="B1" t="s">/, xml)
    refute_match(/inlineStr/, xml)
    assert_equal [999, "South"], rows(wb).last
  end

  def test_constant_memory_stores_strings_inline
    wb = FastXlsx::Workbook.new(memory: :constant)
    wb.add_worksheet << ["North"]

    assert_match(/<c r="A1" t="inlineStr">/, sheet_xml(wb))
  end

  def test_low_memory_rejects_writes_to_flushed_rows
    ws = FastXlsx::Workbook.new(memory: :low).add_worksheet
    ws << ["a"] << ["b"]
    assert_raises(FastXlsx::Error) { ws.write(0, 0, "late") }
  end

  def test_unknown_memory_mode_raises
    error = assert_raises(ArgumentError) { FastXlsx::Workbook.new(memory: :tiny) }
    assert_includes error.message, "tiny"
  end

  def test_old_memory_options_are_gone
    assert_raises(ArgumentError) { FastXlsx::Workbook.new(constant_memory: true) }
  end

  def test_write_returns_the_worksheet_for_chaining
    ws = FastXlsx::Workbook.new.add_worksheet
    assert_same ws, ws.write(0, 0, "a")
  end

  def test_set_prefixed_names_are_gone
    ws = FastXlsx::Workbook.new.add_worksheet
    %i[set_column_width set_column_format set_row_height set_page_breaks set_vertical_page_breaks
       set_header set_footer set_margins].each { |m| refute_respond_to ws, m }
    refute_respond_to FastXlsx::Workbook.new, :set_properties
  end

  def test_saving_twice_and_writing_after_save
    %i[standard constant low].each do |memory|
      wb = FastXlsx::Workbook.new(memory: memory)
      ws = wb.add_worksheet
      ws << ["a", 1]
      assert_equal rows(wb), rows(wb), "save twice, memory: #{memory}"
      ws << ["b", 2]
      assert_equal [["a", 1], ["b", 2]], rows(wb), "write after save, memory: #{memory}"
    end
  end

  def test_constant_memory_writes_all_rows
    wb = FastXlsx::Workbook.new(memory: :constant)
    wb.add_worksheet.concat(Array.new(1000) { |i| [i, "row #{i}"] })

    written = rows(wb)
    assert_equal 1000, written.size
    assert_equal [999, "row 999"], written.last
  end

  # The rows a tall merge spans are held back by rust_xlsxwriter, not yet on
  # disk, so cells beside the merge can still be written.
  def test_cells_beside_a_tall_merge_can_be_written
    %i[constant low].each do |memory|
      wb = FastXlsx::Workbook.new(memory: memory)
      ws = wb.add_worksheet
      ws.merge_range(0, 0, 2, 0, "Tall")
      ws.write(1, 1, "x").write(2, 1, "y")
      ws << ["after"]

      assert_equal [["Tall", nil], [nil, "x"], [nil, "y"], ["after", nil]], rows(wb), "memory: #{memory}"
      assert_match(/<mergeCell ref="A1:A3"/, sheet_xml(wb))
    end
  end

  def test_merge_range_over_flushed_rows_raises
    %i[constant low].each do |memory|
      ws = FastXlsx::Workbook.new(memory: memory).add_worksheet
      ws << ["a"] << ["b"] << ["c"]
      assert_raises(FastXlsx::Error, "memory: #{memory}") { ws.merge_range(0, 0, 0, 2, "Title") }
    end
  end

  def test_constant_memory_rejects_writes_to_flushed_rows
    wb = FastXlsx::Workbook.new(memory: :constant)
    ws = wb.add_worksheet
    ws << ["a"] << ["b"]

    assert_raises(FastXlsx::Error) { ws.write(0, 0, "late") }
  end

  def test_constant_memory_allows_writes_to_the_current_row
    wb = FastXlsx::Workbook.new(memory: :constant)
    ws = wb.add_worksheet
    ws << ["a"]
    ws.write(0, 1, "b")

    assert_equal [%w[a b]], rows(wb)
  end

  def test_invalid_sheet_name_raises
    wb = FastXlsx::Workbook.new
    assert_raises(FastXlsx::Error) { wb.add_worksheet("bad[name]") }
  end

  def test_invalid_sheet_name_leaves_no_sheet_behind
    %i[standard constant].each do |memory|
      wb = FastXlsx::Workbook.new(memory: memory)
      assert_raises(FastXlsx::Error) { wb.add_worksheet("x" * 40) }
      wb.add_worksheet("Good") << ["ok"]

      assert_equal %w[Good], open_xlsx(wb).sheets, "memory: #{memory}"
      assert_equal %w[Good], wb.worksheets.map(&:name)
    end
  end

  # An unnamed sheet takes the next free "SheetN", like Excel, instead of a
  # default name that clashes and fails only when saving.
  def test_unnamed_sheet_skips_names_already_taken
    wb = FastXlsx::Workbook.new
    wb.add_worksheet("Sheet2")
    first = wb.add_worksheet
    second = wb.add_worksheet

    assert_equal %w[Sheet2 Sheet1 Sheet3], [wb.worksheet("Sheet2").name, first.name, second.name]
    assert_equal %w[Sheet2 Sheet1 Sheet3], open_xlsx(wb).sheets
  end

  # Excel sheet names are case-insensitive; the clash should surface at
  # add_worksheet, not at save.
  def test_duplicate_sheet_name_raises_when_added
    wb = FastXlsx::Workbook.new
    wb.add_worksheet("Data")
    error = assert_raises(FastXlsx::Error) { wb.add_worksheet("data") }
    assert_includes error.message, "data"

    assert_equal %w[Data], open_xlsx(wb).sheets
  end

  def defined_names(workbook)
    zip = Zip::File.open_buffer(StringIO.new(workbook.to_xlsx))
    Nokogiri::XML(zip.read("xl/workbook.xml")).remove_namespaces!.css("definedName")
            .map { |n| [n["name"], n["localSheetId"], n.text] }
  end

  def test_define_name_for_the_workbook_and_for_one_sheet
    wb = FastXlsx::Workbook.new
    wb.add_worksheet("Data")

    assert_same wb, wb.define_name("Rate", "=0.96")
    wb.define_name("Data!Sales", "=Data!$A$1:$A$9")

    assert_equal [["Sales", "0", "Data!$A$1:$A$9"], ["Rate", nil, "0.96"]].sort, defined_names(wb).sort
  end

  def test_define_name_rejects_invalid_names
    wb = FastXlsx::Workbook.new
    ["has space", "A1", "!x"].each do |name|
      assert_raises(FastXlsx::Error, name) { wb.define_name(name, "=1") }
    end
  end

  # Both depend on what else the workbook holds by the time it is saved.
  def test_duplicate_or_unknown_sheet_names_raise_when_saving
    wb = FastXlsx::Workbook.new
    wb.add_worksheet
    wb.define_name("Rate", "=1").define_name("rate", "=2")
    assert_raises(FastXlsx::Error) { wb.to_xlsx }

    wb = FastXlsx::Workbook.new
    wb.add_worksheet
    wb.define_name("Nope!Rate", "=1")
    assert_raises(FastXlsx::Error) { wb.to_xlsx }
  end

  def test_properties
    wb = FastXlsx::Workbook.new
    wb.properties(title: "Q3 report", author: "Zac", keywords: "Confidential", company: "Acme")
    wb.add_worksheet

    zip = Zip::File.open_buffer(StringIO.new(wb.to_xlsx))
    core = Nokogiri::XML(zip.read("docProps/core.xml")).remove_namespaces!
    app = Nokogiri::XML(zip.read("docProps/app.xml")).remove_namespaces!
    assert_equal ["Q3 report", "Zac", "Confidential"], %w[title creator keywords].map { |t| core.at(t).text }
    assert_equal "Acme", app.at("Company").text
  end

  def test_properties_rejects_unknown_fields
    error = assert_raises(ArgumentError) { FastXlsx::Workbook.new.properties(titel: "typo") }
    assert_includes error.message, "titel"
  end

  def test_showcase_example_builds_one_sheet_per_feature
    Dir.mktmpdir do |dir|
      path = File.join(dir, "showcase.xlsx")
      example = File.expand_path("../examples/showcase.rb", __dir__)
      assert system(RbConfig.ruby, "-I", File.expand_path("../lib", __dir__), example, path, out: File::NULL),
             "showcase.rb failed"

      assert_equal %w[Values Formats Layout Conditional Validation Table Chart Media Printing],
                   Roo::Excelx.new(path).sheets
    end
  end

  # Precompiled (platform) gems ship one binary per Ruby version under
  # lib/fast_xlsx/<major.minor>/, not lib/fast_xlsx/fast_xlsx.bundle.
  def test_loads_the_extension_from_a_precompiled_gem_layout
    lib = File.expand_path("../lib", __dir__)
    binary = Dir[File.join(lib, "fast_xlsx", "fast_xlsx.{bundle,so,dll}")].first
    skip "compile the extension first" unless binary

    Dir.mktmpdir do |dir|
      FileUtils.cp_r(Dir[File.join(lib, "*")], dir)
      ruby_dir = File.join(dir, "fast_xlsx", RUBY_VERSION[/\d+\.\d+/])
      FileUtils.mkdir_p(ruby_dir)
      FileUtils.mv(File.join(dir, "fast_xlsx", File.basename(binary)), ruby_dir)

      script = 'require "fast_xlsx"; print FastXlsx::Workbook.new.tap { |w| w.add_worksheet << [1] }.to_xlsx[0, 2]'
      # Only `dir` may be visible: no Bundler/RubyGems paths back to the repo's lib.
      env = { "RUBYOPT" => nil, "RUBYLIB" => nil }
      output = IO.popen([env, RbConfig.ruby, "--disable-gems", "-I", dir, "-e", script], err: File::NULL, &:read)
      assert_equal "PK", output
    end
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

  def test_save_accepts_a_pathname
    Dir.mktmpdir do |dir|
      path = Pathname.new(dir).join("out.xlsx")
      wb = FastXlsx::Workbook.new
      wb.add_worksheet << ["saved"]
      wb.save(path)

      assert_equal "saved", Roo::Excelx.new(path.to_s).cell(1, 1)
    end
  end
end
