# frozen_string_literal: true

$LOAD_PATH.unshift File.expand_path("../lib", __dir__)
require "fast_xlsx"

require "minitest/autorun"
require "roo"
require "stringio"
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

  # Raw worksheet XML, for settings roo does not expose (column widths, filters).
  def sheet_xml(workbook, index = 1)
    Zip::File.open_buffer(StringIO.new(workbook.to_xlsx)).read("xl/worksheets/sheet#{index}.xml")
  end

  # { column_number => width } from the <cols> element, 1-based like Excel.
  def column_widths(workbook)
    sheet_xml(workbook).scan(/<col min="(\d+)" max="(\d+)" width="([\d.]+)"/).each_with_object({}) do |(min, max, w), h|
      (min.to_i..max.to_i).each { |c| h[c] = w.to_f }
    end
  end

  # Resolved style of one cell (e.g. "A1") from xl/styles.xml, which roo does not expose.
  def cell_style(workbook, ref)
    zip = Zip::File.open_buffer(StringIO.new(workbook.to_xlsx))
    sheet = Nokogiri::XML(zip.read("xl/worksheets/sheet1.xml")).remove_namespaces!
    styles = Nokogiri::XML(zip.read("xl/styles.xml")).remove_namespaces!
    xf = styles.css("cellXfs > xf")[sheet.at("c[r='#{ref}']")["s"].to_i]
    style_hash(styles, xf)
  end

  private

  def style_hash(styles, cell_xf)
    font = styles.css("fonts > font")[cell_xf["fontId"].to_i]
    fill = styles.css("fills > fill")[cell_xf["fillId"].to_i]
    border = styles.css("borders > border")[cell_xf["borderId"].to_i]
    align = cell_xf.at("alignment")
    {
      font_size: font.at("sz")&.[]("val")&.to_f, font_name: font.at("name")&.[]("val"),
      font_color: font.at("color")&.[]("rgb"), bg_color: fill.at("fgColor")&.[]("rgb"),
      align: align&.[]("horizontal"), valign: align&.[]("vertical"), text_wrap: align&.[]("wrapText") == "1",
      border: %w[left right top bottom].to_h { |side| [side.to_sym, border.at(side)&.[]("style")] }
    }
  end
end
