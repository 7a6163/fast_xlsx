# frozen_string_literal: true

# Builds one workbook with a worksheet per feature, to check the output in
# Excel (or LibreOffice / Numbers) before a release.
#
#   bundle exec rake compile
#   ruby -Ilib examples/showcase.rb [showcase.xlsx]
require "date"
require "stringio"
require "zlib"
require "fast_xlsx"

path = ARGV[0] || "showcase.xlsx"

# A small gradient PNG, generated so the example needs no image file.
def gradient_png(width, height)
  rows = (0...height).map do |y|
    "\0".b + (0...width).map { |x| [x * 255 / width, y * 255 / height, 200].pack("C3") }.join
  end
  chunk = ->(type, data) { [data.bytesize].pack("N") + type + data + [Zlib.crc32(type + data)].pack("N") }
  "\x89PNG\r\n\x1A\n".b +
    chunk.call("IHDR", [width, height, 8, 2, 0, 0, 0].pack("NNC5")) +
    chunk.call("IDAT", Zlib::Deflate.deflate(rows.join)) +
    chunk.call("IEND", "".b)
end

wb = FastXlsx::Workbook.new
wb.properties(title: "fast_xlsx showcase", author: "fast_xlsx", keywords: "example")

bold = FastXlsx::Format.new(bold: true)
header = FastXlsx::Format.new(bold: true, bg_color: "#DDEBF7", border_bottom: :thin, align: :center)
money = FastXlsx::Format.new(num_format: "#,##0.00")
date = FastXlsx::Format.new(num_format: "yyyy-mm-dd")
datetime = FastXlsx::Format.new(num_format: "yyyy-mm-dd hh:mm")

# Values: every cell type.
ws = wb.add_worksheet("Values")
ws.column_width(0, 22).column_width(1, 40)
ws.append(%w[Type Value], format: header)
ws << ["Integer", 42]
ws << ["Float", 3.14159]
ws.append(["Money format", 1_234_567.891], format: [nil, money])
ws.append(["Date", Date.new(2024, 2, 29)], format: [nil, date])
ws.append(["Time", Time.new(2024, 2, 29, 13, 45)], format: [nil, datetime])
ws << ["Boolean", true]
ws << ["nil (empty cell)", nil]
ws << ["Formula =SUM(B2:B3)", FastXlsx::Formula.new("SUM(B2:B3)")]
ws << ["URL", FastXlsx::URL.new("https://github.com/7a6163/fast_xlsx")]
ws << ["URL with text", FastXlsx::URL.new("https://github.com/7a6163/fast_xlsx", text: "fast_xlsx on GitHub")]
ws << ["Rich string",
       FastXlsx::RichString.new(["Bold ", bold], "plain ", ["italic", FastXlsx::Format.new(italic: true)])]
ws << ["Symbol (to_s)", :symbol]

# Formats: one option per row.
ws = wb.add_worksheet("Formats")
ws.column_width(0, 24).column_width(1, 30)
ws.append(%w[Option Sample], format: header)
{
  "bold" => { bold: true }, "italic" => { italic: true }, "strikeout" => { strikeout: true },
  "underline :double" => { underline: :double }, "font_script :superscript" => { font_script: :superscript },
  "font_size 16" => { font_size: 16 }, "font_name Courier New" => { font_name: "Courier New" },
  "font_color #C00000" => { font_color: "#C00000" }, "bg_color #FFF2CC" => { bg_color: "#FFF2CC" },
  "align :right" => { align: :right }, "valign :top" => { valign: :top }, "indent 2" => { indent: 2 },
  "rotation 45" => { rotation: 45 }, "text_wrap" => { text_wrap: true }, "shrink" => { shrink: true },
  "border :medium + color" => { border: :medium, border_color: "#2F5597" },
  "border_bottom :double" => { border_bottom: :double }
}.each do |label, options|
  text = label == "text_wrap" ? "long text that wraps inside the cell width" : "Sample text"
  ws.append([label, text], format: [nil, FastXlsx::Format.new(**options)])
end
ws.row_height(14, 40)

# Layout: widths, autofit, merged title, frozen header, filter.
ws = wb.add_worksheet("Layout")
ws.merge_range(0, 0, 0, 3, "Merged title across A1:D1", FastXlsx::Format.new(bold: true, font_size: 14, align: :center))
ws.append(%w[Region Rep Amount Note], format: header)
20.times { |i| ws << [%w[North South East West][i % 4], "Rep #{i + 1}", (i + 1) * 125.5, "row #{i + 1}"] }
ws.column_format(2, money)
ws.column_width(3, 30) # kept by autofit
ws.autofit
ws.freeze_panes(2, 0)
ws.autofilter(1, 0, 21, 3)

# Conditional formats.
ws = wb.add_worksheet("Conditional")
red = FastXlsx::Format.new(font_color: "#9C0006", bg_color: "#FFC7CE")
green = FastXlsx::Format.new(font_color: "#006100", bg_color: "#C6EFCE")
ws.append(["Cell < 0", "Text contains 'error'", "Data bar", "Color scale", "Formula (A > 50)"], format: header)
values = [-30, 80, 15, -5, 60, 95, 40, -10, 70, 25]
values.each_with_index { |v, i| ws << [v, i.even? ? "ok" : "error #{i}", v.abs, v, v] }
ws.conditional_format(1, 0, 10, 0, type: :cell, criteria: :<, value: 0, format: red)
ws.conditional_format(1, 1, 10, 1, type: :text, criteria: :contains, value: "error", format: red)
ws.conditional_format(1, 2, 10, 2, type: :data_bar)
ws.conditional_format(1, 3, 10, 3, type: :color_scale)
ws.conditional_format(1, 4, 10, 4, type: :formula, value: "=$A2>50", format: green)
ws.column_width(0..4, 22)

# Data validation.
ws = wb.add_worksheet("Validation")
ws.append(["Status (dropdown)", "Quantity 1-10", "Price >= 0", "Code <= 5 chars"], format: header)
ws.data_validation(1, 0, 20, 0, type: :list, value: %w[Open Pending Closed],
                                input_title: "Status", input_message: "Pick a status")
ws.data_validation(1, 1, 20, 1, type: :whole_number, criteria: :between, value: [1, 10],
                                error_title: "Invalid quantity", error_message: "Enter a whole number from 1 to 10")
ws.data_validation(1, 2, 20, 2, type: :decimal, criteria: :>=, value: 0)
ws.data_validation(1, 3, 20, 3, type: :text_length, criteria: :<=, value: 5)
ws.column_width(0..3, 20)

# Table with a total row.
ws = wb.add_worksheet("Table")
sales = [%w[North Ann 1200], %w[South Bob 950], %w[East Cai 1430], %w[West Dee 780]].map { |r, n, a| [r, n, a.to_i] }
ws.add_table(0, 0, sales.size + 1, 2, total_row: true, style: :medium2,
                                      columns: [{ header: "Region", total_label: "Total" }, "Rep",
                                                { header: "Sales", total: :sum, format: money }])
ws.concat(sales)
ws.column_width(0..2, 14)

# Charts.
ws = wb.add_worksheet("Chart")
ws.append(%w[Month Sales Costs], format: header)
ws.concat([["Jan", 10, 7], ["Feb", 25, 12], ["Mar", 18, 11], ["Apr", 30, 16], ["May", 27, 14]])
series = [{ name: "Sales", categories: "Chart!$A$2:$A$6", values: "Chart!$B$2:$B$6" },
          { name: "Costs", categories: "Chart!$A$2:$A$6", values: "Chart!$C$2:$C$6" }]
ws.insert_chart(0, 4, type: :column, series: series, title: "Column", x_axis: "Month", y_axis: "Amount")
ws.insert_chart(16, 4, type: :line, series: series, title: "Line", width: 600, height: 300)
ws.insert_chart(16, 0, type: :pie, series: [series.first], title: "Pie", width: 300, height: 300)

# Image and comment.
ws = wb.add_worksheet("Media")
ws << ["Generated PNG at scale 1, pixel size 240x60, and with an offset. B1 has a comment."]
ws.write_comment(0, 1, "This is a note (comment).", author: "fast_xlsx")
png = gradient_png(120, 40)
ws.insert_image(2, 0, StringIO.new(png), alt_text: "Gradient")
ws.insert_image(2, 3, StringIO.new(png), width: 240, height: 60)
ws.insert_image(8, 0, StringIO.new(png), x_offset: 20, y_offset: 10, scale: 1.5)

# Printing: check File > Print preview.
ws = wb.add_worksheet("Printing")
ws.page_header("&CPage &P of &N").page_footer("&L&A&R&D", margin: 0.2).margins(left: 0.5, right: 0.5)
ws.page_breaks([30])
ws.vertical_page_breaks([4])
ws.append(%w[Row A B C D E F], format: header)
60.times { |i| ws << [i + 1, *Array.new(6) { |c| (i + 1) * (c + 1) }] }

wb.save(path)
puts "wrote #{path}"
