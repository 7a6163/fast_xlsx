# Migrating from fast_excel

fast_xlsx started as a rewrite of [fast_excel](https://github.com/Paxa/fast_excel) on [rust_xlsxwriter](https://github.com/jmcnamara/rust_xlsxwriter), and keeps its spirit: append rows fast, with little memory. Most code moves over with a few renames. This guide lists them, then the differences in behaviour to check.

**Why switch**

- Faster writing (see the [README benchmark](../README.md#speed)), and `memory: :constant` adds about 1 MB however large the file.
- Precompiled gems for Linux, macOS and Windows: no C compiler or FFI at install.
- Dates show as dates without a format, every option is checked (a typo raises instead of being ignored), and saving doesn't block other threads.
- Conditional formats, data validation, tables, charts, images, comments, protection and printing options as Ruby methods.

**Before you start**

- fast_xlsx needs CRuby 3.3 or later (fast_excel supports 2.7+).
- Rows and columns are 0-based in both, and `<<` appends in both.

## A report, before and after

```ruby
# fast_excel
workbook = FastExcel.open(constant_memory: true)
worksheet = workbook.add_worksheet("Orders")
bold = workbook.bold_format
worksheet.append_row(["Order", "Total", "Placed at"], bold)
money = workbook.number_format("#,##0.00")
orders.each { |o| worksheet.append_row([o.number, o.total, o.created_at], [nil, money, nil]) }
worksheet.set_column_width(0, 20)
send_data workbook.read_string, filename: "orders.xlsx"
```

```ruby
# fast_xlsx
workbook = FastXlsx::Workbook.new(memory: :constant)
worksheet = workbook.add_worksheet("Orders")
worksheet.append(["Order", "Total", "Placed at"], format: { bold: true })
money = FastXlsx::Format.new(num_format: "#,##0.00")
orders.each { |o| worksheet.append([o.number, o.total, o.created_at], format: [nil, money, nil]) }
worksheet.column_width(0, 20)
send_data workbook.to_xlsx, filename: "orders.xlsx"
```

`created_at` now shows as `yyyy-mm-dd hh:mm:ss` without a format; with fast_excel it was a number unless you passed a date format.

## Method by method

### Workbook

| fast_excel | fast_xlsx |
|---|---|
| `FastExcel.open` | `FastXlsx::Workbook.new` |
| `FastExcel.open(constant_memory: true)` | `FastXlsx::Workbook.new(memory: :constant)` (or `:low`, which keeps Excel's shared strings) |
| `FastExcel.open("report.xlsx")` … `workbook.close` | `FastXlsx::Workbook.new` … `workbook.save("report.xlsx")` |
| `workbook.read_string` | `workbook.to_xlsx` |
| `workbook.remove_tmp_folder` | not needed: no temp file is created |
| `workbook.add_worksheet(name)` | `workbook.add_worksheet(name)` |
| `workbook.get_worksheet_by_name(name)` | `workbook.worksheet(name)` |
| `workbook.add_format(bold: true)` | `FastXlsx::Format.new(bold: true)`, or pass `{ bold: true }` where a format goes |
| `workbook.bold_format` | `FastXlsx::Format.new(bold: true)` |
| `workbook.number_format("0.00")` | `FastXlsx::Format.new(num_format: "0.00")` |

### Writing

| fast_excel | fast_xlsx |
|---|---|
| `worksheet << values` / `append_row(values, format)` | `worksheet << values` / `append(values, format: format)` |
| `write_value(row, col, value, format)` | `write(row, col, value, format)`, or `write("B2", value, format)`, or `worksheet["B2"] = value` |
| `write_number`, `write_string`, `write_datetime`, `write_formula`, `write_url`, `write_boolean` | `write`: the type comes from the value |
| `write_row(row, values, formats)` | `values.each_with_index { \|v, col\| worksheet.write(row, col, v, format) }`, or `append` when it is the next row |
| `FastExcel::Formula.new("SUM(A1:A9)")` | `FastXlsx::Formula.new("SUM(A1:A9)")` |
| `FastExcel::URL.new(url)` | `FastXlsx::URL.new(url)`, and `URL.new(url, text: "shown")` |
| `FastExcel.date_num(time, offset)` with a date format | `time` as it is: dates are converted and formatted for you |
| `worksheet.last_row_number` | `worksheet.next_row - 1` |

### Columns, rows and filters

| fast_excel | fast_xlsx |
|---|---|
| `set_column_width(col, width)` | `column_width(col, width)` |
| `set_columns_width(first, last, width)` | `column_width(first..last, width)` |
| `set_column(first, last, width, format)` | `column_width(first..last, width)` and `column_format(first..last, format)` |
| `set_row(row, height, format)` | `row_height(row, height)` and `row_format(row, format)` |
| `worksheet.auto_width = true` (before writing) | `worksheet.autofit` (after writing; see below) |
| `enable_filters!(end_col: 3)` | `autofilter(0, 0, worksheet.next_row - 1, 3)`, or `autofilter("A1:D100")` |
| `freeze_panes`, `merge_range` and other libxlsxwriter functions | Ruby methods with the same idea: `freeze_panes("A2")`, `merge_range("A1:D1", "Title")`; see the [Guide](../README.md#guide) |

### Format options

| fast_excel | fast_xlsx |
|---|---|
| `bold:`, `italic:`, `text_wrap:`, `shrink:`, `font_size:`, `font_name:`, `num_format:`, `rotation:`, `indent:`, `bg_color:` | the same |
| `font_family:` | `font_name:` |
| `font_strikeout: true` | `strikeout: true` |
| `underline: :underline_single` (`_double`, `_single_accounting`, `_double_accounting`) | `underline: true` or `:single` (`:double`, `:single_accounting`, `:double_accounting`) |
| `font_script: :font_subscript` / `:font_superscript` | `font_script: :subscript` / `:superscript` |
| `font_color: :orange`, `"#FF0000"`, `0xFF0000` | `"#FF0000"` or `0xFF0000`: color names aren't supported, use their hex value |
| `border: :border_thin` | `border: :thin` (also `:medium`, `:thick`, `:dashed`, `:dotted`, `:double`, `:hair`) |
| `left: :medium`, `top:`, `right:`, `bottom:` | `border_left: :medium`, `border_top:`, `border_right:`, `border_bottom:` |
| `left_color:`, `top_color:` … per side | `border_color:`, one color for every side |
| `align: { h: :align_center, v: :align_vertical_center }` | `align: :center, valign: :center` |

## Differences in behaviour

- **Dates have a format.** fast_excel wrote `Time` and `Date` as plain numbers unless you passed a date format. fast_xlsx shows them as `yyyy-mm-dd` (`Date`) or `yyyy-mm-dd hh:mm:ss` (`Time`, `DateTime`), and adds that to a format without a `num_format` (a bold row stays bold). Pass your own `num_format` to keep a different one. Both use the value's own UTC offset, so times don't shift. Dates before 1900-03-01 now show correctly; fast_excel's showed one day late, because Excel counts a 1900-02-29 that never existed.
- **Errors raise.** Unknown or misspelled options, invalid colors and values Excel can't hold (a string over 32,767 characters, a row past 1,048,576) raise `ArgumentError`, `RangeError` or `FastXlsx::Error`; see [Errors](../README.md#errors). In `:constant` mode, writing to a row already on disk raises `FastXlsx::Error` (fast_excel raised `ArgumentError`).
- **Autofit runs after writing.** fast_excel measured text as you wrote it (strings only). `autofit` measures what is in memory when you call it, so in `:constant` / `:low` mode it only sees the last rows: set widths with `column_width` there. It measures raw values, so give formatted numbers, dates and formulas a width by hand.
- **No temp files.** fast_excel always wrote to a file, a temporary one without a filename. fast_xlsx builds the workbook in memory (`:standard`) or in its own temp files (`:constant` / `:low`) and gives you the result with `to_xlsx` or `save`.
- **Formats are values.** A `FastXlsx::Format` is frozen and created on its own, not from a workbook, so one format can serve several workbooks. Use `format.merge(italic: true)` for a variant.

## Not available

- **The libxlsxwriter functions under `Libxlsxwriter.*`.** fast_xlsx has no raw binding; the [Guide](../README.md#guide) and the [API docs](https://www.rubydoc.info/gems/fast_xlsx) cover what is exposed, and the [roadmap](https://github.com/7a6163/fast_xlsx/milestone/2) what is planned.
- **`FastExcel.open(default_format: ...)`.** Use `column_format` / `row_format`, or pass the format when writing.
- **Some format options:** color names, per-side border colors, `font_outline`, `font_shadow`, `text_justlast`, and alignments other than left, center and right (horizontal) or top, center and bottom (vertical).
- **Reading or editing existing files**: neither library does this.

## Checklist

1. Replace `gem "fast_excel"` with `gem "fast_xlsx"`, and check you are on Ruby 3.3+.
2. Swap `FastExcel.open` / `read_string` / `close` for `FastXlsx::Workbook.new` / `to_xlsx` / `save`.
3. Rename the methods and format options from the tables above; run your code: a missed rename raises `NoMethodError` or `ArgumentError` rather than being ignored.
4. Drop `FastExcel.date_num` and date formats you only added to make dates readable.
5. Replace `auto_width = true` with `autofit` after writing (or `column_width` in `:constant` mode).
6. Open one generated file in Excel and compare.
