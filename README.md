# FastXlsx

[![Gem Version](https://badge.fury.io/rb/fast_xlsx.svg)](https://badge.fury.io/rb/fast_xlsx)
[![codecov](https://codecov.io/gh/7a6163/fast_xlsx/graph/badge.svg)](https://codecov.io/gh/7a6163/fast_xlsx)

Fast `.xlsx` writer for Ruby, built on [rust_xlsxwriter](https://github.com/jmcnamara/rust_xlsxwriter) via [magnus](https://github.com/matsadler/magnus).

> **Status: release candidate.** The API is frozen as of 0.8: 0.8.x releases only fix bugs, and 1.0 follows once it has been in real use without needing changes.

**Roadmap:** [1.0.0](https://github.com/7a6163/fast_xlsx/milestone/1) (release plan) · [1.x](https://github.com/7a6163/fast_xlsx/milestone/2) (planned features)

**Contents:** [Installation](#installation) · [Getting started](#getting-started) · [Guide](#guide) · [Performance](#performance) · [Support and versioning](#support-and-versioning) · [Development](#development)

## Installation

```bash
bundle add fast_xlsx
```

Requires CRuby 3.3 or later. Precompiled gems cover Linux, macOS and Windows; see [Support and versioning](#support-and-versioning).

## Getting started

A sales report, step by step: a header, one row per sale, a totals row, then sized columns, a frozen header and filter buttons.

```ruby
require "date"
require "fast_xlsx"

# Sample data; in an app it comes from your database.
Sale = Struct.new(:region, :product, :units, :price, :sold_on)
sales = [
  Sale.new("North", "Widget", 120, 9.5, Date.new(2026, 7, 3)),
  Sale.new("North", "Gadget", 45, 24.0, Date.new(2026, 7, 9)),
  Sale.new("South", "Widget", 80, 9.5, Date.new(2026, 7, 14)),
  Sale.new("South", "Gizmo", 12, 120.0, Date.new(2026, 7, 21))
]

# 1. A workbook with one worksheet.
wb = FastXlsx::Workbook.new
ws = wb.add_worksheet("Sales")

# 2. A header row, bold on a light blue fill.
header = FastXlsx::Format.new(bold: true, bg_color: "#DDEBF7", border_bottom: :thin)
ws.append(["Region", "Product", "Units", "Price", "Sold on", "Revenue"], format: header)

# 3. One row per sale. Dates show as yyyy-mm-dd on their own; a formula is a
#    FastXlsx::Formula. In formulas rows count from 1, so the first sale is row 2.
money = FastXlsx::Format.new(num_format: "#,##0.00")
sales.each.with_index(2) do |sale, row|
  ws.append([sale.region, sale.product, sale.units, sale.price, sale.sold_on,
             FastXlsx::Formula.new("C#{row}*D#{row}")],
            format: [nil, nil, nil, money, nil, money])
end

# 4. A totals row: the header's look, with the money format added.
last = sales.size + 1
ws.append(["Total", nil, FastXlsx::Formula.new("SUM(C2:C#{last})"), nil, nil,
           FastXlsx::Formula.new("SUM(F2:F#{last})")],
          format: [header, header, header, header, header, header.merge(num_format: "#,##0.00")])

# 5. Easier to read: columns sized to fit (set by hand where autofit can't tell), the header kept in view while
#    scrolling, and filter buttons on the header (rows and columns count from 0).
ws.autofit
ws.column_width(3..5, 12) # money, dates, formulas: autofit measures raw values
ws.freeze_panes("A2")
ws.autofilter("A1:F#{last}")

# 6. Save it. (In Rails: send_data wb.to_xlsx, filename: "sales.xlsx", ...)
wb.save("sales.xlsx")
```

- **Rows and columns count from 0** in method calls (`write(0, 0, ...)` is A1); Excel references like `"A2"` and formulas count from 1 as in Excel.
- **Formulas** are calculated by Excel when it opens the file.
- **Large exports** (tens of thousands of rows and up): use `FastXlsx::Workbook.new(memory: :constant)`. Rows go to disk as they are added, so memory stays flat; fill each worksheet from top to bottom. See [Memory modes](#memory-modes).
- **Rails**: build the workbook in the controller and `send_data wb.to_xlsx`; see [Rails](#rails).

The [Guide](#guide) below covers every feature; the [API docs](https://www.rubydoc.info/gems/fast_xlsx) list every method.

## Guide

[Writing cells](#writing-cells) · [Rails](#rails) · [Memory modes](#memory-modes) · [Formats](#formats) · [Columns and filters](#columns-and-filters) · [Sheet view and printing](#sheet-view-and-printing) · [Cell ranges](#cell-ranges) · [Layout](#layout) · [Outline groups](#outline-groups) · [Defined names](#defined-names) · [Protection](#protection) · [Conditional formats](#conditional-formats) · [Data validation](#data-validation) · [Comments](#comments) · [Images](#images) · [Tables](#tables) · [Charts](#charts) · [Errors](#errors)

### Writing cells

```ruby
require "fast_xlsx"

wb = FastXlsx::Workbook.new                       # or memory: :constant / :low, see below
ws = wb.add_worksheet("Report")                  # later: wb.worksheet("Report"), wb.worksheets

ws << ["id", "name", "created_at"]                # append a row
ws.concat(records.map { |r| [r.id, r.name, r.created_at] })  # append many rows in one call
ws.write(0, 5, 42)                                # write a single cell (row, col, value)
ws.write("F1", 42)                                # or by its Excel reference
ws.write("F1", 42, { bold: true })                # a format positionally, or as format: like append
ws["F1"] = 42                                     # the same (ws[0, 5] = 42 too)

wb.properties(title: "Q3 report", author: "Zac", keywords: "Confidential") # File > Info in Excel
wb.save("report.xlsx")                            # or wb.to_xlsx => binary String
```

`Workbook.new` and `add_worksheet` also take a block, which gets the new object:

```ruby
wb = FastXlsx::Workbook.new do |wb|
  wb.add_worksheet("Users") do |ws|
    ws.append(%w[id name email], format: { bold: true })
    User.find_each { |u| ws << u.attributes.values_at("id", "name", "email") } # a Hash's values in column order
  end
end
```

### Rails

```ruby
def export
  wb = FastXlsx::Workbook.new(memory: :constant) # large exports: memory stays flat
  ws = wb.add_worksheet("Orders")
  ws.append(["Order", "Customer", "Total", "Placed at"], format: { bold: true })
  Order.includes(:customer).find_each { |o| ws << [o.number, o.customer.name, o.total, o.created_at] }
  send_data wb.to_xlsx, filename: "orders.xlsx",
                        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
end
```

### Memory modes

By default every cell stays in memory until the file is saved. For large exports, two modes write each finished row to a temp file instead:

```ruby
FastXlsx::Workbook.new                          # memory: :standard (default): everything in memory, any write order
FastXlsx::Workbook.new(memory: :constant)       # rows on disk, strings stored inline in each cell
FastXlsx::Workbook.new(memory: :low)            # rows on disk, strings in Excel's shared string table
```

| | `:standard` (default) | `:constant` | `:low` |
|---|---|---|---|
| Finished rows | kept in memory | written to disk | written to disk |
| Memory grows with | all cells | nothing (flat) | the number of unique strings |
| Write order | any | top to bottom only | top to bottom only |
| `autofit` | full | only sees rows still in memory | only sees rows still in memory |
| Strings | shared string table | inline in each cell | shared string table |
| Output | standard | some readers (e.g. xsv) don't support inline strings | standard |

Which one:

- **`:standard`** for normal reports, when you need to go back and change earlier rows, or rely on `autofit`.
- **`:constant`** for large exports written row by row: memory stays flat whatever the data, and it is the fastest mode.
- **`:low`** for large exports that other programs will read: memory stays low when strings repeat (regions, statuses, …) and the file uses the standard shared string table. With many unique strings it keeps those strings in memory until `save`.

In both disk-backed modes, writing to a row that was already written to disk raises `FastXlsx::Error`, and tables must be added before their data (see [Tables](#tables)). An unknown mode raises `ArgumentError`.

### Formats

```ruby
header = FastXlsx::Format.new(bold: true, bg_color: "#DDEBF7", border_bottom: :thin, align: :center)
date   = FastXlsx::Format.new(num_format: "yyyy-mm-dd")

ws.append(["id", "name", "created_at"], format: header)  # format every cell in the row
ws.append([1, "a", Time.now], format: [nil, nil, date]) # or one format (or nil) per cell
ws.write(1, 2, Date.today, date)                         # format one cell
```

| Option | Values |
|---|---|
| `bold`, `italic`, `strikeout`, `text_wrap`, `shrink` | `true` / `false` |
| `underline` | `true` (single), `:single`, `:double`, `:single_accounting`, `:double_accounting` |
| `font_script` | `:superscript`, `:subscript` |
| `rotation` | degrees, `-90..90`, or `270` for stacked text |
| `indent` | indent level, e.g. `2` |
| `font_size` | number, e.g. `14` |
| `font_name` | e.g. `"Arial"` |
| `font_color`, `bg_color` | `"#RRGGBB"` or `0xRRGGBB` |
| `num_format` | Excel number format, e.g. `"#,##0.00"`, `"yyyy-mm-dd"` |
| `align` | `:left`, `:center`, `:right` |
| `valign` | `:top`, `:center`, `:bottom` |
| `border`, `border_left`, `border_right`, `border_top`, `border_bottom` | `:thin`, `:medium`, `:thick`, `:dashed`, `:dotted`, `:double`, `:hair` |
| `border_color` | `"#RRGGBB"` or `0xRRGGBB` |
| `locked` | `false` keeps the cell editable on a protected sheet (default `true`) |
| `hidden` | `true` hides the cell's formula on a protected sheet |

A Hash of options works wherever a format goes (equal Hashes share one format), and a format can be extended:

```ruby
ws.append(["Total", 1_234], format: { bold: true })
title = header.merge(font_size: 16)   # header's options plus font_size; header.to_h lists them
```

In a loop over many cells, create the `Format` once: building one from a Hash on every call is slower.

Per-side borders override `border`. Unknown options and invalid values raise `ArgumentError`.

Dates and times are shown as `yyyy-mm-dd` or `yyyy-mm-dd hh:mm:ss`. A format without a `num_format` (yours, a table column's or `column_format`'s) keeps its look and gets that date format added, so `append(row, format: bold)` gives bold dates; a format with a `num_format` is used as is.

### Columns and filters

```ruby
ws.column_width(0, 20)        # column A, width in characters
ws.column_width(1..3, 12)     # columns B–D
ws.column_format(4, FastXlsx::Format.new(num_format: "#,##0.00")) # default for cells in E written without a format
ws.autofit                        # size other columns to the data written so far; set widths are kept
ws.autofilter("A1:D101")          # filter buttons on A1:D101
```

### Sheet view and printing

```ruby
ws.zoom(150)                # 10..400 percent
ws.tab_color("#C00000")
ws.hide_gridlines
ws.activate                 # Excel opens on this sheet (un-hides it if hidden)
ws.selection("B2:C3")       # the cells selected when it opens
ws.top_left_cell("A50")     # and the cell scrolled to the top left
ws.ignore_error("A2:A1000", :number_stored_as_text) # no green triangles (one kind per range; don't overlap ranges)
other.hide                  # the sheet Excel opens on can't be hidden: activate another one first

ws.page_setup(landscape: true, paper: :a4,  # or :letter, :legal, :tabloid, :a3, :a5, Excel's paper number, 0 = printer default
              fit_width: 1,                 # 1 page wide, as many pages tall as needed
              repeat_rows: 0,               # print the header row on every page (an index or a Range)
              print_area: "A1:D100",        # any cell range style
              gridlines: true)              # print the gridlines
```

### Cell ranges

`autofilter`, `merge_range`, `conditional_format`, `data_validation` and `add_table` take a range in any of these styles:

```ruby
ws.autofilter(0, 0, 100, 3)   # four 0-based numbers: first_row, first_col, last_row, last_col
ws.autofilter("A1:D101")      # an Excel reference; "$A$1:$D$101" and a single "B2" work too
ws.autofilter(0..100, 0..3)   # rows and columns, each an Integer or a Range
ws.merge_range(0, 0..3, "Q3 report", title) # row 1, columns A–D
```

Methods that take one cell (`write`, `write_comment`, `insert_image`, `insert_chart`, `freeze_panes`) take `(row, col)` or a reference like `"B2"`.

`autofit` only sees rows still in memory, so in `:constant` / `:low` memory mode it ignores rows already written to disk; set widths with `column_width` instead.

### Layout

```ruby
ws.freeze_panes(1, 0)                          # keep the first row visible while scrolling
ws.row_height(0, 30)                       # row 1, height in points
ws.default_row_height(18)                  # every other row; call it before row_format, hide_rows, group_rows
ws.row_format(0, header)                   # default for cells in row 1 written without a format
ws.hide_rows(10..20)                       # also hide_columns(3), an index or a Range
ws.merge_range(0, 0, 0, 3, "Q3 report", title) # merge A1:D1; the value can be any cell type
ws.page_breaks([50, 100])                  # print a new page before rows 51 and 101
ws.vertical_page_breaks([8])               # and before column I
ws.page_header("&CPage &P of &N")               # printed header, Excel header/footer codes
ws.page_footer("&L&A", margin: 0.2)             # sheet name on the left; margin in inches
ws.margins(left: 0.5, top: 1)              # other margins keep Excel's defaults
```

### Outline groups

```ruby
ws.group_rows(1..10)                     # rows 2–11 get an expand/collapse button
ws.group_rows(1..4)                      # grouping again nests them (up to 7 levels)
ws.group_columns(2..3, collapsed: true)  # columns C–D, collapsed until expanded
```

`group_rows` needs `memory: :standard`: in `:constant` / `:low` mode rust_xlsxwriter writes rows without their outline level, so it raises. `group_columns` works in every mode.

### Defined names

```ruby
wb.define_name("Rate", "=0.96")                      # workbook-wide; use as =A1*Rate
wb.define_name("Report!Sales", "=Report!$B$2:$B$13") # only on the Report sheet
```

Invalid names raise `FastXlsx::Error` right away; duplicate names, and names for a sheet that doesn't exist, raise when saving.

### Protection

```ruby
input = FastXlsx::Format.new(locked: false)
ws.write(1, 1, 0, input)                                  # B2 stays editable
ws.protect                                                # lock everything else
ws.protect(password: "secret", allow: %i[sort use_autofilter]) # or with a password and allowed actions
ws.unprotect_range("B2:D10", name: "Inputs")               # a range users can still edit (optional password:)
```

`allow:` takes `:format_cells`, `:format_columns`, `:format_rows`, `:insert_columns`, `:insert_rows`, `:insert_links`, `:delete_columns`, `:delete_rows`, `:sort`, `:use_autofilter`, `:use_pivot_tables`, `:edit_scenarios`, `:edit_objects`. The password only stops editing in Excel; it does not encrypt the file.

### Conditional formats

```ruby
red = FastXlsx::Format.new(font_color: "#9C0006", bg_color: "#FFC7CE")

# rows 1–100 of column B (first_row, first_col, last_row, last_col)
ws.conditional_format(0, 1, 99, 1, type: :cell, criteria: :<, value: 0, format: red)
ws.conditional_format(0, 1, 99, 1, type: :cell, criteria: :between, value: [1, 10], format: red)
ws.conditional_format(0, 0, 99, 0, type: :text, criteria: :contains, value: "error", format: red)
ws.conditional_format(0, 0, 99, 3, type: :formula, value: "=$D1>100", format: red)
ws.conditional_format(0, 2, 99, 2, type: :data_bar)
ws.conditional_format(0, 2, 99, 2, type: :color_scale)            # 3-color; colors: 2 for 2-color
```

| `type` | `criteria` | `value` |
|---|---|---|
| `:cell` | `:==`, `:!=`, `:>`, `:>=`, `:<`, `:<=`, `:between`, `:not_between` | number or string; `[min, max]` for the range criteria |
| `:text` | `:contains`, `:not_contains`, `:begins_with`, `:ends_with` | string |
| `:formula` | — | formula string, relative to the top-left cell |
| `:data_bar`, `:color_scale` | — | — |

### Data validation

```ruby
ws.data_validation(1, 2, 100, 2, type: :list, value: %w[Open Closed])     # dropdown in C2:C101
ws.data_validation(1, 2, 100, 2, type: :list, value: "=$Z$1:$Z$10")       # dropdown from a range
ws.data_validation(1, 3, 100, 3, type: :whole_number, criteria: :between, value: [1, 10],
                   input_title: "Quantity", input_message: "1 to 10",
                   error_title: "Invalid", error_message: "Enter a whole number from 1 to 10")
```

`type` is `:list`, `:whole_number`, `:decimal` or `:text_length`; the number types take the same `criteria` as `:cell` conditional formats.

### Comments

```ruby
ws.write_comment(0, 0, "Checked by finance", author: "Zac")   # Excel shows it as a note on A1
```

### Images

```ruby
ws.insert_image(0, 0, "logo.png")                                  # top-left corner in A1
ws.insert_image(0, 5, StringIO.new(blob.download), scale: 0.5, x_offset: 10, y_offset: 4, alt_text: "Logo")
ws.insert_image(10, 0, "chart.png", width: 320, height: 180)       # pixel size; one of them keeps the aspect ratio
```

A String is always treated as a path, so wrap raw bytes (such as Active Storage's `blob.download`) in a `StringIO`.

The source is a file path or any IO responding to `#read` (PNG, JPEG, GIF or BMP). Offsets are in pixels. Data that is not a supported image raises `FastXlsx::Error`.

### Tables

```ruby
ws.concat([%w[Region Rep Sales], *sales])            # header row, then the data
ws.add_table(0, 0, sales.size + 1, 2, total_row: true, style: :medium2, # +1 row for the totals
             columns: [{ header: "Region", total_label: "Total" }, "Rep", { header: "Sales", total: :sum }])
```

The range includes the header row and, with `total_row: true`, the total row; the table writes the headers. `columns` must match the range width. Options: `style` (`:light1`–`:light21`, `:medium1`–`:medium28`, `:dark1`–`:dark11`, `:none`), `name`, `total_row`, `banded_rows`, `autofilter`. Column totals: `:sum`, `:average`, `:count`, `:count_numbers`, `:max`, `:min`, `:std_dev`, `:var`.

You can also add the table first and then append the data: after `add_table`, `<<` / `append` / `concat` continue right under the header row. In `:constant` / `:low` memory mode this is the only order that works; adding a table whose header row was already written to disk raises `FastXlsx::Error`.

### Charts

```ruby
ws.concat([%w[Month Sales Costs], ["Jan", 10, 7], ["Feb", 25, 12], ["Mar", 18, 11]])

ws.insert_chart(1, 4, type: :column,
                series: [
                  { name: "Sales", categories: "Sheet1!$A$2:$A$4", values: "Sheet1!$B$2:$B$4" },
                  { name: "Costs", categories: "Sheet1!$A$2:$A$4", values: "Sheet1!$C$2:$C$4" }
                ],
                title: "Q1", x_axis: "Month", y_axis: "Amount", width: 600, height: 360)
```

`type`: `:column`, `:column_stacked`, `:bar`, `:bar_stacked`, `:line`, `:line_stacked`, `:area`, `:area_stacked`, `:pie`, `:doughnut`, `:radar`, `:scatter`. Ranges use Excel syntax, so a chart can plot data from another worksheet. Sizes are in pixels (default 480 × 288).

Values are mapped by type:

| Ruby | Excel |
|---|---|
| `Integer`, `Float`, any `Numeric` | number |
| `String` | string |
| `Time`, `DateTime` | date and time (`yyyy-mm-dd hh:mm:ss` unless formatted), in its own offset |
| `Date` | date (`yyyy-mm-dd` unless formatted) |
| `FastXlsx::Formula.new("SUM(A1:A9)")` | formula |
| `FastXlsx::URL.new("https://…")`, `URL.new(url, text: "Title")` | hyperlink (optionally showing other text) |
| `FastXlsx::RichString.new(["Total: ", bold], "1,234")` | text with a format per segment |
| `true` / `false` | boolean |
| `nil` | empty cell (a blank cell with the format, when one is given) |
| anything else | `to_s` as string |

`<<` and `concat` append after the last row written to that worksheet. In `:constant` / `:low` memory mode rows are written to disk as you go, so fill each worksheet top to bottom.

### Errors

| Error | Raised for |
|---|---|
| `TypeError` | an argument of the wrong type, e.g. a String where a row number goes |
| `RangeError` | a row or column outside the sheet (1,048,576 × 16,384), or a negative count (an infinite Float raises `FloatDomainError`, a `RangeError`) |
| `ArgumentError` | the right type but a value that isn't allowed: an unknown option or symbol, an invalid color or cell reference, a reversed range or single-cell merge, a size or margin that is negative, NaN or beyond Excel's limits |
| `FastXlsx::Error` | what the workbook can't do: duplicate or invalid names, writing to rows already on disk, overlapping merges, hiding the sheet Excel opens on, text over Excel's limits (cells, URLs, headers and footers), unsupported images |

Numbers are converted like Ruby's own methods do (`to_int`), so `1.9` as a row is row `1`.

## Performance

Apple Silicon, Ruby 4.0.5. Each library uses its own idiomatic row-append API; xlsxtream is a streaming writer with fewer features.

### Speed

20,000 rows × 5 columns (integer, string, integer, `Time`, float), build + serialize to a String, median of 7 runs (3 for rubyXL):

| Library | Time | vs fastest | Ruby objects allocated |
|---|---:|---:|---:|
| **fast_xlsx** (`memory: :constant`) | **77 ms** | 1.0x | 7 |
| **fast_xlsx** (`memory: :low`) | **86 ms** | 1.1x | 7 |
| **fast_xlsx** | **89 ms** | 1.1x | 10 |
| [xlsxtream](https://github.com/felixbuenemann/xlsxtream) 3.1 | 183 ms | 2.4x | 561,728 |
| [fast_excel](https://github.com/Paxa/fast_excel) 0.5 (constant_memory) | 200 ms | 2.6x | 20,079 |
| [fast_excel](https://github.com/Paxa/fast_excel) 0.5 | 238 ms | 3.1x | 320,076 |
| [write_xlsx](https://github.com/cxn03651/write_xlsx) 1.15 | 608 ms | 7.8x | 1,483,899 |
| [caxlsx](https://github.com/caxlsx/caxlsx) 4.5 | 689 ms | 8.9x | 745,122 |
| [rubyXL](https://github.com/weshatheleopard/rubyXL) 3.4 | 2755 ms | 35.6x | 8,700,448 |

All outputs are 689–750 KB.

### Memory

200,000 rows × 5 columns saved to a file; extra peak RSS over a process that only builds the data, median of 3 runs. "Unique" gives every row a different 100-character string; "repeated" uses a handful of values (regions, statuses), as most reports do:

| Library | Unique strings | Repeated strings |
|---|---:|---:|
| **fast_xlsx** (`memory: :constant`) | **+1 MB** | **+1 MB** |
| **fast_xlsx** (`memory: :low`) | +61 MB | **+1 MB** |
| **fast_xlsx** | +271 MB | +217 MB |
| fast_excel 0.5 (constant_memory) | +10 MB | +9 MB |
| fast_excel 0.5 | +182 MB | +151 MB |

The `:standard` mode uses more memory than fast_excel's: when saving, rust_xlsxwriter assembles each worksheet's XML in memory (so several worksheets can be built in parallel) instead of streaming it from a temp file. Use `memory: :constant` or `memory: :low` for large exports.

### Threads

Most of the time goes into saving (building the XML and compressing it). `to_xlsx` and `save` do that without holding Ruby's global lock, so in a threaded server (Puma, Sidekiq) other threads keep running while a large export is saved. An interrupt (Ctrl-C, `Timeout`, `Thread#raise`) takes effect once the save finishes.

### Reproduce

```bash
bundle exec rake compile
BUNDLE_GEMFILE=bench/Gemfile bundle install
BUNDLE_GEMFILE=bench/Gemfile bundle exec ruby bench/compare.rb   # speed; optional row count argument

# memory: peak RSS of one run (use /usr/bin/time -v on Linux); subtract the baseline
BUNDLE_GEMFILE=bench/Gemfile /usr/bin/time -l bundle exec ruby bench/memory.rb baseline unique
BUNDLE_GEMFILE=bench/Gemfile /usr/bin/time -l bundle exec ruby bench/memory.rb fast_xlsx:low unique
```

## Support and versioning

- **Ruby:** CRuby 3.3 and later. A Ruby version is dropped only after it reaches its end of life, and only in a minor release. JRuby and TruffleRuby are not supported (this is a native extension).
- **Precompiled gems:** Linux (x86_64 and aarch64, glibc and musl; ARM musl), macOS (arm64 and x86_64) and Windows (x64). Before each release the Linux x86_64, macOS arm64 and Windows gems are installed and loaded on Ruby 3.3, 3.4 and 4.0. Other platforms build from source and need a Rust toolchain.
- **Versions:** from 1.0, [Semantic Versioning](https://semver.org): no breaking API change before 2.0. Until then a minor release (0.x.0) can change behaviour; the [changelog](CHANGELOG.md) lists every change under "Changed".
- **Deprecations:** from 1.0, a method or option to be removed first prints a warning (with `warn`, so on by default) for at least one minor release, with the CHANGELOG naming its replacement; it goes only in the next major version.
- **Output:** `.xlsx` files that open in Excel 2007 and later, LibreOffice and Google Sheets. Updating rust_xlsxwriter can change the bytes of the file but not what it contains.

## Development

```bash
bin/setup
bundle exec rake compile   # build the Rust extension into lib/fast_xlsx/
bundle exec rake test
bundle exec rake           # compile + test + rubocop
```

Coverage of the Rust extension while the Ruby tests run (needs [cargo-llvm-cov](https://github.com/taiki-e/cargo-llvm-cov) and `rustup component add llvm-tools-preview`):

```bash
rm -rf tmp lib/fast_xlsx/fast_xlsx.bundle           # force an instrumented rebuild
eval "$(cargo llvm-cov show-env --export-prefix)"
bundle exec rake compile test
cargo llvm-cov report --release                     # or --lcov / --html
```

Rebuild in a clean shell afterwards (`rm -rf tmp && bundle exec rake compile`) so the everyday build is not instrumented.

Mutation tests ([cargo-mutants](https://mutants.rs)) change the Rust source one small edit at a time and check that a Ruby test fails for each change. `ext/fast_xlsx/tests/ruby_suite.rs` is what connects the two: it rebuilds the extension and runs the Ruby suite, so `cargo test` covers the Ruby tests too.

```bash
cargo install cargo-mutants
cargo mutants --file ext/fast_xlsx/src/lib.rs --jobs 4   # ~7 minutes; results in mutants.out/
```

Mutants listed in `mutants.out/missed.txt` point at behaviour no test checks. The `Mutation tests` workflow runs this weekly and on demand.

### Releasing

First generate the showcase workbook and open it in Excel to check every feature renders (the test suite only inspects the XML):

```bash
bundle exec rake compile
ruby -Ilib examples/showcase.rb showcase.xlsx
```

Then bump `FastXlsx::VERSION`, update `CHANGELOG.md`, commit, and push a matching tag:

```bash
git tag v0.1.0 && git push origin v0.1.0
```

The `Build gems` workflow builds the source gem plus precompiled gems for each platform and pushes them to RubyGems (trusted publishing) and GitHub Packages. It refuses to publish if the tag and `VERSION` differ.

## License

MIT
