# FastXlsx

[![codecov](https://codecov.io/gh/7a6163/fast_xlsx/graph/badge.svg)](https://codecov.io/gh/7a6163/fast_xlsx)

Fast `.xlsx` writer for Ruby, built on [rust_xlsxwriter](https://github.com/jmcnamara/rust_xlsxwriter) via [magnus](https://github.com/matsadler/magnus).

> **Status: early.** Column formats, freeze panes, merged cells and charts are not implemented yet.

## Usage

```ruby
require "fast_xlsx"

wb = FastXlsx::Workbook.new                       # or Workbook.new(constant_memory: true)
ws = wb.add_worksheet("Report")

ws << ["id", "name", "created_at"]                # append a row
ws.concat(records.map { |r| [r.id, r.name, r.created_at] })  # append many rows in one call
ws.write(0, 5, 42)                                # write a single cell (row, col, value)

wb.save("report.xlsx")                            # or wb.to_xlsx => binary String
```

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
| `bold`, `italic`, `underline`, `text_wrap` | `true` / `false` |
| `font_size` | number, e.g. `14` |
| `font_name` | e.g. `"Arial"` |
| `font_color`, `bg_color` | `"#RRGGBB"` or `0xRRGGBB` |
| `num_format` | Excel number format, e.g. `"#,##0.00"`, `"yyyy-mm-dd"` |
| `align` | `:left`, `:center`, `:right` |
| `valign` | `:top`, `:center`, `:bottom` |
| `border`, `border_left`, `border_right`, `border_top`, `border_bottom` | `:thin`, `:medium`, `:thick`, `:dashed`, `:dotted`, `:double`, `:hair` |

Per-side borders override `border`. Unknown options and invalid values raise `ArgumentError`.

### Columns and filters

```ruby
ws.set_column_width(0, 20)        # column A, width in characters
ws.set_column_width(1..3, 12)     # columns B–D
ws.set_column_format(4, FastXlsx::Format.new(num_format: "#,##0.00")) # default for cells in E written without a format
ws.autofit                        # size columns to the data written so far
ws.autofilter(0, 0, 100, 3)       # filter buttons on A1:D101 (first_row, first_col, last_row, last_col)
```

`autofit` only sees rows still in memory, so it has no effect on rows already flushed in `constant_memory` mode.

### Layout

```ruby
ws.freeze_panes(1, 0)                          # keep the first row visible while scrolling
ws.set_row_height(0, 30)                       # row 1, height in points
ws.merge_range(0, 0, 0, 3, "Q3 report", title) # merge A1:D1; the value can be any cell type
```

Values are mapped by type:

| Ruby | Excel |
|---|---|
| `Integer`, `Float`, any `Numeric` | number |
| `String` | string |
| `Time` | number (Excel serial date, local time) |
| `Date`, `DateTime` | number (Excel serial date, own offset) |
| `FastXlsx::Formula.new("SUM(A1:A9)")` | formula |
| `FastXlsx::URL.new("https://…")` | hyperlink |
| `true` / `false` | boolean |
| `nil` | empty cell |
| anything else | `to_s` as string |

`<<` and `concat` append after the last row written to that worksheet. In `constant_memory` mode rows are flushed as they are written, so fill each worksheet top to bottom.

Errors from the writer (invalid sheet names, out-of-order rows in constant memory mode, …) raise `FastXlsx::Error`.

## Performance

20,000 rows × 5 columns (integer, string, integer, `Time`, float), build + serialize to a String, median of 7 runs (3 for rubyXL), Apple Silicon, Ruby 4.0.5:

| Library | Time | vs fast_xlsx | Ruby objects allocated |
|---|---:|---:|---:|
| **fast_xlsx** (constant_memory) | **93 ms** | 1.0x | 20,006 |
| **fast_xlsx** | **102 ms** | 1.1x | 20,009 |
| [xlsxtream](https://github.com/felixbuenemann/xlsxtream) 3.1 | 184 ms | 2.0x | 561,740 |
| [fast_excel](https://github.com/Paxa/fast_excel) 0.5 (constant_memory) | 203 ms | 2.2x | 20,079 |
| [fast_excel](https://github.com/Paxa/fast_excel) 0.5 | 244 ms | 2.6x | 320,076 |
| [write_xlsx](https://github.com/cxn03651/write_xlsx) 1.15 | 594 ms | 6.4x | 1,483,899 |
| [caxlsx](https://github.com/caxlsx/caxlsx) 4.5 | 701 ms | 7.6x | 745,122 |
| [rubyXL](https://github.com/weshatheleopard/rubyXL) 3.4 | 2636 ms | 28.5x | 8,700,448 |

All outputs are 705–750 KB. Each library uses its own idiomatic row-append API; xlsxtream is a streaming writer with fewer features.

Reproduce:

```bash
bundle exec rake compile
BUNDLE_GEMFILE=bench/Gemfile bundle install
BUNDLE_GEMFILE=bench/Gemfile bundle exec ruby bench/compare.rb   # optional row count argument
```

## Installation

```bash
bundle add fast_xlsx
```

Precompiled gems are built for common platforms; other platforms need a Rust toolchain to install.

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

### Releasing

Bump `FastXlsx::VERSION`, update `CHANGELOG.md`, commit, then push a matching tag:

```bash
git tag v0.1.0 && git push origin v0.1.0
```

The `Build gems` workflow builds the source gem plus precompiled gems for each platform and pushes them to RubyGems (trusted publishing) and GitHub Packages. It refuses to publish if the tag and `VERSION` differ.

## License

MIT
