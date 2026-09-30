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
ws.autofit                        # size columns to the data written so far
ws.autofilter(0, 0, 100, 3)       # filter buttons on A1:D101 (first_row, first_col, last_row, last_col)
```

`autofit` only sees rows still in memory, so it has no effect on rows already flushed in `constant_memory` mode.

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

20,000 rows × 5 columns (integer, string, integer, `Time`, float), write + serialize, median of 7, Apple Silicon, Ruby 4.0:

| | default | constant_memory |
|---|---|---|
| fast_xlsx `<<` | 103 ms | 94 ms |
| fast_xlsx `concat` | 101 ms | 91 ms |
| [fast_excel](https://github.com/Paxa/fast_excel) 0.5 `<<` | 241 ms | 199 ms |

Reproduce with `bench/write.rb`.

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
