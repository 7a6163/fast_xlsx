# FastXlsx

Fast `.xlsx` writer for Ruby, built on [rust_xlsxwriter](https://github.com/jmcnamara/rust_xlsxwriter) via [magnus](https://github.com/matsadler/magnus).

> **Status: early.** Values only — formats, formulas, URLs, column widths and charts are not implemented yet.

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

Values are mapped by type:

| Ruby | Excel |
|---|---|
| `Integer`, `Float`, any `Numeric` | number |
| `String` | string |
| `Time` | number (Excel serial date, local time) |
| `Date`, `DateTime` | number (Excel serial date, own offset) |
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

## License

MIT
