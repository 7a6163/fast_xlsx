## [Unreleased]

### Added

- `Worksheet#hide_rows`, `#hide_columns`, `#row_format` (default format for a row's cells, winning over `column_format` as in Excel) and `#default_row_height`.

### Fixed

- `nil` written with a format (e.g. a border, or a row format in `append`) writes a blank cell with that format; the format was dropped.

## [0.5.0] - 2026-10-01

### Added

- API documentation: every public method has YARD docs (shown on rubydoc.info).

### Changed

- Errors follow one rule (see the README): a row or column outside the sheet raises `RangeError` (row 1,048,576, column 16,384 or a row of more than 16,384 cells raised `FastXlsx::Error` or `ArgumentError`); a reversed range or single-cell merge raises `ArgumentError` (was `FastXlsx::Error`); a header or footer over 255 characters raises `FastXlsx::Error` like other text over Excel's limits (was `ArgumentError`); a negative or NaN image size, font size or margin raises `ArgumentError` (was written into the file or ignored); `zoom` converts numbers with `to_int` like other arguments (`zoom(150.9)` raised `TypeError`); `column_width` outside 0..255 and `row_height` outside 0..409 raise `ArgumentError` (a negative one hid the column or row).
- Date and time cells written with a format that has no `num_format` (e.g. `append(row, format: bold)`) get the default date format added, instead of showing as serial numbers. This applies to a cell's own format, a table column's and `column_format`'s; a format with a `num_format` is used as is.

## [0.4.0] - 2026-10-01

### Changed

- Dates and times written without a format now get `yyyy-mm-dd` (`Date`) or `yyyy-mm-dd hh:mm:ss` (`Time`, `DateTime`) instead of showing as serial numbers. A cell's own format, its table column's and `column_format` still win. Applying the format makes writing date-heavy data about 15% slower, the same as passing a date format yourself.

## [0.3.0] - 2026-10-01

### Changed

- `to_xlsx` and `save` release Ruby's global lock while building and compressing the file, so other threads (e.g. in Puma or Sidekiq) keep running. An interrupt (Ctrl-C, `Timeout`, `Thread#raise`) takes effect once the save finishes.
- Saving is 20-30% faster: files are compressed with zlib-rs instead of C zlib (which also drops the `libz-sys` dependency). Adding rows is about 15% faster from building with link-time optimisation.

### Added

- `write`, `write_comment`, `insert_image`, `insert_chart` and `freeze_panes` also take a cell reference such as `"B2"`. `write` is now native, which makes it about 15% faster.
- `Worksheet#zoom`, `#tab_color`, `#hide_gridlines`, `#activate`, `#hide`, and `#page_setup` (orientation, paper size, fit to pages, rows/columns repeated on every page, print area, printed gridlines).

### Fixed

- A `merge_range` that overlaps an earlier merge raises before writing anything; before, it also blanked the earlier merge's value.

## [0.2.0] - 2026-10-01

### Added

- Range methods (`autofilter`, `merge_range`, `conditional_format`, `data_validation`, `add_table`) also take an Excel reference (`"A1:D10"`) or rows and columns as Integers or Ranges (`0..9, 0..3`).
- `Workbook#define_name` for workbook-wide and sheet-scoped defined names.
- `Worksheet#group_rows` and `#group_columns` (outline groups, optionally collapsed); `group_rows` raises in `:constant` / `:low` mode.
- `Worksheet#protect(password:, allow:)` locks a sheet; `Format` options `locked: false` and `hidden: true` keep cells editable or hide formulas.

### Changed

- `column_width`, `column_format` and the other methods that take a Range raise `ArgumentError` for an empty, reversed or endless Range, or one that isn't Integers, instead of a `TypeError` or `RangeError`.

## [0.1.2] - 2026-09-30

### Fixed

- Values Excel cannot hold (a string over 32,767 characters, invalid UTF-8, a bad or over-long URL, more than 16,384 columns) raise before anything is written, instead of leaving the row half written.
- `merge_range` with such a value no longer leaves the range merged; its value also takes its table column's format.
- An unnamed worksheet takes the first free `SheetN` instead of a name that clashes with one given earlier (which failed only when saving).
- Cell values and formats are copied out of Ruby objects when a row is converted, so Ruby code run while converting (a `to_s`) cannot change or free them before they are written.
- Ruby code run while converting a value (a `to_s`, `jd`, ...) that touches the same workbook, or another thread writing to it, no longer deadlocks. A value that fails to convert leaves its row unwritten.
- In `:constant` / `:low` mode, `merge_range` and `add_table` on rows already written to disk raise `FastXlsx::Error` instead of being dropped; cells beside a tall merge can still be written.
- Table column formats apply to rows written after `add_table`.
- Invalid worksheet names raise without leaving an extra sheet behind; names that differ only in case raise at `add_worksheet`, not at save.
- Strings in other encodings (e.g. Windows-1252) are converted to UTF-8.
- `data_validation` lists accept numbers and other values (listed by `to_s`).
- Dates match Excel's 1900 date system (serials before 1900-03-01 were one too high); dates before 1900 raise `ArgumentError`.
- `Workbook#save` accepts a `Pathname`.

## [0.1.1] - 2026-09-30

### Fixed

- Precompiled (platform) gems failed to `require`: the extension was only looked up where a source install puts it, not in the per-Ruby-version directory that platform gems use. 0.1.0's platform gems cannot be loaded; upgrade to 0.1.1 (`bundle update fast_xlsx` if your lockfile has 0.1.0).

### Changed

- Releases now install and load the built platform gems on Linux, macOS and Windows with Ruby 3.3, 3.4 and 4.0 before publishing.

## [0.1.0] - 2026-09-30

First release: a fast `.xlsx` writer built on rust_xlsxwriter. Early API; it may still change before 1.0.

### Workbooks and worksheets

- `FastXlsx::Workbook.new(memory: :standard | :constant | :low)`: keep every cell in memory, or write finished rows to disk with strings inline (`:constant`) or in the shared string table (`:low`)
- `Workbook#add_worksheet`, `#worksheet(name)`, `#worksheets`, `#properties` (document properties), `#to_xlsx`, `#save`
- `Worksheet#<<`, `#append`, `#concat`, `#write` (all return the worksheet); writes to rows already on disk raise `FastXlsx::Error`
- Rows and columns are 0-based

### Cell values

- Numeric, String, Time, Date/DateTime, boolean and nil
- `FastXlsx::Formula`, `FastXlsx::URL` (optionally showing other text), `FastXlsx::RichString` (a format per text segment)

### Formatting

- `FastXlsx::Format`: fonts, colors, number formats, alignment, wrapping, rotation, indent, borders and underline styles; unknown options and invalid values raise `ArgumentError`
- One format per row or per cell in `append`; `Worksheet#column_format` for column defaults

### Layout and features

- `column_width`, `autofit` (keeps widths set explicitly), `row_height`, `freeze_panes`, `merge_range`, `autofilter`
- `conditional_format` (cell, text, formula, data bar, color scale), `data_validation` (lists, numbers, text length, messages)
- `add_table` (styles, total row, column totals), `insert_chart`, `insert_image` (path or IO, scale or pixel size), `write_comment`
- Printing: `page_header`, `page_footer`, `margins`, `page_breaks`, `vertical_page_breaks`
- Unknown options to the option-hash methods raise `ArgumentError`

### Packaging

- Precompiled gems for Linux (glibc and musl, x86_64 and aarch64, arm), macOS (arm64 and x86_64) and Windows (x64); requires CRuby 3.3+
