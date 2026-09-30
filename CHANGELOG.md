## [Unreleased]

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
