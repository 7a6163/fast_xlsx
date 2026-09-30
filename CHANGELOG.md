## [Unreleased]

- `Workbook` / `Worksheet` with `<<`, `concat`, `write`, `to_xlsx`, `save`
- `constant_memory` mode
- Numeric, String, Time, Date/DateTime, boolean and nil cell values
- `FastXlsx::Formula` cell values
- `FastXlsx::URL` hyperlink cell values
- `FastXlsx::Format` (bold, italic, underline, num_format) for `Worksheet#write` and `Worksheet#append`
- Raise `FastXlsx::Error` instead of silently dropping writes to already-flushed rows in `constant_memory` mode
- `Worksheet#set_column_width`, `#autofit` and `#autofilter`
- Format options `font_size`, `font_name`, `font_color`, `bg_color`, `align`, `valign`, `text_wrap` and borders
- `Worksheet#append(values, format: [...])` takes one format per cell
- `Worksheet#set_column_format` for column default formats
