use std::cell::Cell;
use std::sync::{Arc, Mutex};

use magnus::{
    function, method, prelude::*, r_hash::ForEach, typed_data::Obj, value::Lazy, Error,
    ExceptionClass, Integer, RArray, RClass, RHash, RModule, RString, Ruby, Symbol, TryConvert,
    Value,
};
use rust_xlsxwriter::{
    Color, ConditionalFormat2ColorScale, ConditionalFormat3ColorScale, ConditionalFormatCell,
    ConditionalFormatCellRule, ConditionalFormatDataBar, ConditionalFormatFormula,
    ConditionalFormatText, ConditionalFormatTextRule, ConditionalFormatValue, DataValidation,
    DataValidationRule, FormatAlign, FormatBorder, FormatScript, FormatUnderline,
    IntoDataValidationValue, IntoExcelData, XlsxError,
};

use rust_xlsxwriter::{Chart, ChartType, DocProperties, Image, Note};

// These constants are defined in lib/fast_xlsx.rb before this extension loads.
fn fast_xlsx_const<T: TryConvert>(ruby: &Ruby, name: &str) -> T {
    ruby.class_object()
        .const_get::<_, RModule>("FastXlsx")
        .and_then(|m| m.const_get(name))
        .unwrap()
}

static ERROR: Lazy<ExceptionClass> = Lazy::new(|ruby| fast_xlsx_const(ruby, "Error"));
static FORMULA: Lazy<RClass> = Lazy::new(|ruby| fast_xlsx_const(ruby, "Formula"));
static URL: Lazy<RClass> = Lazy::new(|ruby| fast_xlsx_const(ruby, "URL"));

fn xerr(e: XlsxError) -> Error {
    let ruby = Ruby::get().unwrap();
    Error::new(ruby.get_inner(&ERROR), e.to_string())
}

type Shared = Arc<Mutex<rust_xlsxwriter::Workbook>>;

#[magnus::wrap(class = "FastXlsx::Workbook", free_immediately)]
struct Workbook {
    inner: Shared,
    constant_memory: bool,
}

#[magnus::wrap(class = "FastXlsx::Worksheet", free_immediately)]
struct Worksheet {
    wb: Shared,
    index: usize,
    next_row: Cell<u32>,
    constant_memory: bool,
}

impl Workbook {
    fn new(constant_memory: bool) -> Self {
        Workbook {
            inner: Arc::new(Mutex::new(rust_xlsxwriter::Workbook::new())),
            constant_memory,
        }
    }

    fn add_worksheet(&self, name: Option<String>) -> Result<Worksheet, Error> {
        let mut wb = self.inner.lock().unwrap();
        let ws = if self.constant_memory {
            wb.add_worksheet_with_constant_memory()
        } else {
            wb.add_worksheet()
        };
        if let Some(name) = name {
            ws.set_name(name).map_err(xerr)?;
        }
        Ok(Worksheet {
            wb: self.inner.clone(),
            index: wb.worksheets().len() - 1,
            next_row: Cell::new(0),
            constant_memory: self.constant_memory,
        })
    }

    fn to_xlsx(ruby: &Ruby, rb_self: &Self) -> Result<RString, Error> {
        let buf = rb_self
            .inner
            .lock()
            .unwrap()
            .save_to_buffer()
            .map_err(xerr)?;
        Ok(ruby.str_from_slice(&buf))
    }

    fn save(&self, path: String) -> Result<(), Error> {
        self.inner.lock().unwrap().save(path).map_err(xerr)
    }

    fn set_properties(ruby: &Ruby, rb_self: &Self, fields: RHash) -> Result<(), Error> {
        let mut props = DocProperties::new();
        fields.foreach(|key: Symbol, value: String| {
            props = match &*key.name()? {
                "title" => props.clone().set_title(value),
                "subject" => props.clone().set_subject(value),
                "author" => props.clone().set_author(value),
                "manager" => props.clone().set_manager(value),
                "company" => props.clone().set_company(value),
                "category" => props.clone().set_category(value),
                "keywords" => props.clone().set_keywords(value),
                "comments" => props.clone().set_comment(value),
                "status" => props.clone().set_status(value),
                other => {
                    return Err(Error::new(
                        ruby.exception_arg_error(),
                        format!("unknown property: {other}"),
                    ))
                }
            };
            Ok(ForEach::Continue)
        })?;
        rb_self.inner.lock().unwrap().set_properties(&props);
        Ok(())
    }
}

// Excel stores datetimes as days since 1900-01-01 in local time.
fn excel_time(v: Value) -> Result<f64, Error> {
    let secs: f64 = v.funcall("to_f", ())?;
    let offset: i64 = v.funcall("utc_offset", ())?;
    Ok((secs + offset as f64) / 86400.0 + 25569.0)
}

// Date / DateTime: Julian day and day fraction are both in the object's own
// offset. JD 2415019 is Excel serial 0 (1899-12-30).
fn excel_date(v: Value) -> Result<f64, Error> {
    let jd: i64 = v.funcall("jd", ())?;
    let fraction: f64 = f64::try_convert(v.funcall("day_fraction", ())?)?;
    Ok((jd - 2_415_019) as f64 + fraction)
}

fn emit<T: IntoExcelData>(
    ws: &mut rust_xlsxwriter::Worksheet,
    row: u32,
    col: u16,
    data: T,
    format: Option<&Format>,
) -> Result<(), Error> {
    match format {
        Some(f) => ws.write_with_format(row, col, data, &f.0),
        None => ws.write(row, col, data),
    }
    .map(|_| ())
    .map_err(xerr)
}

fn put(
    ruby: &Ruby,
    ws: &mut rust_xlsxwriter::Worksheet,
    row: u32,
    col: u16,
    v: Value,
    format: Option<&Format>,
) -> Result<(), Error> {
    if v.is_nil() {
        Ok(())
    } else if let Some(s) = RString::from_value(v) {
        // SAFETY: the borrowed str is copied by the writer before any Ruby code runs.
        emit(ws, row, col, unsafe { s.as_str()? }, format)
    } else if v.is_kind_of(ruby.class_numeric()) {
        emit(ws, row, col, f64::try_convert(v)?, format)
    } else if v.is_kind_of(ruby.class_time()) {
        emit(ws, row, col, excel_time(v)?, format)
    } else if v.is_kind_of(ruby.class_true_class()) || v.is_kind_of(ruby.class_false_class()) {
        emit(ws, row, col, v.to_bool(), format)
    } else if v.is_kind_of(ruby.get_inner(&FORMULA)) {
        let expression: String = v.funcall("expression", ())?;
        emit(
            ws,
            row,
            col,
            rust_xlsxwriter::Formula::new(expression),
            format,
        )
    } else if v.is_kind_of(ruby.get_inner(&URL)) {
        let url: String = v.funcall("url", ())?;
        let mut link = rust_xlsxwriter::Url::new(url);
        if let Some(text) = v.funcall::<_, _, Option<String>>("text", ())? {
            link = link.set_text(text);
        }
        emit(ws, row, col, link, format)
    } else if v.respond_to("jd", false)? {
        emit(ws, row, col, excel_date(v)?, format)
    } else {
        let s: String = v.funcall("to_s", ())?;
        emit(ws, row, col, s, format)
    }
}

// "#RRGGBB" or 0xRRGGBB.
fn color(ruby: &Ruby, value: Value) -> Result<Color, Error> {
    let rgb = if let Some(i) = Integer::from_value(value) {
        i.to_u32().ok().filter(|n| *n <= 0xFF_FFFF)
    } else if let Some(s) = RString::from_value(value) {
        let s = s.to_string()?;
        s.strip_prefix('#')
            .filter(|hex| hex.len() == 6 && hex.chars().all(|c| c.is_ascii_hexdigit()))
            .and_then(|hex| u32::from_str_radix(hex, 16).ok())
    } else {
        None
    };
    rgb.map(Color::RGB).ok_or_else(|| {
        Error::new(
            ruby.exception_arg_error(),
            format!(
                "invalid color {}: use \"#RRGGBB\" or 0xRRGGBB",
                value.inspect()
            ),
        )
    })
}

// A symbol option that must be one of `choices`.
fn choice<T: Clone>(
    ruby: &Ruby,
    option: &str,
    value: Value,
    choices: &[(&str, T)],
) -> Result<T, Error> {
    let name = Symbol::from_value(value).and_then(|s| s.name().ok());
    name.and_then(|n| choices.iter().find(|(k, _)| *k == n))
        .map(|(_, v)| v.clone())
        .ok_or_else(|| {
            let names: Vec<_> = choices.iter().map(|(k, _)| format!(":{k}")).collect();
            Error::new(
                ruby.exception_arg_error(),
                format!(
                    "invalid {option} {}: expected one of {}",
                    value.inspect(),
                    names.join(", ")
                ),
            )
        })
}

// A number or string compared against in a conditional format rule.
fn cf_value(ruby: &Ruby, v: Value) -> Result<ConditionalFormatValue, Error> {
    if let Some(s) = RString::from_value(v) {
        Ok(s.to_string()?.into())
    } else if v.is_kind_of(ruby.class_numeric()) {
        Ok(f64::try_convert(v)?.into())
    } else {
        Err(Error::new(
            ruby.exception_arg_error(),
            format!(
                "invalid conditional format value {}: use a number or string",
                v.inspect()
            ),
        ))
    }
}

// A comparison criteria and its operand(s): one value, or [min, max] for the
// range criteria.
fn comparison(
    ruby: &Ruby,
    what: &str,
    criteria: Value,
    value: Value,
) -> Result<(&'static str, Value, Option<Value>), Error> {
    let op = choice(
        ruby,
        what,
        criteria,
        &[
            ("==", "=="),
            ("!=", "!="),
            (">", ">"),
            (">=", ">="),
            ("<", "<"),
            ("<=", "<="),
            ("between", "between"),
            ("not_between", "not_between"),
        ],
    )?;
    if op != "between" && op != "not_between" {
        return Ok((op, value, None));
    }
    let bounds = RArray::from_value(value)
        .filter(|a| a.len() == 2)
        .ok_or_else(|| {
            Error::new(
                ruby.exception_arg_error(),
                format!("{op} needs value: [min, max], got {}", value.inspect()),
            )
        })?;
    Ok((op, bounds.entry(0)?, Some(bounds.entry(1)?)))
}

// Maps a comparison() result onto a rust_xlsxwriter rule enum, converting the
// operands with $convert.
macro_rules! comparison_rule {
    ($Rule:ident, $cmp:expr, $convert:expr) => {{
        let (op, a, b) = $cmp;
        let a = $convert(a)?;
        match (op, b) {
            ("between", Some(b)) => $Rule::Between(a, $convert(b)?),
            ("not_between", Some(b)) => $Rule::NotBetween(a, $convert(b)?),
            ("==", _) => $Rule::EqualTo(a),
            ("!=", _) => $Rule::NotEqualTo(a),
            (">", _) => $Rule::GreaterThan(a),
            (">=", _) => $Rule::GreaterThanOrEqualTo(a),
            ("<", _) => $Rule::LessThan(a),
            _ => $Rule::LessThanOrEqualTo(a),
        }
    }};
}

fn cell_rule(
    ruby: &Ruby,
    criteria: Value,
    value: Value,
) -> Result<ConditionalFormatCellRule<ConditionalFormatValue>, Error> {
    let cmp = comparison(ruby, "cell criteria", criteria, value)?;
    Ok(comparison_rule!(ConditionalFormatCellRule, cmp, |v| {
        cf_value(ruby, v)
    }))
}

fn validation_rule<T: TryConvert + IntoDataValidationValue>(
    ruby: &Ruby,
    criteria: Value,
    value: Value,
) -> Result<DataValidationRule<T>, Error> {
    let cmp = comparison(ruby, "validation criteria", criteria, value)?;
    Ok(comparison_rule!(DataValidationRule, cmp, T::try_convert))
}

// Excel's 255-character limit, counted the way rust_xlsxwriter does (with the
// long &[Page]-style codes shortened), which otherwise drops the text silently.
fn check_header_footer(ruby: &Ruby, what: &str, text: &str) -> Result<(), Error> {
    let mut expanded = text.to_string();
    for (long, short) in [
        ("&[Tab]", "&A"),
        ("&[Date]", "&D"),
        ("&[File]", "&F"),
        ("&[Page]", "&P"),
        ("&[Path]", "&Z"),
        ("&[Time]", "&T"),
        ("&[Pages]", "&N"),
        ("&[Picture]", "&G"),
    ] {
        expanded = expanded.replace(long, short);
    }
    let len = expanded.chars().count();
    if len > 255 {
        return Err(Error::new(
            ruby.exception_arg_error(),
            format!("{what} is {len} characters; Excel allows 255"),
        ));
    }
    Ok(())
}

// Rejects option keys outside `allowed`, so a typo raises instead of being ignored.
fn check_keys(ruby: &Ruby, options: RHash, allowed: &[&str], what: &str) -> Result<(), Error> {
    options.foreach(|key: Symbol, _: Value| {
        let name = key.name()?;
        if allowed.contains(&&*name) {
            Ok(ForEach::Continue)
        } else {
            Err(Error::new(
                ruby.exception_arg_error(),
                format!(
                    "unknown {what} option: {name} (expected one of {})",
                    allowed.join(", ")
                ),
            ))
        }
    })
}

const CHART_TYPES: &[(&str, ChartType)] = &[
    ("area", ChartType::Area),
    ("area_stacked", ChartType::AreaStacked),
    ("bar", ChartType::Bar),
    ("bar_stacked", ChartType::BarStacked),
    ("column", ChartType::Column),
    ("column_stacked", ChartType::ColumnStacked),
    ("line", ChartType::Line),
    ("line_stacked", ChartType::LineStacked),
    ("pie", ChartType::Pie),
    ("doughnut", ChartType::Doughnut),
    ("radar", ChartType::Radar),
    ("scatter", ChartType::Scatter),
];

const BORDERS: &[(&str, FormatBorder)] = &[
    ("thin", FormatBorder::Thin),
    ("medium", FormatBorder::Medium),
    ("thick", FormatBorder::Thick),
    ("dashed", FormatBorder::Dashed),
    ("dotted", FormatBorder::Dotted),
    ("double", FormatBorder::Double),
    ("hair", FormatBorder::Hair),
];

#[magnus::wrap(class = "FastXlsx::Format", free_immediately)]
struct Format(rust_xlsxwriter::Format);

impl Format {
    fn new(ruby: &Ruby, options: RHash) -> Result<Self, Error> {
        let mut f = rust_xlsxwriter::Format::new();
        options.foreach(|key: Symbol, value: Value| {
            let taken = std::mem::take(&mut f);
            f = match &*key.name()? {
                "bold" if value.to_bool() => taken.set_bold(),
                "italic" if value.to_bool() => taken.set_italic(),
                "underline" if value.is_kind_of(ruby.class_true_class()) => {
                    taken.set_underline(FormatUnderline::Single)
                }
                "underline" if value.to_bool() => taken.set_underline(choice(
                    ruby,
                    "underline",
                    value,
                    &[
                        ("single", FormatUnderline::Single),
                        ("double", FormatUnderline::Double),
                        ("single_accounting", FormatUnderline::SingleAccounting),
                        ("double_accounting", FormatUnderline::DoubleAccounting),
                    ],
                )?),
                "strikeout" if value.to_bool() => taken.set_font_strikethrough(),
                "font_script" => taken.set_font_script(choice(
                    ruby,
                    "font_script",
                    value,
                    &[
                        ("superscript", FormatScript::Superscript),
                        ("subscript", FormatScript::Subscript),
                    ],
                )?),
                "rotation" => {
                    let degrees = i16::try_convert(value)?;
                    if !(-90..=90).contains(&degrees) && degrees != 270 {
                        return Err(Error::new(
                            ruby.exception_arg_error(),
                            format!("invalid rotation {degrees}: use -90..90 or 270"),
                        ));
                    }
                    taken.set_rotation(degrees)
                }
                "indent" => taken.set_indent(u8::try_convert(value)?),
                "shrink" if value.to_bool() => taken.set_shrink(),
                "border_color" => taken.set_border_color(color(ruby, value)?),
                "num_format" => taken.set_num_format(String::try_convert(value)?),
                "font_size" => taken.set_font_size(f64::try_convert(value)?),
                "font_name" => taken.set_font_name(String::try_convert(value)?),
                "font_color" => taken.set_font_color(color(ruby, value)?),
                "bg_color" => taken.set_background_color(color(ruby, value)?),
                "align" => taken.set_align(choice(
                    ruby,
                    "align",
                    value,
                    &[
                        ("left", FormatAlign::Left),
                        ("center", FormatAlign::Center),
                        ("right", FormatAlign::Right),
                    ],
                )?),
                "valign" => taken.set_align(choice(
                    ruby,
                    "valign",
                    value,
                    &[
                        ("top", FormatAlign::Top),
                        ("center", FormatAlign::VerticalCenter),
                        ("bottom", FormatAlign::Bottom),
                    ],
                )?),
                "text_wrap" if value.to_bool() => taken.set_text_wrap(),
                "border" => taken.set_border(choice(ruby, "border", value, BORDERS)?),
                "border_left" => taken.set_border_left(choice(ruby, "border", value, BORDERS)?),
                "border_right" => taken.set_border_right(choice(ruby, "border", value, BORDERS)?),
                "border_top" => taken.set_border_top(choice(ruby, "border", value, BORDERS)?),
                "border_bottom" => taken.set_border_bottom(choice(ruby, "border", value, BORDERS)?),
                "bold" | "italic" | "underline" | "text_wrap" | "strikeout" | "shrink" => taken,
                other => {
                    return Err(Error::new(
                        ruby.exception_arg_error(),
                        format!("unknown format option: {other}"),
                    ))
                }
            };
            Ok(ForEach::Continue)
        })?;
        Ok(Format(f))
    }
}

// A row's format: one Format (or nil) for every cell, or an Array with one per cell.
enum RowFormat {
    Same(Option<&'static Format>),
    PerCell(RArray),
}

impl RowFormat {
    fn from_value(v: Value) -> Result<Self, Error> {
        match RArray::from_value(v) {
            Some(formats) => Ok(RowFormat::PerCell(formats)),
            None => Ok(RowFormat::Same(Option::<&Format>::try_convert(v)?)),
        }
    }

    fn at(&self, col: usize) -> Result<Option<&Format>, Error> {
        match self {
            RowFormat::Same(f) => Ok(*f),
            RowFormat::PerCell(formats) => {
                Option::<&Format>::try_convert(formats.entry::<Value>(col as isize)?)
            }
        }
    }
}

impl Worksheet {
    fn advance(&self, row: u32) {
        self.next_row.set(self.next_row.get().max(row + 1));
    }

    fn with_ws<T>(
        &self,
        f: impl FnOnce(&mut rust_xlsxwriter::Worksheet) -> Result<T, Error>,
    ) -> Result<T, Error> {
        let mut wb = self.wb.lock().unwrap();
        f(wb.worksheet_from_index(self.index).map_err(xerr)?)
    }

    fn write_row(
        &self,
        ruby: &Ruby,
        row: u32,
        cells: RArray,
        format: &RowFormat,
    ) -> Result<(), Error> {
        self.with_ws(|ws| {
            for (i, v) in cells.into_iter().enumerate() {
                let col = u16::try_from(i)
                    .map_err(|_| Error::new(ruby.exception_arg_error(), "too many columns"))?;
                put(ruby, ws, row, col, v, format.at(i)?)?;
            }
            Ok(())
        })?;
        self.advance(row);
        Ok(())
    }

    fn write(
        ruby: &Ruby,
        rb_self: &Self,
        row: u32,
        col: u16,
        v: Value,
        format: Option<&Format>,
    ) -> Result<(), Error> {
        // rust_xlsxwriter silently drops writes to rows it has already flushed.
        if rb_self.constant_memory && row + 1 < rb_self.next_row.get() {
            return Err(Error::new(
                ruby.get_inner(&ERROR),
                format!("row {row} was already flushed in constant_memory mode"),
            ));
        }
        rb_self.with_ws(|ws| put(ruby, ws, row, col, v, format))?;
        rb_self.advance(row);
        Ok(())
    }

    fn append(
        ruby: &Ruby,
        rb_self: Obj<Self>,
        cells: RArray,
        format: Value,
    ) -> Result<Obj<Self>, Error> {
        let format = RowFormat::from_value(format)?;
        rb_self.write_row(ruby, rb_self.next_row.get(), cells, &format)?;
        Ok(rb_self)
    }

    fn push(ruby: &Ruby, rb_self: Obj<Self>, cells: RArray) -> Result<Obj<Self>, Error> {
        rb_self.write_row(ruby, rb_self.next_row.get(), cells, &RowFormat::Same(None))?;
        Ok(rb_self)
    }

    fn concat(ruby: &Ruby, rb_self: Obj<Self>, rows: RArray) -> Result<Obj<Self>, Error> {
        for r in rows.into_iter() {
            let cells = RArray::try_convert(r)?;
            rb_self.write_row(ruby, rb_self.next_row.get(), cells, &RowFormat::Same(None))?;
        }
        Ok(rb_self)
    }

    fn set_column_width(&self, first: u16, last: u16, width: f64) -> Result<(), Error> {
        self.with_ws(|ws| {
            ws.set_column_range_width(first, last, width)
                .map(|_| ())
                .map_err(xerr)
        })
    }

    fn set_column_format(&self, first: u16, last: u16, format: &Format) -> Result<(), Error> {
        self.with_ws(|ws| {
            ws.set_column_range_format(first, last, &format.0)
                .map(|_| ())
                .map_err(xerr)
        })
    }

    fn autofilter(
        rb_self: Obj<Self>,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
    ) -> Result<Obj<Self>, Error> {
        rb_self.with_ws(|ws| {
            ws.autofilter(first_row, first_col, last_row, last_col)
                .map(|_| ())
                .map_err(xerr)
        })?;
        Ok(rb_self)
    }

    fn autofit(rb_self: Obj<Self>) -> Result<Obj<Self>, Error> {
        rb_self.with_ws(|ws| {
            ws.autofit();
            Ok(())
        })?;
        Ok(rb_self)
    }

    fn freeze_panes(rb_self: Obj<Self>, row: u32, col: u16) -> Result<Obj<Self>, Error> {
        rb_self.with_ws(|ws| ws.set_freeze_panes(row, col).map(|_| ()).map_err(xerr))?;
        Ok(rb_self)
    }

    fn set_row_height(rb_self: Obj<Self>, row: u32, height: f64) -> Result<Obj<Self>, Error> {
        rb_self.with_ws(|ws| ws.set_row_height(row, height).map(|_| ()).map_err(xerr))?;
        Ok(rb_self)
    }

    fn set_header(ruby: &Ruby, rb_self: &Self, text: String) -> Result<(), Error> {
        check_header_footer(ruby, "header", &text)?;
        rb_self.with_ws(|ws| {
            ws.set_header(text);
            Ok(())
        })
    }

    fn set_footer(ruby: &Ruby, rb_self: &Self, text: String) -> Result<(), Error> {
        check_header_footer(ruby, "footer", &text)?;
        rb_self.with_ws(|ws| {
            ws.set_footer(text);
            Ok(())
        })
    }

    // Inches; a negative value keeps the current margin.
    fn set_margins(
        &self,
        left: f64,
        right: f64,
        top: f64,
        bottom: f64,
        header: f64,
        footer: f64,
    ) -> Result<(), Error> {
        self.with_ws(|ws| {
            ws.set_margins(left, right, top, bottom, header, footer);
            Ok(())
        })
    }

    fn set_page_breaks(rb_self: Obj<Self>, rows: Vec<u32>) -> Result<Obj<Self>, Error> {
        rb_self.with_ws(|ws| ws.set_page_breaks(&rows).map(|_| ()).map_err(xerr))?;
        Ok(rb_self)
    }

    fn set_vertical_page_breaks(rb_self: Obj<Self>, cols: Vec<u32>) -> Result<Obj<Self>, Error> {
        rb_self.with_ws(|ws| ws.set_vertical_page_breaks(&cols).map(|_| ()).map_err(xerr))?;
        Ok(rb_self)
    }

    // rust_xlsxwriter only merges with a string, so merge with "" and then write
    // the value into the first cell, which keeps its type.
    #[allow(clippy::too_many_arguments)]
    fn merge_range(
        ruby: &Ruby,
        rb_self: Obj<Self>,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
        v: Value,
        format: Option<&Format>,
    ) -> Result<Obj<Self>, Error> {
        let default = rust_xlsxwriter::Format::new();
        let merge_format = format.map_or(&default, |f| &f.0);
        rb_self.with_ws(|ws| {
            ws.merge_range(first_row, first_col, last_row, last_col, "", merge_format)
                .map_err(xerr)?;
            put(ruby, ws, first_row, first_col, v, format)
        })?;
        rb_self.advance(last_row);
        Ok(rb_self)
    }

    #[allow(clippy::too_many_arguments)]
    fn conditional_format(
        ruby: &Ruby,
        rb_self: Obj<Self>,
        fr: u32,
        fc: u16,
        lr: u32,
        lc: u16,
        options: RHash,
    ) -> Result<Obj<Self>, Error> {
        let nil = ruby.qnil().as_value();
        let opt = |name: &str| options.get(ruby.to_symbol(name)).unwrap_or(nil);
        check_keys(
            ruby,
            options,
            &["type", "criteria", "value", "format", "colors"],
            "conditional_format",
        )?;
        let format = Option::<&Format>::try_convert(opt("format"))?;
        let value = opt("value");
        let kind = choice(
            ruby,
            "conditional format type",
            opt("type"),
            &[
                ("cell", "cell"),
                ("text", "text"),
                ("formula", "formula"),
                ("data_bar", "data_bar"),
                ("color_scale", "color_scale"),
            ],
        )?;
        rb_self.with_ws(|ws| {
            let r = match kind {
                "cell" => {
                    let mut cf = ConditionalFormatCell::new().set_rule(cell_rule(
                        ruby,
                        opt("criteria"),
                        value,
                    )?);
                    if let Some(f) = format {
                        cf = cf.set_format(&f.0);
                    }
                    ws.add_conditional_format(fr, fc, lr, lc, &cf)
                }
                "text" => {
                    let text = String::try_convert(value)?;
                    let rule = match choice(
                        ruby,
                        "text criteria",
                        opt("criteria"),
                        &[
                            ("contains", 0),
                            ("not_contains", 1),
                            ("begins_with", 2),
                            ("ends_with", 3),
                        ],
                    )? {
                        0 => ConditionalFormatTextRule::Contains(text),
                        1 => ConditionalFormatTextRule::DoesNotContain(text),
                        2 => ConditionalFormatTextRule::BeginsWith(text),
                        _ => ConditionalFormatTextRule::EndsWith(text),
                    };
                    let mut cf = ConditionalFormatText::new().set_rule(rule);
                    if let Some(f) = format {
                        cf = cf.set_format(&f.0);
                    }
                    ws.add_conditional_format(fr, fc, lr, lc, &cf)
                }
                "formula" => {
                    let mut cf = ConditionalFormatFormula::new()
                        .set_rule(String::try_convert(value)?.as_str());
                    if let Some(f) = format {
                        cf = cf.set_format(&f.0);
                    }
                    ws.add_conditional_format(fr, fc, lr, lc, &cf)
                }
                "data_bar" => {
                    ws.add_conditional_format(fr, fc, lr, lc, &ConditionalFormatDataBar::new())
                }
                _ => {
                    let colors = if opt("colors").is_nil() {
                        3
                    } else {
                        u8::try_convert(opt("colors"))?
                    };
                    match colors {
                        2 => ws.add_conditional_format(
                            fr,
                            fc,
                            lr,
                            lc,
                            &ConditionalFormat2ColorScale::new(),
                        ),
                        3 => ws.add_conditional_format(
                            fr,
                            fc,
                            lr,
                            lc,
                            &ConditionalFormat3ColorScale::new(),
                        ),
                        n => {
                            return Err(Error::new(
                                ruby.exception_arg_error(),
                                format!("invalid colors {n}: use 2 or 3"),
                            ))
                        }
                    }
                }
            };
            r.map(|_| ()).map_err(xerr)
        })?;
        Ok(rb_self)
    }

    #[allow(clippy::too_many_arguments)]
    fn data_validation(
        ruby: &Ruby,
        rb_self: Obj<Self>,
        fr: u32,
        fc: u16,
        lr: u32,
        lc: u16,
        options: RHash,
    ) -> Result<Obj<Self>, Error> {
        let nil = ruby.qnil().as_value();
        let opt = |name: &str| options.get(ruby.to_symbol(name)).unwrap_or(nil);
        check_keys(
            ruby,
            options,
            &[
                "type",
                "criteria",
                "value",
                "input_title",
                "input_message",
                "error_title",
                "error_message",
            ],
            "data_validation",
        )?;
        let (criteria, value) = (opt("criteria"), opt("value"));
        let dv = DataValidation::new();
        let mut dv = match choice(
            ruby,
            "data validation type",
            opt("type"),
            &[
                ("list", "list"),
                ("whole_number", "whole_number"),
                ("decimal", "decimal"),
                ("text_length", "text_length"),
            ],
        )? {
            "list" => match RArray::from_value(value) {
                Some(items) => dv
                    .allow_list_strings(&items.to_vec::<String>()?)
                    .map_err(xerr)?,
                None => dv
                    .allow_list_formula(rust_xlsxwriter::Formula::new(String::try_convert(value)?)),
            },
            "whole_number" => dv.allow_whole_number(validation_rule(ruby, criteria, value)?),
            "decimal" => dv.allow_decimal_number(validation_rule(ruby, criteria, value)?),
            _ => dv.allow_text_length(validation_rule(ruby, criteria, value)?),
        };
        for (name, set) in [
            (
                "input_title",
                DataValidation::set_input_title as fn(_, String) -> _,
            ),
            ("input_message", DataValidation::set_input_message),
            ("error_title", DataValidation::set_error_title),
            ("error_message", DataValidation::set_error_message),
        ] {
            let text = opt(name);
            if !text.is_nil() {
                dv = set(dv, String::try_convert(text)?).map_err(xerr)?;
            }
        }
        rb_self.with_ws(|ws| {
            ws.add_data_validation(fr, fc, lr, lc, &dv)
                .map(|_| ())
                .map_err(xerr)
        })?;
        Ok(rb_self)
    }

    fn write_comment(
        rb_self: Obj<Self>,
        row: u32,
        col: u16,
        text: String,
        author: Option<String>,
    ) -> Result<Obj<Self>, Error> {
        // Keep the text as written; rust_xlsxwriter would prefix "Author:\n".
        let mut note = Note::new(text).add_author_prefix(false);
        if let Some(author) = author {
            note = note.set_author(author);
        }
        rb_self.with_ws(|ws| ws.insert_note(row, col, &note).map(|_| ()).map_err(xerr))?;
        Ok(rb_self)
    }

    fn insert_image(
        ruby: &Ruby,
        rb_self: Obj<Self>,
        row: u32,
        col: u16,
        bytes: RString,
        options: RHash,
    ) -> Result<Obj<Self>, Error> {
        let nil = ruby.qnil().as_value();
        let opt = |name: &str| options.get(ruby.to_symbol(name)).unwrap_or(nil);
        check_keys(
            ruby,
            options,
            &[
                "scale", "width", "height", "x_offset", "y_offset", "alt_text",
            ],
            "insert_image",
        )?;
        let offset = |name: &str| -> Result<u32, Error> {
            let v = opt(name);
            if v.is_nil() {
                Ok(0)
            } else {
                u32::try_convert(v)
            }
        };
        // SAFETY: the bytes are copied into the Image before any Ruby code runs.
        let mut image = Image::new_from_buffer(unsafe { bytes.as_slice() }).map_err(xerr)?;
        let pixels = |name: &str| -> Result<Option<f64>, Error> {
            let v = opt(name);
            if v.is_nil() {
                Ok(None)
            } else {
                f64::try_convert(v).map(Some)
            }
        };
        let (width, height) = (pixels("width")?, pixels("height")?);
        if !opt("scale").is_nil() {
            if width.is_some() || height.is_some() {
                return Err(Error::new(
                    ruby.exception_arg_error(),
                    "pass either scale: or width:/height:, not both",
                ));
            }
            let scale = f64::try_convert(opt("scale"))?;
            image = image.set_scale_width(scale).set_scale_height(scale);
        } else if width.is_some() || height.is_some() {
            // Displayed size at scale 1, as rust_xlsxwriter computes it from the DPI.
            let natural_w = image.width() * 96.0 / image.width_dpi();
            let natural_h = image.height() * 96.0 / image.height_dpi();
            let scale_w = width.map(|w| w / natural_w);
            let scale_h = height.map(|h| h / natural_h);
            // With only one dimension, scale the other by the same factor.
            let (sw, sh) = (scale_w.or(scale_h).unwrap(), scale_h.or(scale_w).unwrap());
            image = image.set_scale_width(sw).set_scale_height(sh);
        }
        if !opt("alt_text").is_nil() {
            image = image.set_alt_text(String::try_convert(opt("alt_text"))?);
        }
        let (x, y) = (offset("x_offset")?, offset("y_offset")?);
        rb_self.with_ws(|ws| {
            ws.insert_image_with_offset(row, col, &image, x, y)
                .map(|_| ())
                .map_err(xerr)
        })?;
        Ok(rb_self)
    }

    fn insert_chart(
        ruby: &Ruby,
        rb_self: Obj<Self>,
        row: u32,
        col: u16,
        options: RHash,
    ) -> Result<Obj<Self>, Error> {
        let nil = ruby.qnil().as_value();
        let arg_error = |msg: String| Error::new(ruby.exception_arg_error(), msg);
        let text = |v: Value| -> Result<Option<String>, Error> {
            if v.is_nil() {
                Ok(None)
            } else {
                String::try_convert(v).map(Some)
            }
        };
        let opt = |name: &str| options.get(ruby.to_symbol(name)).unwrap_or(nil);
        check_keys(
            ruby,
            options,
            &[
                "type", "series", "title", "x_axis", "y_axis", "width", "height",
            ],
            "insert_chart",
        )?;

        let mut chart = Chart::new(choice(ruby, "chart type", opt("type"), CHART_TYPES)?);
        let series = RArray::try_convert(opt("series"))?;
        if series.is_empty() {
            return Err(arg_error(
                "series: needs at least one { values:, categories:, name: } hash".into(),
            ));
        }
        for s in series.into_iter() {
            let s = RHash::try_convert(s)?;
            check_keys(ruby, s, &["values", "categories", "name"], "chart series")?;
            let get = |name: &str| s.get(ruby.to_symbol(name)).unwrap_or(nil);
            let values = text(get("values"))?
                .ok_or_else(|| arg_error(format!("series {} needs values:", s.inspect())))?;
            let cs = chart.add_series().set_values(values.as_str());
            if let Some(categories) = text(get("categories"))? {
                cs.set_categories(categories.as_str());
            }
            if let Some(name) = text(get("name"))? {
                cs.set_name(name.as_str());
            }
        }
        if let Some(title) = text(opt("title"))? {
            chart.title().set_name(title.as_str());
        }
        if let Some(name) = text(opt("x_axis"))? {
            chart.x_axis().set_name(name.as_str());
        }
        if let Some(name) = text(opt("y_axis"))? {
            chart.y_axis().set_name(name.as_str());
        }
        if !opt("width").is_nil() {
            chart.set_width(u32::try_convert(opt("width"))?);
        }
        if !opt("height").is_nil() {
            chart.set_height(u32::try_convert(opt("height"))?);
        }
        rb_self.with_ws(|ws| ws.insert_chart(row, col, &chart).map(|_| ()).map_err(xerr))?;
        Ok(rb_self)
    }

    fn name(&self) -> Result<String, Error> {
        self.with_ws(|ws| Ok(ws.name()))
    }

    fn next_row(&self) -> u32 {
        self.next_row.get()
    }
}

#[magnus::init]
fn init(ruby: &Ruby) -> Result<(), Error> {
    let module = ruby.define_module("FastXlsx")?;

    let wb = module.define_class("Workbook", ruby.class_object())?;
    wb.define_singleton_method("_new", function!(Workbook::new, 1))?;
    wb.define_method("_add_worksheet", method!(Workbook::add_worksheet, 1))?;
    wb.define_method("to_xlsx", method!(Workbook::to_xlsx, 0))?;
    wb.define_method("save", method!(Workbook::save, 1))?;
    wb.define_method("_set_properties", method!(Workbook::set_properties, 1))?;

    let ws = module.define_class("Worksheet", ruby.class_object())?;
    ws.define_method("_write", method!(Worksheet::write, 4))?;
    ws.define_method("_append", method!(Worksheet::append, 2))?;
    ws.define_method("_set_column_width", method!(Worksheet::set_column_width, 3))?;
    ws.define_method(
        "_set_column_format",
        method!(Worksheet::set_column_format, 3),
    )?;
    ws.define_method("_autofit", method!(Worksheet::autofit, 0))?;
    ws.define_method("autofilter", method!(Worksheet::autofilter, 4))?;
    ws.define_method("name", method!(Worksheet::name, 0))?;
    ws.define_method("_write_comment", method!(Worksheet::write_comment, 4))?;
    ws.define_method("_insert_image", method!(Worksheet::insert_image, 4))?;
    ws.define_method("_insert_chart", method!(Worksheet::insert_chart, 3))?;
    ws.define_method(
        "_conditional_format",
        method!(Worksheet::conditional_format, 5),
    )?;
    ws.define_method("_data_validation", method!(Worksheet::data_validation, 5))?;
    ws.define_method("freeze_panes", method!(Worksheet::freeze_panes, 2))?;
    ws.define_method("set_row_height", method!(Worksheet::set_row_height, 2))?;
    ws.define_method("set_page_breaks", method!(Worksheet::set_page_breaks, 1))?;
    ws.define_method("_set_header", method!(Worksheet::set_header, 1))?;
    ws.define_method("_set_footer", method!(Worksheet::set_footer, 1))?;
    ws.define_method("_set_margins", method!(Worksheet::set_margins, 6))?;
    ws.define_method(
        "set_vertical_page_breaks",
        method!(Worksheet::set_vertical_page_breaks, 1),
    )?;
    ws.define_method("_merge_range", method!(Worksheet::merge_range, 6))?;

    let format = module.define_class("Format", ruby.class_object())?;
    format.define_singleton_method("_new", function!(Format::new, 1))?;
    ws.define_method("<<", method!(Worksheet::push, 1))?;
    ws.define_method("concat", method!(Worksheet::concat, 1))?;
    ws.define_method("next_row", method!(Worksheet::next_row, 0))?;
    Ok(())
}
