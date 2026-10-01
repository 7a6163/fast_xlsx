use std::cell::{Cell, RefCell};
use std::sync::atomic::{AtomicUsize, Ordering};
use std::sync::LazyLock;
use std::sync::{Arc, Mutex};

use magnus::{
    function, method, prelude::*, r_hash::ForEach, typed_data::Obj, value::Lazy, Error,
    ExceptionClass, Integer, RArray, RClass, RHash, RModule, RString, Ruby, Symbol, TryConvert,
    Value,
};
use rust_xlsxwriter::{
    Color, ConditionalFormat, ConditionalFormat2ColorScale, ConditionalFormat3ColorScale,
    ConditionalFormatCell, ConditionalFormatCellRule, ConditionalFormatDataBar,
    ConditionalFormatFormula, ConditionalFormatText, ConditionalFormatTextRule,
    ConditionalFormatValue, DataValidation, DataValidationRule, FormatAlign, FormatBorder,
    FormatScript, FormatUnderline, IgnoreError, IntoDataValidationValue, IntoExcelData, XlsxError,
};

use rust_xlsxwriter::{
    Chart, ChartType, DocProperties, Image, Note, ProtectionOptions, Table, TableColumn,
    TableFunction, TableStyle,
};

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
static RICH_STRING: Lazy<RClass> = Lazy::new(|ruby| fast_xlsx_const(ruby, "RichString"));

// A row or column outside the sheet is a RangeError, like a negative one
// (which fails converting to an unsigned number before reaching the writer).
fn xerr(e: XlsxError) -> Error {
    let ruby = Ruby::get().unwrap();
    let class = match e {
        XlsxError::RowColumnLimitError => ruby.exception_range_error(),
        // Values that aren't allowed, like a reversed range.
        XlsxError::RowColumnOrderError | XlsxError::MergeRangeSingleCell => {
            ruby.exception_arg_error()
        }
        _ => ruby.get_inner(&ERROR),
    };
    Error::new(class, e.to_string())
}

// Excel's limits that rust_xlsxwriter doesn't check: a negative width or
// height hides the column or row, a larger one is capped or invalid.
// A size that must be a finite number above 0 (NaN or a negative one would
// be written into the file as is).
fn check_positive(ruby: &Ruby, what: &str, size: f64) -> Result<f64, Error> {
    if size.is_finite() && size > 0.0 {
        return Ok(size);
    }
    Err(Error::new(
        ruby.exception_arg_error(),
        format!("invalid {what} {size}: use a number above 0"),
    ))
}

fn check_size(ruby: &Ruby, what: &str, size: f64, max: f64) -> Result<(), Error> {
    if (0.0..=max).contains(&size) {
        return Ok(());
    }
    Err(Error::new(
        ruby.exception_arg_error(),
        format!("invalid {what} {size}: use 0..{max}"),
    ))
}

type Shared = Arc<Mutex<rust_xlsxwriter::Workbook>>;

// Runs `f` without Ruby's global lock, so other Ruby threads run meanwhile.
// `f` must not touch any Ruby object, and must release the workbook mutex
// before returning: a Ruby thread waiting on that mutex holds the lock this
// thread then needs back.
//
// Uses the "2" variant: rb_thread_call_without_gvl raises pending interrupts
// (Thread#raise, Timeout, Ctrl-C) by longjmp-ing over these Rust frames,
// which skips their destructors. This one never raises; if an interrupt is
// already pending it returns without calling `f`, which then runs with the
// lock held. Either way Ruby raises the interrupt after the method returns.
// Without an unblocking function, an interrupt waits for `f` to finish.
fn without_gvl<R>(f: impl FnOnce() -> R) -> R {
    unsafe extern "C" fn call<F: FnOnce() -> R, R>(
        data: *mut std::ffi::c_void,
    ) -> *mut std::ffi::c_void {
        let (f, result) = &mut *(data as *mut (Option<F>, Option<std::thread::Result<R>>));
        // A panic must not unwind into Ruby's C code; it is resumed below.
        *result = Some(std::panic::catch_unwind(std::panic::AssertUnwindSafe(
            f.take().unwrap(),
        )));
        std::ptr::null_mut()
    }
    fn run<F: FnOnce() -> R, R>(f: F) -> R {
        let mut data: (Option<F>, Option<std::thread::Result<R>>) = (Some(f), None);
        unsafe {
            rb_sys::rb_thread_call_without_gvl2(
                Some(call::<F, R>),
                &mut data as *mut _ as *mut std::ffi::c_void,
                None,
                std::ptr::null_mut(),
            );
        }
        let Some(result) = data.1 else {
            // An interrupt was pending, so `call` never ran.
            return (data.0.take().unwrap())();
        };
        match result {
            Ok(result) => result,
            Err(panic) => std::panic::resume_unwind(panic),
        }
    }
    run(f)
}

#[magnus::wrap(class = "FastXlsx::Workbook", free_immediately)]
struct Workbook {
    inner: Shared,
    // Index of the sheet Excel opens on (0 unless one is activated), shared
    // with the worksheets: it can't be hidden.
    active: Arc<AtomicUsize>,
    constant_memory: bool,
    low_memory: bool,
}

#[magnus::wrap(class = "FastXlsx::Worksheet", free_immediately)]
struct Worksheet {
    wb: Shared,
    index: usize,
    active: Arc<AtomicUsize>,
    // Where << / append write next.
    next_row: Cell<u32>,
    // Highest row with cells written. In :constant / :low mode rows above it
    // are on disk; rows a merge spans are held back, so merges don't count.
    last_written_row: Cell<u32>,
    // constant_memory and low_memory worksheets write finished rows to disk.
    flushes_rows: bool,
    // Formats of table columns, for cells written into a table's data rows
    // after add_table (rust_xlsxwriter only formats cells that already exist).
    // Owned copies, so they don't depend on the Ruby Format objects living on.
    table_formats: RefCell<Vec<TableColumnFormat>>,
    // Merged ranges as (first_row, first_col, last_row, last_col).
    merges: RefCell<Vec<(u32, u16, u32, u16)>>,
    // Formats given with column_format: (first, last, style).
    formatted_columns: RefCell<Vec<(u16, u16, Arc<Style>)>>,
    // Formats given with row_format: (first, last, style).
    formatted_rows: RefCell<Vec<(u32, u32, Arc<Style>)>>,
}

struct TableColumnFormat {
    rows: std::ops::RangeInclusive<u32>,
    col: u16,
    format: Arc<Style>,
}

impl Workbook {
    fn new(constant_memory: bool, low_memory: bool) -> Self {
        Workbook {
            inner: Arc::new(Mutex::new(rust_xlsxwriter::Workbook::new())),
            active: Arc::new(AtomicUsize::new(0)),
            constant_memory,
            low_memory,
        }
    }

    fn add_worksheet(&self, name: Option<String>) -> Result<Worksheet, Error> {
        let mut wb = self.inner.lock().unwrap();
        // Excel sheet names are case-insensitive; rust_xlsxwriter only notices
        // a clash when saving.
        let taken: Vec<String> = wb
            .worksheets()
            .iter()
            .map(|ws| ws.name().to_lowercase())
            .collect();
        let name = match name {
            Some(name) => {
                // Validate the name before adding: rust_xlsxwriter adds the
                // sheet first, so a bad name would leave a "SheetN" behind.
                rust_xlsxwriter::Worksheet::new()
                    .set_name(&name)
                    .map_err(xerr)?;
                if taken.contains(&name.to_lowercase()) {
                    let ruby = Ruby::get().unwrap();
                    return Err(Error::new(
                        ruby.get_inner(&ERROR),
                        format!("a worksheet named {name:?} already exists (names ignore case)"),
                    ));
                }
                name
            }
            // Like Excel, the first free "SheetN"; rust_xlsxwriter's default
            // counts sheets, which can clash with a name given earlier.
            None => (1..)
                .map(|n| format!("Sheet{n}"))
                .find(|n| !taken.contains(&n.to_lowercase()))
                .unwrap(),
        };
        let ws = if self.low_memory {
            wb.add_worksheet_with_low_memory()
        } else if self.constant_memory {
            wb.add_worksheet_with_constant_memory()
        } else {
            wb.add_worksheet()
        };
        ws.set_name(name).map_err(xerr)?;
        Ok(Worksheet {
            wb: self.inner.clone(),
            index: wb.worksheets().len() - 1,
            active: self.active.clone(),
            next_row: Cell::new(0),
            last_written_row: Cell::new(0),
            flushes_rows: self.constant_memory || self.low_memory,
            table_formats: RefCell::new(Vec::new()),
            merges: RefCell::new(Vec::new()),
            formatted_columns: RefCell::new(Vec::new()),
            formatted_rows: RefCell::new(Vec::new()),
        })
    }

    // Saving (XML and compression) can take a while, so it runs without
    // Ruby's global lock. The mutex guard is dropped inside the closure.
    fn to_xlsx(ruby: &Ruby, rb_self: &Self) -> Result<RString, Error> {
        let inner = &rb_self.inner;
        let buf = without_gvl(|| inner.lock().unwrap().save_to_buffer()).map_err(xerr)?;
        Ok(ruby.str_from_slice(&buf))
    }

    fn save(&self, path: String) -> Result<(), Error> {
        let inner = &self.inner;
        without_gvl(|| inner.lock().unwrap().save(path).map(|_| ())).map_err(xerr)
    }

    // "Name" for the whole workbook, "Sheet1!Name" for one sheet. Duplicate
    // names and unknown sheets are reported when saving, since sheets can be
    // added after the name.
    fn define_name(rb_self: Obj<Self>, name: String, formula: String) -> Result<Obj<Self>, Error> {
        rb_self
            .inner
            .lock()
            .unwrap()
            .define_name(name, &formula)
            .map_err(xerr)?;
        Ok(rb_self)
    }

    fn set_properties(ruby: &Ruby, rb_self: &Self, fields: RHash) -> Result<(), Error> {
        let mut props = DocProperties::new();
        fields.foreach(|key: Symbol, value: String| {
            let taken = std::mem::take(&mut props);
            props = match &*key.name()? {
                "title" => taken.set_title(value),
                "subject" => taken.set_subject(value),
                "author" => taken.set_author(value),
                "manager" => taken.set_manager(value),
                "company" => taken.set_company(value),
                "category" => taken.set_category(value),
                "keywords" => taken.set_keywords(value),
                "comments" => taken.set_comment(value),
                "status" => taken.set_status(value),
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

// Excel serial date from days (and fraction) since 1899-12-30. Excel's 1900
// date system counts a non-existent 1900-02-29 (serial 60), so real dates
// before 1900-03-01 are one lower; it has no dates before 1900-01-01.
fn excel_serial(days: f64) -> Result<f64, Error> {
    if days < 2.0 {
        let ruby = Ruby::get().unwrap();
        return Err(Error::new(
            ruby.exception_arg_error(),
            "Excel cannot represent dates before 1900-01-01",
        ));
    }
    Ok(if days < 61.0 { days - 1.0 } else { days })
}

// Time: seconds since 1970-01-01 (serial day 25569), in its own offset.
fn excel_time(v: Value) -> Result<f64, Error> {
    let secs: f64 = v.funcall("to_f", ())?;
    let offset: i64 = v.funcall("utc_offset", ())?;
    excel_serial((secs + offset as f64) / 86400.0 + 25569.0)
}

// Date / DateTime: Julian day and day fraction are both in the object's own
// offset. JD 2415019 is 1899-12-30.
fn excel_date(v: Value) -> Result<f64, Error> {
    let jd: i64 = v.funcall("jd", ())?;
    let fraction: f64 = f64::try_convert(v.funcall("day_fraction", ())?)?;
    excel_serial((jd - 2_415_019) as f64 + fraction)
}

fn emit<T: IntoExcelData>(
    ws: &mut rust_xlsxwriter::Worksheet,
    row: u32,
    col: u16,
    data: T,
    format: Option<&rust_xlsxwriter::Format>,
) -> Result<(), Error> {
    match format {
        Some(f) => ws.write_with_format(row, col, data, f),
        None => ws.write(row, col, data),
    }
    .map(|_| ())
    .map_err(xerr)
}

// Excel's limit on the text in a cell.
const MAX_CHARS: usize = 32_767;
// Excel's column and row counts.
const MAX_COLS: usize = 16_384;
const MAX_ROWS: u32 = 1_048_576;

// A cell value converted from Ruby. Converting may run Ruby code (to_s, jd,
// url, ...), so it happens before the workbook lock is taken: Ruby code that
// touches the same workbook, or another thread, would otherwise deadlock on it.
// Writing a CellValue runs no Ruby code.
// It holds Rust-owned copies, never Ruby objects: Ruby code run while the
// rest of the row converts could change or drop them, and the GC does not
// see references kept in a Rust Vec.
enum CellValue {
    Empty,
    Text(String),
    Number(f64),
    // An Excel serial date; true when it has a time of day (Time, DateTime).
    Date(f64, bool),
    Bool(bool),
    Formula(String),
    Url(String, Option<String>),
    // RichString segments; None means the default font.
    Rich(Vec<(Option<Arc<Style>>, String)>),
}

impl CellValue {
    fn from_ruby(ruby: &Ruby, v: Value) -> Result<Self, Error> {
        let value = if v.is_nil() {
            CellValue::Empty
        } else if let Some(s) = RString::from_value(v) {
            // Other encodings (e.g. Windows-1252 / Big5 from a legacy CSV) are
            // converted; invalid bytes, or binary non-ASCII, raise.
            CellValue::Text(s.to_string()?)
        } else if v.is_kind_of(ruby.class_numeric()) {
            CellValue::Number(f64::try_convert(v)?)
        } else if v.is_kind_of(ruby.class_time()) {
            CellValue::Date(excel_time(v)?, true)
        } else if v.is_kind_of(ruby.class_true_class()) || v.is_kind_of(ruby.class_false_class()) {
            CellValue::Bool(v.to_bool())
        } else if v.is_kind_of(ruby.get_inner(&FORMULA)) {
            CellValue::Formula(v.funcall("expression", ())?)
        } else if v.is_kind_of(ruby.get_inner(&URL)) {
            CellValue::Url(v.funcall("url", ())?, v.funcall("text", ())?)
        } else if v.is_kind_of(ruby.get_inner(&RICH_STRING)) {
            let segments: RArray = v.funcall("segments", ())?;
            let mut parts = Vec::with_capacity(segments.len());
            each_entry(segments, |_, segment| {
                let (text, seg_format): (String, Value) = TryConvert::try_convert(segment)?;
                let seg_format = Option::<&Format>::try_convert(seg_format)?;
                parts.push((seg_format.map(|f| f.0.clone()), text));
                Ok(())
            })?;
            CellValue::Rich(parts)
        } else if v.respond_to("jd", false)? {
            // Date has no #hour; DateTime does.
            CellValue::Date(excel_date(v)?, v.respond_to("hour", false)?)
        } else {
            CellValue::Text(v.funcall("to_s", ())?)
        };
        value.check()?;
        Ok(value)
    }

    // Some(has a time of day) for date cells.
    fn date_kind(&self) -> Option<bool> {
        match self {
            CellValue::Date(_, time) => Some(*time),
            _ => None,
        }
    }

    // Rejects what the writer would, so a row or merge fails before anything
    // is written. Numbers, booleans and formulas only fail on a bad row/column.
    fn check(&self) -> Result<(), Error> {
        thread_local! {
            // Reused: every check writes to A1, replacing the one before.
            static SCRATCH: RefCell<rust_xlsxwriter::Worksheet> =
                RefCell::new(rust_xlsxwriter::Worksheet::new());
        }
        match self {
            CellValue::Text(s) if s.chars().count() > MAX_CHARS => {
                Err(xerr(XlsxError::MaxStringLengthExceeded))
            }
            // URLs and rich strings have more rules; let the writer apply them.
            CellValue::Url(..) | CellValue::Rich(_) => {
                SCRATCH.with_borrow_mut(|ws| self.write(ws, 0, 0, None))
            }
            _ => Ok(()),
        }
    }

    fn write(
        &self,
        ws: &mut rust_xlsxwriter::Worksheet,
        row: u32,
        col: u16,
        format: Option<&rust_xlsxwriter::Format>,
    ) -> Result<(), Error> {
        match self {
            // nil writes nothing, unless it has a format (e.g. a border).
            CellValue::Empty => match format {
                Some(f) => ws.write_blank(row, col, f).map(|_| ()).map_err(xerr),
                None => Ok(()),
            },
            CellValue::Text(s) => emit(ws, row, col, s.as_str(), format),
            CellValue::Number(n) | CellValue::Date(n, _) => emit(ws, row, col, *n, format),
            CellValue::Bool(b) => emit(ws, row, col, *b, format),
            CellValue::Formula(f) => emit(
                ws,
                row,
                col,
                rust_xlsxwriter::Formula::new(f.as_str()),
                format,
            ),
            CellValue::Url(url, text) => {
                let mut link = rust_xlsxwriter::Url::new(url.as_str());
                if let Some(text) = text {
                    link = link.set_text(text.as_str());
                }
                emit(ws, row, col, link, format)
            }
            CellValue::Rich(parts) => {
                let default = rust_xlsxwriter::Format::default();
                let rich: Vec<(&rust_xlsxwriter::Format, &str)> = parts
                    .iter()
                    .map(|(f, text)| (f.as_deref().map_or(&default, |s| &s.format), text.as_str()))
                    .collect();
                match format {
                    Some(f) => ws.write_rich_string_with_format(row, col, &rich, f),
                    None => ws.write_rich_string(row, col, &rich),
                }
                .map(|_| ())
                .map_err(xerr)
            }
        }
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
            ruby.get_inner(&ERROR),
            format!("{what} is {len} characters; Excel allows 255"),
        ));
    }
    Ok(())
}

// Option `name` of a Ruby options Hash as given; nil when missing. For
// choice(), which reports an invalid (or missing) value itself.
fn raw_opt(ruby: &Ruby, options: RHash, name: &str) -> Value {
    options
        .get(ruby.to_symbol(name))
        .unwrap_or_else(|| ruby.qnil().as_value())
}

// Option `name` of a Ruby options Hash, converted; missing or nil is None.
fn opt<T: TryConvert>(ruby: &Ruby, options: RHash, name: &str) -> Result<Option<T>, Error> {
    Option::<T>::try_convert(raw_opt(ruby, options, name))
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

// :none, :light1..:light21, :medium1..:medium28, :dark1..:dark11.
const TABLE_STYLES: &[(&str, TableStyle)] = &[
    ("none", TableStyle::None),
    ("light1", TableStyle::Light1),
    ("light2", TableStyle::Light2),
    ("light3", TableStyle::Light3),
    ("light4", TableStyle::Light4),
    ("light5", TableStyle::Light5),
    ("light6", TableStyle::Light6),
    ("light7", TableStyle::Light7),
    ("light8", TableStyle::Light8),
    ("light9", TableStyle::Light9),
    ("light10", TableStyle::Light10),
    ("light11", TableStyle::Light11),
    ("light12", TableStyle::Light12),
    ("light13", TableStyle::Light13),
    ("light14", TableStyle::Light14),
    ("light15", TableStyle::Light15),
    ("light16", TableStyle::Light16),
    ("light17", TableStyle::Light17),
    ("light18", TableStyle::Light18),
    ("light19", TableStyle::Light19),
    ("light20", TableStyle::Light20),
    ("light21", TableStyle::Light21),
    ("medium1", TableStyle::Medium1),
    ("medium2", TableStyle::Medium2),
    ("medium3", TableStyle::Medium3),
    ("medium4", TableStyle::Medium4),
    ("medium5", TableStyle::Medium5),
    ("medium6", TableStyle::Medium6),
    ("medium7", TableStyle::Medium7),
    ("medium8", TableStyle::Medium8),
    ("medium9", TableStyle::Medium9),
    ("medium10", TableStyle::Medium10),
    ("medium11", TableStyle::Medium11),
    ("medium12", TableStyle::Medium12),
    ("medium13", TableStyle::Medium13),
    ("medium14", TableStyle::Medium14),
    ("medium15", TableStyle::Medium15),
    ("medium16", TableStyle::Medium16),
    ("medium17", TableStyle::Medium17),
    ("medium18", TableStyle::Medium18),
    ("medium19", TableStyle::Medium19),
    ("medium20", TableStyle::Medium20),
    ("medium21", TableStyle::Medium21),
    ("medium22", TableStyle::Medium22),
    ("medium23", TableStyle::Medium23),
    ("medium24", TableStyle::Medium24),
    ("medium25", TableStyle::Medium25),
    ("medium26", TableStyle::Medium26),
    ("medium27", TableStyle::Medium27),
    ("medium28", TableStyle::Medium28),
    ("dark1", TableStyle::Dark1),
    ("dark2", TableStyle::Dark2),
    ("dark3", TableStyle::Dark3),
    ("dark4", TableStyle::Dark4),
    ("dark5", TableStyle::Dark5),
    ("dark6", TableStyle::Dark6),
    ("dark7", TableStyle::Dark7),
    ("dark8", TableStyle::Dark8),
    ("dark9", TableStyle::Dark9),
    ("dark10", TableStyle::Dark10),
    ("dark11", TableStyle::Dark11),
];

const TABLE_TOTALS: &[(&str, TableFunction)] = &[
    ("sum", TableFunction::Sum),
    ("average", TableFunction::Average),
    ("count", TableFunction::Count),
    ("count_numbers", TableFunction::CountNumbers),
    ("max", TableFunction::Max),
    ("min", TableFunction::Min),
    ("std_dev", TableFunction::StdDev),
    ("var", TableFunction::Var),
];

// A table column: a header String, or { header:, total:, total_label:, format: }.
// Also returns a copy of the column's format, if any, for later writes.
fn table_column(ruby: &Ruby, v: Value) -> Result<(TableColumn, Option<Arc<Style>>), Error> {
    let Some(spec) = RHash::from_value(v) else {
        return Ok((TableColumn::new().set_header(String::try_convert(v)?), None));
    };
    check_keys(
        ruby,
        spec,
        &["header", "total", "total_label", "format"],
        "table column",
    )?;
    let header = String::try_convert(raw_opt(ruby, spec, "header"))?;
    let mut column = TableColumn::new().set_header(header);
    if let Some(total) = opt::<Value>(ruby, spec, "total")? {
        column = column.set_total_function(choice(ruby, "table total", total, TABLE_TOTALS)?);
    }
    if let Some(label) = opt::<String>(ruby, spec, "total_label")? {
        column = column.set_total_label(label);
    }
    let format = opt::<&Format>(ruby, spec, "format")?.map(|f| f.0.clone());
    if let Some(f) = &format {
        column = column.set_format(&f.format);
    }
    Ok((column, format))
}

// Actions users may still take on a protected sheet (selecting cells is
// always allowed).
type ProtectionFlag = fn(&mut ProtectionOptions) -> &mut bool;
const PROTECTION_ALLOW: &[(&str, ProtectionFlag)] = &[
    ("format_cells", |o| &mut o.format_cells),
    ("format_columns", |o| &mut o.format_columns),
    ("format_rows", |o| &mut o.format_rows),
    ("insert_columns", |o| &mut o.insert_columns),
    ("insert_rows", |o| &mut o.insert_rows),
    ("insert_links", |o| &mut o.insert_links),
    ("delete_columns", |o| &mut o.delete_columns),
    ("delete_rows", |o| &mut o.delete_rows),
    ("sort", |o| &mut o.sort),
    ("use_autofilter", |o| &mut o.use_autofilter),
    ("use_pivot_tables", |o| &mut o.use_pivot_tables),
    ("edit_scenarios", |o| &mut o.edit_scenarios),
    ("edit_objects", |o| &mut o.edit_objects),
];

// Not IgnoreError::TwoDigitTextYear: rust_xlsxwriter 0.99 writes it as
// "TwoDigitTextYear", not the schema's "twoDigitTextYear", which Excel would
// flag as damaged content. Add it once that is fixed upstream.
const IGNORE_ERRORS: &[(&str, IgnoreError)] = &[
    ("number_stored_as_text", IgnoreError::NumberStoredAsText),
    ("formula_error", IgnoreError::FormulaError),
    ("formula_differs", IgnoreError::FormulaDiffers),
    (
        "formula_refers_to_empty_cells",
        IgnoreError::FormulaRefersToEmptyCells,
    ),
    ("formula_omits_cells", IgnoreError::FormulaOmitsCells),
    ("data_validation_error", IgnoreError::DataValidationError),
    (
        "unlocked_cells_with_formula",
        IgnoreError::UnlockedCellsWithFormula,
    ),
    (
        "inconsistent_column_formula",
        IgnoreError::InconsistentColumnFormula,
    ),
];

// Excel's paper size codes; page_setup also takes the number itself.
const PAPER_SIZES: &[(&str, u8)] = &[
    ("letter", 1),
    ("tabloid", 3),
    ("legal", 5),
    ("a3", 8),
    ("a4", 9),
    ("a5", 11),
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
struct Format(Arc<Style>);

// A Format as written: the format itself, and for one without a num_format,
// copies with the default date formats added, used for date cells (Excel
// shows a date with no number format as its serial number).
struct Style {
    format: rust_xlsxwriter::Format,
    date: Option<rust_xlsxwriter::Format>,
    date_time: Option<rust_xlsxwriter::Format>,
}

const DATE_FORMAT: &str = "yyyy-mm-dd";
const DATE_TIME_FORMAT: &str = "yyyy-mm-dd hh:mm:ss";

impl Style {
    fn new(format: rust_xlsxwriter::Format, has_num_format: bool) -> Self {
        let with =
            |num_format: &str| (!has_num_format).then(|| format.clone().set_num_format(num_format));
        Style {
            date: with(DATE_FORMAT),
            date_time: with(DATE_TIME_FORMAT),
            format,
        }
    }

    // The date variant for a date cell (`time`: it has a time of day), if
    // this format has no num_format of its own.
    fn for_date(&self, time: bool) -> Option<&rust_xlsxwriter::Format> {
        if time {
            self.date_time.as_ref()
        } else {
            self.date.as_ref()
        }
    }

    // What a value is written with: `date` is Some(has time) for date cells.
    fn pick(&self, date: Option<bool>) -> &rust_xlsxwriter::Format {
        date.and_then(|time| self.for_date(time))
            .unwrap_or(&self.format)
    }
}

impl Format {
    fn new(ruby: &Ruby, options: RHash) -> Result<Self, Error> {
        let mut f = rust_xlsxwriter::Format::new();
        let mut has_num_format = false;
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
                // For protected sheets: locked: false leaves a cell editable,
                // hidden: true hides its formula.
                "locked" if value.to_bool() => taken.set_locked(),
                "locked" => taken.set_unlocked(),
                "hidden" if value.to_bool() => taken.set_hidden(),
                "border_color" => taken.set_border_color(color(ruby, value)?),
                "num_format" => {
                    has_num_format = true;
                    taken.set_num_format(String::try_convert(value)?)
                }
                "font_size" => taken.set_font_size(check_positive(
                    ruby,
                    "font_size",
                    f64::try_convert(value)?,
                )?),
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
                "bold" | "italic" | "underline" | "text_wrap" | "strikeout" | "shrink"
                | "hidden" => taken,
                other => {
                    return Err(Error::new(
                        ruby.exception_arg_error(),
                        format!("unknown format option: {other}"),
                    ))
                }
            };
            Ok(ForEach::Continue)
        })?;
        Ok(Format(Arc::new(Style::new(f, has_num_format))))
    }
}

// Visits each element by index. RArray::into_iter dups the array first, and
// dup'ing an Array longer than 3 elements turns the caller's array into a
// shared root: one extra live Ruby object per row for as long as the data lives.
// Reading the length each step also stays correct if Ruby code run by `f`
// (e.g. a to_s) changes the array.
fn each_entry(
    ary: RArray,
    mut f: impl FnMut(usize, Value) -> Result<(), Error>,
) -> Result<(), Error> {
    let mut i = 0;
    while i < ary.len() {
        f(i, ary.entry(i as isize)?)?;
        i += 1;
    }
    Ok(())
}

// A row's format: one Format (or nil) for every cell, or an Array with one per cell.
enum RowFormat {
    Same(Option<Arc<Style>>),
    PerCell(RArray),
}

impl RowFormat {
    fn from_value(v: Value) -> Result<Self, Error> {
        match RArray::from_value(v) {
            Some(formats) => Ok(RowFormat::PerCell(formats)),
            None => Ok(RowFormat::Same(
                Option::<&Format>::try_convert(v)?.map(|f| f.0.clone()),
            )),
        }
    }

    fn at(&self, col: usize) -> Result<Option<Arc<Style>>, Error> {
        match self {
            RowFormat::Same(f) => Ok(f.clone()),
            RowFormat::PerCell(formats) => Ok(Option::<&Format>::try_convert(
                formats.entry::<Value>(col as isize)?,
            )?
            .map(|f| f.0.clone())),
        }
    }
}

impl Worksheet {
    fn advance(&self, row: u32) {
        self.next_row.set(self.next_row.get().max(row + 1));
    }

    fn note_written(&self, row: u32) {
        self.last_written_row
            .set(self.last_written_row.get().max(row));
    }

    // rust_xlsxwriter silently drops writes to rows it has already flushed.
    fn check_not_flushed(&self, ruby: &Ruby, row: u32) -> Result<(), Error> {
        if self.flushes_rows && row < self.last_written_row.get() {
            return Err(Error::new(
                ruby.get_inner(&ERROR),
                format!(
                    "row {row} was already written to disk (constant_memory / low_memory mode)"
                ),
            ));
        }
        Ok(())
    }

    fn with_ws<T>(
        &self,
        f: impl FnOnce(&mut rust_xlsxwriter::Worksheet) -> Result<T, Error>,
    ) -> Result<T, Error> {
        let mut wb = self.wb.lock().unwrap();
        f(wb.worksheet_from_index(self.index).map_err(xerr)?)
    }

    // The format a value is written with: its own, else its table column's.
    // A date cell takes that format's date variant (when it has no
    // num_format); with neither, its row_format's, else its column_format's
    // date variant (rust_xlsxwriter's order), else the default date format.
    // Other cells without one get None, so rust_xlsxwriter applies any row or
    // column format itself.
    fn format_for<'a>(
        tables: &'a [TableColumnFormat],
        rows: &'a [(u32, u32, Arc<Style>)],
        columns: &'a [(u16, u16, Arc<Style>)],
        row: u32,
        col: u16,
        own: Option<&'a Style>,
        value: &CellValue,
    ) -> Option<&'a rust_xlsxwriter::Format> {
        static DEFAULTS: LazyLock<Style> =
            LazyLock::new(|| Style::new(rust_xlsxwriter::Format::new(), false));
        let date = value.date_kind();
        let style = own.or_else(|| {
            tables
                .iter()
                .rev()
                .find(|t| t.col == col && t.rows.contains(&row))
                .map(|t| &*t.format)
        });
        if let Some(style) = style {
            return Some(style.pick(date));
        }
        let time = date?;
        let row_style = rows
            .iter()
            .rev()
            .find(|(first, last, _)| (*first..=*last).contains(&row))
            .map(|(_, _, style)| style);
        let column_style = || {
            columns
                .iter()
                .rev()
                .find(|(first, last, _)| (*first..=*last).contains(&col))
                .map(|(_, _, style)| style)
        };
        match row_style.or_else(column_style) {
            // None when it has a num_format: rust_xlsxwriter applies it.
            Some(style) => style.for_date(time),
            None => DEFAULTS.for_date(time),
        }
    }

    fn write_row(
        &self,
        ruby: &Ruby,
        row: u32,
        cells: RArray,
        format: &RowFormat,
    ) -> Result<(), Error> {
        // Convert the whole row first: no lock is held while Ruby code runs,
        // and a value that fails to convert leaves the row unwritten.
        let mut values = Vec::with_capacity(cells.len());
        each_entry(cells, |i, v| {
            if i >= MAX_COLS {
                return Err(Error::new(
                    ruby.exception_range_error(),
                    format!("too many columns: Excel allows {MAX_COLS}"),
                ));
            }
            let col = i as u16;
            values.push((col, CellValue::from_ruby(ruby, v)?, format.at(i)?));
            Ok(())
        })?;
        let tables = self.table_formats.borrow();
        let rows = self.formatted_rows.borrow();
        let columns = self.formatted_columns.borrow();
        self.with_ws(|ws| {
            // Values were checked when converted, so only a bad row number
            // fails here, and it fails on the first cell.
            for (col, value, format) in &values {
                let format = Self::format_for(
                    &tables,
                    &rows,
                    &columns,
                    row,
                    *col,
                    format.as_deref(),
                    value,
                );
                value.write(ws, row, *col, format)?;
            }
            Ok(())
        })?;
        self.advance(row);
        self.note_written(row);
        Ok(())
    }

    // Worksheet#write. The (row, col, value, format = nil) form is handled
    // here rather than in a Ruby wrapper, since it runs once per cell; the
    // rest ("B2", or a wrong argument count) goes to Ruby's _write_ref.
    fn write_any(ruby: &Ruby, rb_self: Obj<Self>, args: &[Value]) -> Result<Obj<Self>, Error> {
        let index = |i: usize| args.get(i).and_then(|v| Integer::from_value(*v));
        if let (3 | 4, Some(row), Some(col)) = (args.len(), index(0), index(1)) {
            let format = match args.get(3) {
                Some(f) => Option::<&Format>::try_convert(*f)?,
                None => None,
            };
            Self::write(
                ruby,
                &rb_self,
                row.to_u32()?,
                col.to_u16()?,
                args[2],
                format,
            )?;
            return Ok(rb_self);
        }
        rb_self.funcall("_write_ref", args)
    }

    fn write(
        ruby: &Ruby,
        rb_self: &Self,
        row: u32,
        col: u16,
        v: Value,
        format: Option<&Format>,
    ) -> Result<(), Error> {
        rb_self.check_not_flushed(ruby, row)?;
        let value = CellValue::from_ruby(ruby, v)?;
        let tables = rb_self.table_formats.borrow();
        let rows = rb_self.formatted_rows.borrow();
        let columns = rb_self.formatted_columns.borrow();
        rb_self.with_ws(|ws| {
            value.write(
                ws,
                row,
                col,
                Self::format_for(
                    &tables,
                    &rows,
                    &columns,
                    row,
                    col,
                    format.map(|f| &*f.0),
                    &value,
                ),
            )
        })?;
        rb_self.advance(row);
        rb_self.note_written(row);
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
        each_entry(rows, |_, r| {
            let cells = RArray::try_convert(r)?;
            rb_self.write_row(ruby, rb_self.next_row.get(), cells, &RowFormat::Same(None))
        })?;
        Ok(rb_self)
    }

    fn set_column_width(
        ruby: &Ruby,
        rb_self: &Self,
        first: u16,
        last: u16,
        width: f64,
    ) -> Result<(), Error> {
        check_size(ruby, "column width", width, 255.0)?;
        rb_self.with_ws(|ws| {
            ws.set_column_range_width(first, last, width)
                .map(|_| ())
                .map_err(xerr)
        })
    }

    // Row options are written with the row, so rows on disk ignore them.
    fn hide_rows(
        ruby: &Ruby,
        rb_self: Obj<Self>,
        first: u32,
        last: u32,
    ) -> Result<Obj<Self>, Error> {
        rb_self.check_not_flushed(ruby, first)?;
        rb_self.with_ws(|ws| {
            for row in first..=last {
                ws.set_row_hidden(row).map_err(xerr)?;
            }
            Ok(())
        })?;
        Ok(rb_self)
    }

    fn selection(
        rb_self: Obj<Self>,
        fr: u32,
        fc: u16,
        lr: u32,
        lc: u16,
    ) -> Result<Obj<Self>, Error> {
        rb_self.with_ws(|ws| ws.set_selection(fr, fc, lr, lc).map(|_| ()).map_err(xerr))?;
        Ok(rb_self)
    }

    fn top_left_cell(rb_self: Obj<Self>, row: u32, col: u16) -> Result<Obj<Self>, Error> {
        rb_self.with_ws(|ws| ws.set_top_left_cell(row, col).map(|_| ()).map_err(xerr))?;
        Ok(rb_self)
    }

    #[allow(clippy::too_many_arguments)]
    fn ignore_error(
        ruby: &Ruby,
        rb_self: Obj<Self>,
        fr: u32,
        fc: u16,
        lr: u32,
        lc: u16,
        kind: Value,
    ) -> Result<Obj<Self>, Error> {
        let error = choice(ruby, "error to ignore", kind, IGNORE_ERRORS)?;
        rb_self.with_ws(|ws| {
            ws.ignore_error_range(fr, fc, lr, lc, error)
                .map(|_| ())
                .map_err(xerr)
        })?;
        Ok(rb_self)
    }

    #[allow(clippy::too_many_arguments)]
    fn unprotect_range(
        rb_self: Obj<Self>,
        fr: u32,
        fc: u16,
        lr: u32,
        lc: u16,
        name: Option<String>,
        password: Option<String>,
    ) -> Result<Obj<Self>, Error> {
        rb_self.with_ws(|ws| {
            ws.unprotect_range_with_options(
                fr,
                fc,
                lr,
                lc,
                name.as_deref().unwrap_or(""),
                password.as_deref().unwrap_or(""),
            )
            .map(|_| ())
            .map_err(xerr)
        })?;
        Ok(rb_self)
    }

    fn hide_columns(rb_self: Obj<Self>, first: u16, last: u16) -> Result<Obj<Self>, Error> {
        rb_self.with_ws(|ws| {
            ws.set_column_range_hidden(first, last)
                .map(|_| ())
                .map_err(xerr)
        })?;
        Ok(rb_self)
    }

    fn set_row_format(
        ruby: &Ruby,
        rb_self: &Self,
        first: u32,
        last: u32,
        format: &Format,
    ) -> Result<(), Error> {
        rb_self.check_not_flushed(ruby, first)?;
        rb_self.with_ws(|ws| {
            for row in first..=last {
                ws.set_row_format(row, &format.0.format).map_err(xerr)?;
            }
            Ok(())
        })?;
        rb_self
            .formatted_rows
            .borrow_mut()
            .push((first, last, format.0.clone()));
        Ok(())
    }

    fn default_row_height(
        ruby: &Ruby,
        rb_self: Obj<Self>,
        height: f64,
    ) -> Result<Obj<Self>, Error> {
        check_size(ruby, "row height", height, 409.0)?;
        rb_self.with_ws(|ws| {
            ws.set_default_row_height(height);
            Ok(())
        })?;
        Ok(rb_self)
    }

    fn set_column_format(&self, first: u16, last: u16, format: &Format) -> Result<(), Error> {
        self.with_ws(|ws| {
            ws.set_column_range_format(first, last, &format.0.format)
                .map(|_| ())
                .map_err(xerr)
        })?;
        self.formatted_columns
            .borrow_mut()
            .push((first, last, format.0.clone()));
        Ok(())
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

    fn set_row_height(
        ruby: &Ruby,
        rb_self: Obj<Self>,
        row: u32,
        height: f64,
    ) -> Result<Obj<Self>, Error> {
        check_size(ruby, "row height", height, 409.0)?;
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

    // Inches; None keeps the current margin.
    #[allow(clippy::too_many_arguments)]
    fn set_margins(
        ruby: &Ruby,
        rb_self: &Self,
        left: Option<f64>,
        right: Option<f64>,
        top: Option<f64>,
        bottom: Option<f64>,
        header: Option<f64>,
        footer: Option<f64>,
    ) -> Result<(), Error> {
        // rust_xlsxwriter keeps a margin given as a negative number.
        let inches = |margin: Option<f64>| -> Result<f64, Error> {
            match margin {
                None => Ok(-1.0),
                Some(m) if m.is_finite() && m >= 0.0 => Ok(m),
                Some(m) => Err(Error::new(
                    ruby.exception_arg_error(),
                    format!("invalid margin {m}: use 0 or more inches"),
                )),
            }
        };
        let (left, right, top) = (inches(left)?, inches(right)?, inches(top)?);
        let (bottom, header, footer) = (inches(bottom)?, inches(header)?, inches(footer)?);
        rb_self.with_ws(|ws| {
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
        rb_self.check_not_flushed(ruby, first_row)?;
        // rust_xlsxwriter blanks the range before it notices an overlap, which
        // wipes the earlier merge's value, so check first.
        // A range rust_xlsxwriter rejects anyway (reversed, too big) is left
        // to it, so its error names the real problem.
        // ponytail: linear scan; sheets have few merges.
        let range = (first_row, first_col, last_row, last_col);
        let valid = first_row <= last_row
            && first_col <= last_col
            && last_row < MAX_ROWS
            && usize::from(last_col) < MAX_COLS;
        let overlap = rb_self
            .merges
            .borrow()
            .iter()
            .copied()
            .find(|&(fr, fc, lr, lc)| {
                first_row <= lr && fr <= last_row && first_col <= lc && fc <= last_col
            })
            .filter(|_| valid);
        if let Some((fr, fc, lr, lc)) = overlap {
            return Err(Error::new(
                ruby.get_inner(&ERROR),
                format!(
                    "merge range {} overlaps the earlier merge {}",
                    rust_xlsxwriter::utility::cell_range(first_row, first_col, last_row, last_col),
                    rust_xlsxwriter::utility::cell_range(fr, fc, lr, lc),
                ),
            ));
        }
        let value = CellValue::from_ruby(ruby, v)?;
        let tables = rb_self.table_formats.borrow();
        let rows = rb_self.formatted_rows.borrow();
        let columns = rb_self.formatted_columns.borrow();
        let format = Self::format_for(
            &tables,
            &rows,
            &columns,
            first_row,
            first_col,
            format.map(|f| &*f.0),
            &value,
        );
        let default = rust_xlsxwriter::Format::new();
        rb_self.with_ws(|ws| {
            // The value was checked when converted, so once the range is
            // merged, writing it cannot fail.
            ws.merge_range(
                first_row,
                first_col,
                last_row,
                last_col,
                "",
                format.unwrap_or(&default),
            )
            .map_err(xerr)?;
            value.write(ws, first_row, first_col, format)
        })?;
        rb_self.merges.borrow_mut().push(range);
        rb_self.advance(last_row);
        rb_self.note_written(first_row);
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
        check_keys(
            ruby,
            options,
            &["type", "criteria", "value", "format", "colors"],
            "conditional_format",
        )?;
        let format = opt::<&Format>(ruby, options, "format")?;
        let value = raw_opt(ruby, options, "value");
        let criteria = raw_opt(ruby, options, "criteria");
        let kind = choice(
            ruby,
            "conditional format type",
            raw_opt(ruby, options, "type"),
            &[
                ("cell", "cell"),
                ("text", "text"),
                ("formula", "formula"),
                ("data_bar", "data_bar"),
                ("color_scale", "color_scale"),
            ],
        )?;
        // Built before taking the lock: building converts Ruby values.
        type AddRule = Box<dyn FnOnce(&mut rust_xlsxwriter::Worksheet) -> Result<(), XlsxError>>;
        fn rule<T: ConditionalFormat + Send + Sync + 'static>(
            (fr, fc, lr, lc): (u32, u16, u32, u16),
            cf: T,
        ) -> AddRule {
            Box::new(move |ws| ws.add_conditional_format(fr, fc, lr, lc, &cf).map(|_| ()))
        }
        let range = (fr, fc, lr, lc);
        let add_rule = match kind {
            "cell" => {
                let mut cf =
                    ConditionalFormatCell::new().set_rule(cell_rule(ruby, criteria, value)?);
                if let Some(f) = format {
                    cf = cf.set_format(&f.0.format);
                }
                rule(range, cf)
            }
            "text" => {
                let text = String::try_convert(value)?;
                let text_rule = match choice(
                    ruby,
                    "text criteria",
                    criteria,
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
                let mut cf = ConditionalFormatText::new().set_rule(text_rule);
                if let Some(f) = format {
                    cf = cf.set_format(&f.0.format);
                }
                rule(range, cf)
            }
            "formula" => {
                let mut cf =
                    ConditionalFormatFormula::new().set_rule(String::try_convert(value)?.as_str());
                if let Some(f) = format {
                    cf = cf.set_format(&f.0.format);
                }
                rule(range, cf)
            }
            "data_bar" => rule(range, ConditionalFormatDataBar::new()),
            _ => match opt::<u8>(ruby, options, "colors")?.unwrap_or(3) {
                2 => rule(range, ConditionalFormat2ColorScale::new()),
                3 => rule(range, ConditionalFormat3ColorScale::new()),
                n => {
                    return Err(Error::new(
                        ruby.exception_arg_error(),
                        format!("invalid colors {n}: use 2 or 3"),
                    ))
                }
            },
        };
        rb_self.with_ws(|ws| add_rule(ws).map_err(xerr))?;
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
        let criteria = raw_opt(ruby, options, "criteria");
        let value = raw_opt(ruby, options, "value");
        let dv = DataValidation::new();
        let mut dv = match choice(
            ruby,
            "data validation type",
            raw_opt(ruby, options, "type"),
            &[
                ("list", "list"),
                ("whole_number", "whole_number"),
                ("decimal", "decimal"),
                ("text_length", "text_length"),
            ],
        )? {
            "list" => match RArray::from_value(value) {
                // Items are shown as text in the dropdown, so numbers and
                // symbols are listed by their to_s.
                Some(items) => {
                    let mut labels = Vec::with_capacity(items.len());
                    each_entry(items, |_, item| {
                        labels.push(item.funcall::<_, _, String>("to_s", ())?);
                        Ok(())
                    })?;
                    dv.allow_list_strings(&labels).map_err(xerr)?
                }
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
            if let Some(text) = opt::<String>(ruby, options, name)? {
                dv = set(dv, text).map_err(xerr)?;
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
        check_keys(
            ruby,
            options,
            &[
                "scale", "width", "height", "x_offset", "y_offset", "alt_text",
            ],
            "insert_image",
        )?;
        // SAFETY: the bytes are copied into the Image before any Ruby code runs.
        let mut image = Image::new_from_buffer(unsafe { bytes.as_slice() }).map_err(xerr)?;
        let positive = |name: &str| -> Result<Option<f64>, Error> {
            opt::<f64>(ruby, options, name)?
                .map(|v| check_positive(ruby, name, v))
                .transpose()
        };
        let width = positive("width")?;
        let height = positive("height")?;
        if let Some(scale) = positive("scale")? {
            if width.is_some() || height.is_some() {
                return Err(Error::new(
                    ruby.exception_arg_error(),
                    "pass either scale: or width:/height:, not both",
                ));
            }
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
        if let Some(alt_text) = opt::<String>(ruby, options, "alt_text")? {
            image = image.set_alt_text(alt_text);
        }
        let x = opt::<u32>(ruby, options, "x_offset")?.unwrap_or(0);
        let y = opt::<u32>(ruby, options, "y_offset")?.unwrap_or(0);
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
        let arg_error = |msg: String| Error::new(ruby.exception_arg_error(), msg);
        check_keys(
            ruby,
            options,
            &[
                "type", "series", "title", "x_axis", "y_axis", "width", "height",
            ],
            "insert_chart",
        )?;

        let chart_type = choice(
            ruby,
            "chart type",
            raw_opt(ruby, options, "type"),
            CHART_TYPES,
        )?;
        let mut chart = Chart::new(chart_type);
        let series = RArray::try_convert(raw_opt(ruby, options, "series"))?;
        if series.is_empty() {
            return Err(arg_error(
                "series: needs at least one { values:, categories:, name: } hash".into(),
            ));
        }
        each_entry(series, |_, s| {
            let s = RHash::try_convert(s)?;
            check_keys(ruby, s, &["values", "categories", "name"], "chart series")?;
            let values = opt::<String>(ruby, s, "values")?
                .ok_or_else(|| arg_error(format!("series {} needs values:", s.inspect())))?;
            let cs = chart.add_series().set_values(values.as_str());
            if let Some(categories) = opt::<String>(ruby, s, "categories")? {
                cs.set_categories(categories.as_str());
            }
            if let Some(name) = opt::<String>(ruby, s, "name")? {
                cs.set_name(name.as_str());
            }
            Ok(())
        })?;
        if let Some(title) = opt::<String>(ruby, options, "title")? {
            chart.title().set_name(title.as_str());
        }
        if let Some(name) = opt::<String>(ruby, options, "x_axis")? {
            chart.x_axis().set_name(name.as_str());
        }
        if let Some(name) = opt::<String>(ruby, options, "y_axis")? {
            chart.y_axis().set_name(name.as_str());
        }
        if let Some(width) = opt::<u32>(ruby, options, "width")? {
            chart.set_width(width);
        }
        if let Some(height) = opt::<u32>(ruby, options, "height")? {
            chart.set_height(height);
        }
        rb_self.with_ws(|ws| ws.insert_chart(row, col, &chart).map(|_| ()).map_err(xerr))?;
        Ok(rb_self)
    }

    #[allow(clippy::too_many_arguments)]
    fn add_table(
        ruby: &Ruby,
        rb_self: Obj<Self>,
        fr: u32,
        fc: u16,
        lr: u32,
        lc: u16,
        options: RHash,
    ) -> Result<Obj<Self>, Error> {
        check_keys(
            ruby,
            options,
            &[
                "columns",
                "style",
                "name",
                "total_row",
                "banded_rows",
                "autofilter",
            ],
            "add_table",
        )?;
        // Ruby truthiness, as before: any non-nil, non-false value enables.
        let flag = |name: &str, default: bool| -> Result<bool, Error> {
            Ok(opt::<Value>(ruby, options, name)?.map_or(default, |v| v.to_bool()))
        };
        // The table writes its header row, which must not be flushed yet.
        rb_self.check_not_flushed(ruby, fr)?;
        let total_row = flag("total_row", false)?;
        let mut table = Table::new()
            .set_total_row(total_row)
            .set_banded_rows(flag("banded_rows", true)?)
            .set_autofilter(flag("autofilter", true)?);
        let mut column_formats = Vec::new();
        if let Some(specs) = opt::<RArray>(ruby, options, "columns")? {
            let width = usize::from(lc.saturating_sub(fc)) + 1;
            if specs.len() != width {
                return Err(Error::new(
                    ruby.exception_arg_error(),
                    format!(
                        "columns: has {} entries but the range is {width} columns wide",
                        specs.len()
                    ),
                ));
            }
            let mut columns = Vec::with_capacity(specs.len());
            each_entry(specs, |i, v| {
                let (column, format) = table_column(ruby, v)?;
                columns.push(column);
                if let Some(format) = format {
                    column_formats.push((fc + i as u16, format));
                }
                Ok(())
            })?;
            table = table.set_columns(&columns);
        }
        if let Some(style) = opt::<Value>(ruby, options, "style")? {
            table = table.set_style(choice(ruby, "table style", style, TABLE_STYLES)?);
        }
        if let Some(name) = opt::<String>(ruby, options, "name")? {
            table = table.set_name(name);
        }
        rb_self.with_ws(|ws| {
            ws.add_table(fr, fc, lr, lc, &table)
                .map(|_| ())
                .map_err(xerr)
        })?;
        // Data rows: under the header, above the total row.
        let data_rows = (fr + 1)..=lr.saturating_sub(u32::from(total_row));
        rb_self
            .table_formats
            .borrow_mut()
            .extend(
                column_formats
                    .into_iter()
                    .map(|(col, format)| TableColumnFormat {
                        rows: data_rows.clone(),
                        col,
                        format,
                    }),
            );
        // Continue appending under the header, so `<<` fills the table.
        rb_self.advance(fr);
        rb_self.note_written(fr);
        Ok(rb_self)
    }

    fn activate(rb_self: Obj<Self>) -> Result<Obj<Self>, Error> {
        let mut wb = rb_self.wb.lock().unwrap();
        // rust_xlsxwriter leaves sheets activated earlier selected, which
        // groups them in Excel (edits then go to all of them).
        for (i, ws) in wb.worksheets_mut().iter_mut().enumerate() {
            ws.set_active(i == rb_self.index);
            ws.set_selected(i == rb_self.index);
        }
        rb_self.active.store(rb_self.index, Ordering::Relaxed);
        Ok(rb_self)
    }

    fn hide(ruby: &Ruby, rb_self: Obj<Self>) -> Result<Obj<Self>, Error> {
        // rust_xlsxwriter would quietly unhide it when saving.
        if rb_self.active.load(Ordering::Relaxed) == rb_self.index {
            return Err(Error::new(
                ruby.get_inner(&ERROR),
                "can't hide the sheet Excel opens on (the first one unless another is activated): activate another sheet first",
            ));
        }
        rb_self.with_ws(|ws| {
            ws.set_hidden(true);
            Ok(())
        })?;
        Ok(rb_self)
    }

    fn zoom(ruby: &Ruby, rb_self: Obj<Self>, percent: i64) -> Result<Obj<Self>, Error> {
        // rust_xlsxwriter only prints a warning for these.
        if !(10..=400).contains(&percent) {
            return Err(Error::new(
                ruby.exception_arg_error(),
                format!("invalid zoom {percent}: use 10..400"),
            ));
        }
        rb_self.with_ws(|ws| {
            ws.set_zoom(percent as u16);
            Ok(())
        })?;
        Ok(rb_self)
    }

    fn tab_color(ruby: &Ruby, rb_self: Obj<Self>, value: Value) -> Result<Obj<Self>, Error> {
        let rgb = color(ruby, value)?;
        rb_self.with_ws(|ws| {
            ws.set_tab_color(rgb);
            Ok(())
        })?;
        Ok(rb_self)
    }

    fn hide_gridlines(rb_self: Obj<Self>) -> Result<Obj<Self>, Error> {
        rb_self.with_ws(|ws| {
            ws.set_screen_gridlines(false);
            Ok(())
        })?;
        Ok(rb_self)
    }

    // Ranges arrive as [first, last] / [first_row, first_col, last_row,
    // last_col], already parsed by lib/fast_xlsx.rb.
    fn page_setup(ruby: &Ruby, rb_self: &Self, options: RHash) -> Result<(), Error> {
        check_keys(
            ruby,
            options,
            &[
                "landscape",
                "paper",
                "fit_width",
                "fit_height",
                "repeat_rows",
                "repeat_columns",
                "print_area",
                "gridlines",
            ],
            "page_setup",
        )?;
        let landscape = opt::<Value>(ruby, options, "landscape")?.map(|v| v.to_bool());
        let paper = match opt::<Value>(ruby, options, "paper")? {
            None => None,
            Some(v) if Integer::from_value(v).is_some() => Some(
                Integer::from_value(v)
                    .and_then(|n| n.to_u8().ok())
                    .ok_or_else(|| {
                        Error::new(
                            ruby.exception_arg_error(),
                            format!("invalid paper {}: use a symbol or Excel's paper number (0 for the printer's default)", v.inspect()),
                        )
                    })?,
            ),
            Some(v) => Some(choice(ruby, "paper", v, PAPER_SIZES)?),
        };
        let fit_width = opt::<u16>(ruby, options, "fit_width")?;
        let fit_height = opt::<u16>(ruby, options, "fit_height")?;
        let repeat_rows = opt::<(u32, u32)>(ruby, options, "repeat_rows")?;
        let repeat_columns = opt::<(u16, u16)>(ruby, options, "repeat_columns")?;
        let print_area = opt::<(u32, u16, u32, u16)>(ruby, options, "print_area")?;
        let gridlines = opt::<Value>(ruby, options, "gridlines")?.map(|v| v.to_bool());
        rb_self.with_ws(|ws| {
            match landscape {
                Some(true) => ws.set_landscape(),
                Some(false) => ws.set_portrait(),
                None => ws,
            };
            if let Some(paper) = paper {
                ws.set_paper_size(paper);
            }
            // 0 means as many pages as the content needs.
            if fit_width.unwrap_or(0) > 0 || fit_height.unwrap_or(0) > 0 {
                ws.set_print_fit_to_pages(fit_width.unwrap_or(0), fit_height.unwrap_or(0));
            }
            if let Some((first, last)) = repeat_rows {
                ws.set_repeat_rows(first, last).map_err(xerr)?;
            }
            if let Some((first, last)) = repeat_columns {
                ws.set_repeat_columns(first, last).map_err(xerr)?;
            }
            if let Some((fr, fc, lr, lc)) = print_area {
                ws.set_print_area(fr, fc, lr, lc).map_err(xerr)?;
            }
            if let Some(gridlines) = gridlines {
                ws.set_print_gridlines(gridlines);
            }
            Ok(())
        })
    }

    fn group_rows(
        ruby: &Ruby,
        rb_self: Obj<Self>,
        first: u32,
        last: u32,
        collapsed: bool,
    ) -> Result<Obj<Self>, Error> {
        // rust_xlsxwriter writes no outline levels when it writes rows to
        // disk as it goes.
        if rb_self.flushes_rows {
            return Err(Error::new(
                ruby.get_inner(&ERROR),
                "group_rows needs memory: :standard (rows written to disk as they go lose their outline level)",
            ));
        }
        rb_self.with_ws(|ws| {
            if collapsed {
                ws.group_rows_collapsed(first, last)
            } else {
                ws.group_rows(first, last)
            }
            .map(|_| ())
            .map_err(xerr)
        })?;
        Ok(rb_self)
    }

    fn group_columns(
        rb_self: Obj<Self>,
        first: u16,
        last: u16,
        collapsed: bool,
    ) -> Result<Obj<Self>, Error> {
        rb_self.with_ws(|ws| {
            if collapsed {
                ws.group_columns_collapsed(first, last)
            } else {
                ws.group_columns(first, last)
            }
            .map(|_| ())
            .map_err(xerr)
        })?;
        Ok(rb_self)
    }

    fn protect(
        ruby: &Ruby,
        rb_self: Obj<Self>,
        password: Option<String>,
        allow: RArray,
    ) -> Result<Obj<Self>, Error> {
        let mut options = ProtectionOptions::new();
        each_entry(allow, |_, action| {
            *choice(ruby, "protect action", action, PROTECTION_ALLOW)?(&mut options) = true;
            Ok(())
        })?;
        rb_self.with_ws(|ws| {
            // Always set the password, so protecting again replaces it ("" is none).
            ws.protect_with_password(password.as_deref().unwrap_or(""));
            ws.protect_with_options(&options); // keeps the password
            Ok(())
        })?;
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
    wb.define_singleton_method("_new", function!(Workbook::new, 2))?;
    wb.define_method("_add_worksheet", method!(Workbook::add_worksheet, 1))?;
    wb.define_method("to_xlsx", method!(Workbook::to_xlsx, 0))?;
    wb.define_method("_save", method!(Workbook::save, 1))?;
    wb.define_method("define_name", method!(Workbook::define_name, 2))?;
    wb.define_method("_properties", method!(Workbook::set_properties, 1))?;

    let ws = module.define_class("Worksheet", ruby.class_object())?;
    ws.define_method("_write", method!(Worksheet::write, 4))?;
    ws.define_method("write", method!(Worksheet::write_any, -1))?;
    ws.define_method("_append", method!(Worksheet::append, 2))?;
    ws.define_method("_column_width", method!(Worksheet::set_column_width, 3))?;
    ws.define_method("_column_format", method!(Worksheet::set_column_format, 3))?;
    ws.define_method("_autofit", method!(Worksheet::autofit, 0))?;
    ws.define_method("_autofilter", method!(Worksheet::autofilter, 4))?;
    ws.define_method("name", method!(Worksheet::name, 0))?;
    ws.define_method("_write_comment", method!(Worksheet::write_comment, 4))?;
    ws.define_method("_insert_image", method!(Worksheet::insert_image, 4))?;
    ws.define_method("_insert_chart", method!(Worksheet::insert_chart, 3))?;
    ws.define_method("_add_table", method!(Worksheet::add_table, 5))?;
    ws.define_method(
        "_conditional_format",
        method!(Worksheet::conditional_format, 5),
    )?;
    ws.define_method("_data_validation", method!(Worksheet::data_validation, 5))?;
    ws.define_method("_freeze_panes", method!(Worksheet::freeze_panes, 2))?;
    ws.define_method("row_height", method!(Worksheet::set_row_height, 2))?;
    ws.define_method("page_breaks", method!(Worksheet::set_page_breaks, 1))?;
    ws.define_method("_page_header", method!(Worksheet::set_header, 1))?;
    ws.define_method("_page_footer", method!(Worksheet::set_footer, 1))?;
    ws.define_method("_margins", method!(Worksheet::set_margins, 6))?;
    ws.define_method(
        "vertical_page_breaks",
        method!(Worksheet::set_vertical_page_breaks, 1),
    )?;
    ws.define_method("_hide_rows", method!(Worksheet::hide_rows, 2))?;
    ws.define_method("_hide_columns", method!(Worksheet::hide_columns, 2))?;
    ws.define_method("_row_format", method!(Worksheet::set_row_format, 3))?;
    ws.define_method(
        "default_row_height",
        method!(Worksheet::default_row_height, 1),
    )?;
    ws.define_method("_selection", method!(Worksheet::selection, 4))?;
    ws.define_method("_top_left_cell", method!(Worksheet::top_left_cell, 2))?;
    ws.define_method("_ignore_error", method!(Worksheet::ignore_error, 5))?;
    ws.define_method("_unprotect_range", method!(Worksheet::unprotect_range, 6))?;
    ws.define_method("activate", method!(Worksheet::activate, 0))?;
    ws.define_method("hide", method!(Worksheet::hide, 0))?;
    ws.define_method("zoom", method!(Worksheet::zoom, 1))?;
    ws.define_method("tab_color", method!(Worksheet::tab_color, 1))?;
    ws.define_method("hide_gridlines", method!(Worksheet::hide_gridlines, 0))?;
    ws.define_method("_page_setup", method!(Worksheet::page_setup, 1))?;
    ws.define_method("_group_rows", method!(Worksheet::group_rows, 3))?;
    ws.define_method("_group_columns", method!(Worksheet::group_columns, 3))?;
    ws.define_method("_protect", method!(Worksheet::protect, 2))?;
    ws.define_method("_merge_range", method!(Worksheet::merge_range, 6))?;

    let format = module.define_class("Format", ruby.class_object())?;
    format.define_singleton_method("_new", function!(Format::new, 1))?;
    ws.define_method("<<", method!(Worksheet::push, 1))?;
    ws.define_method("concat", method!(Worksheet::concat, 1))?;
    ws.define_method("next_row", method!(Worksheet::next_row, 0))?;
    Ok(())
}
