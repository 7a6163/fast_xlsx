use std::cell::Cell;
use std::sync::{Arc, Mutex};

use magnus::{
    function, method, prelude::*, r_hash::ForEach, typed_data::Obj, value::Lazy, Error,
    ExceptionClass, Integer, RArray, RClass, RHash, RModule, RString, Ruby, Symbol, TryConvert,
    Value,
};
use rust_xlsxwriter::{
    Color, FormatAlign, FormatBorder, FormatUnderline, IntoExcelData, XlsxError,
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
        emit(ws, row, col, rust_xlsxwriter::Url::new(url), format)
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
                "underline" if value.to_bool() => taken.set_underline(FormatUnderline::Single),
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
                "bold" | "italic" | "underline" | "text_wrap" => taken,
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

    let ws = module.define_class("Worksheet", ruby.class_object())?;
    ws.define_method("_write", method!(Worksheet::write, 4))?;
    ws.define_method("_append", method!(Worksheet::append, 2))?;
    ws.define_method("_set_column_width", method!(Worksheet::set_column_width, 3))?;
    ws.define_method("autofit", method!(Worksheet::autofit, 0))?;
    ws.define_method("autofilter", method!(Worksheet::autofilter, 4))?;

    let format = module.define_class("Format", ruby.class_object())?;
    format.define_singleton_method("_new", function!(Format::new, 1))?;
    ws.define_method("<<", method!(Worksheet::push, 1))?;
    ws.define_method("concat", method!(Worksheet::concat, 1))?;
    ws.define_method("next_row", method!(Worksheet::next_row, 0))?;
    Ok(())
}
