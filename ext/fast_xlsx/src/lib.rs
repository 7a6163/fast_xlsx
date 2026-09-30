use std::cell::Cell;
use std::sync::{Arc, Mutex};

use magnus::{
    function, method, prelude::*, r_hash::ForEach, typed_data::Obj, value::Lazy, Error,
    ExceptionClass, RArray, RClass, RHash, RModule, RString, Ruby, Symbol, TryConvert, Value,
};
use rust_xlsxwriter::{FormatUnderline, IntoExcelData, XlsxError};

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
                "bold" | "italic" | "underline" => taken,
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

impl Worksheet {
    fn advance(&self, row: u32) {
        self.next_row.set(self.next_row.get().max(row + 1));
    }

    fn write_row(
        &self,
        ruby: &Ruby,
        row: u32,
        cells: RArray,
        format: Option<&Format>,
    ) -> Result<(), Error> {
        let mut wb = self.wb.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index).map_err(xerr)?;
        for (col, v) in cells.into_iter().enumerate() {
            let col = u16::try_from(col)
                .map_err(|_| Error::new(ruby.exception_arg_error(), "too many columns"))?;
            put(ruby, ws, row, col, v, format)?;
        }
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
        let mut wb = rb_self.wb.lock().unwrap();
        let ws = wb.worksheet_from_index(rb_self.index).map_err(xerr)?;
        put(ruby, ws, row, col, v, format)?;
        rb_self.advance(row);
        Ok(())
    }

    fn append(
        ruby: &Ruby,
        rb_self: Obj<Self>,
        cells: RArray,
        format: Option<&Format>,
    ) -> Result<Obj<Self>, Error> {
        rb_self.write_row(ruby, rb_self.next_row.get(), cells, format)?;
        Ok(rb_self)
    }

    fn push(ruby: &Ruby, rb_self: Obj<Self>, cells: RArray) -> Result<Obj<Self>, Error> {
        Self::append(ruby, rb_self, cells, None)
    }

    fn concat(ruby: &Ruby, rb_self: Obj<Self>, rows: RArray) -> Result<Obj<Self>, Error> {
        for r in rows.into_iter() {
            rb_self.write_row(ruby, rb_self.next_row.get(), RArray::try_convert(r)?, None)?;
        }
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

    let format = module.define_class("Format", ruby.class_object())?;
    format.define_singleton_method("_new", function!(Format::new, 1))?;
    ws.define_method("<<", method!(Worksheet::push, 1))?;
    ws.define_method("concat", method!(Worksheet::concat, 1))?;
    ws.define_method("next_row", method!(Worksheet::next_row, 0))?;
    Ok(())
}
