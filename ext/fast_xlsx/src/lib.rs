use std::cell::Cell;
use std::sync::{Arc, Mutex};

use magnus::{
    function, method, prelude::*, typed_data::Obj, value::Lazy, Error, ExceptionClass, RArray,
    RModule, RString, Ruby, Value,
};
use rust_xlsxwriter::XlsxError;

// FastXlsx::Error is defined in lib/fast_xlsx.rb before this extension loads.
static ERROR: Lazy<ExceptionClass> = Lazy::new(|ruby| {
    ruby.class_object()
        .const_get::<_, RModule>("FastXlsx")
        .and_then(|m| m.const_get("Error"))
        .unwrap()
});

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
        let buf = rb_self.inner.lock().unwrap().save_to_buffer().map_err(xerr)?;
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

fn put(
    ruby: &Ruby,
    ws: &mut rust_xlsxwriter::Worksheet,
    row: u32,
    col: u16,
    v: Value,
) -> Result<(), Error> {
    let r = if v.is_nil() {
        return Ok(());
    } else if let Some(s) = RString::from_value(v) {
        // SAFETY: the borrowed str is copied by write_string before any Ruby code runs.
        ws.write_string(row, col, unsafe { s.as_str()? })
    } else if v.is_kind_of(ruby.class_numeric()) {
        ws.write_number(row, col, f64::try_convert(v)?)
    } else if v.is_kind_of(ruby.class_time()) {
        ws.write_number(row, col, excel_time(v)?)
    } else if v.is_kind_of(ruby.class_true_class()) || v.is_kind_of(ruby.class_false_class()) {
        ws.write_boolean(row, col, v.to_bool())
    } else {
        let s: String = v.funcall("to_s", ())?;
        ws.write_string(row, col, s)
    };
    r.map(|_| ()).map_err(xerr)
}

impl Worksheet {
    fn advance(&self, row: u32) {
        self.next_row.set(self.next_row.get().max(row + 1));
    }

    fn write_row(&self, ruby: &Ruby, row: u32, cells: RArray) -> Result<(), Error> {
        let mut wb = self.wb.lock().unwrap();
        let ws = wb.worksheet_from_index(self.index).map_err(xerr)?;
        for (col, v) in cells.into_iter().enumerate() {
            let col = u16::try_from(col)
                .map_err(|_| Error::new(ruby.exception_arg_error(), "too many columns"))?;
            put(ruby, ws, row, col, v)?;
        }
        self.advance(row);
        Ok(())
    }

    fn write(ruby: &Ruby, rb_self: &Self, row: u32, col: u16, v: Value) -> Result<(), Error> {
        let mut wb = rb_self.wb.lock().unwrap();
        let ws = wb.worksheet_from_index(rb_self.index).map_err(xerr)?;
        put(ruby, ws, row, col, v)?;
        rb_self.advance(row);
        Ok(())
    }

    fn push(ruby: &Ruby, rb_self: Obj<Self>, cells: RArray) -> Result<Obj<Self>, Error> {
        rb_self.write_row(ruby, rb_self.next_row.get(), cells)?;
        Ok(rb_self)
    }

    fn concat(ruby: &Ruby, rb_self: Obj<Self>, rows: RArray) -> Result<Obj<Self>, Error> {
        for r in rows.into_iter() {
            rb_self.write_row(ruby, rb_self.next_row.get(), RArray::try_convert(r)?)?;
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
    ws.define_method("write", method!(Worksheet::write, 3))?;
    ws.define_method("<<", method!(Worksheet::push, 1))?;
    ws.define_method("concat", method!(Worksheet::concat, 1))?;
    ws.define_method("next_row", method!(Worksheet::next_row, 0))?;
    Ok(())
}
