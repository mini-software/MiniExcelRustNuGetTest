use std::cell::RefCell;
use std::ffi::{CStr, c_char};
use std::panic::{AssertUnwindSafe, catch_unwind};
use std::ptr;
use std::str::FromStr;

use miniexcel::{
    CellReference, CellValue, CommentPerson, CommentTimestamp, CsvConfiguration, CsvEncoding,
    CsvReadOptions, CsvWriteOptions, DynamicRow, ExistingSheetPolicy, HeaderMode, InsertOptions,
    MergeSameCellsOptions, MiniExcel, ReadOptions, SheetType, SheetVisibility,
    TargetRelationshipPolicy, TemplateOptions, WriteOptions,
};

const ABI_VERSION: u32 = 1;
const RESULT_END: i32 = 0;
const RESULT_BATCH: i32 = 1;
const ERROR_INVALID_ARGUMENT: i32 = -1;
const ERROR_QUERY: i32 = -2;
const ERROR_PANIC: i32 = -3;
const ERROR_WRITE: i32 = -4;

thread_local! {
    static LAST_ERROR: RefCell<Vec<u8>> = const { RefCell::new(Vec::new()) };
}

pub struct QueryHandle {
    rows: Box<dyn Iterator<Item = miniexcel::Result<DynamicRow>> + Send>,
    frame: Vec<u8>,
}

pub struct BufferHandle {
    frame: Vec<u8>,
}

struct QueryOpenOptions {
    path: *const c_char,
    use_header_row: u8,
    sheet_name: *const c_char,
    start_cell: *const c_char,
    end_cell: *const c_char,
    ignore_empty_rows: u8,
    fill_merged_cells: u8,
    trim_headers: u8,
    enable_shared_string_cache: u8,
    shared_string_cache_size: u64,
    shared_string_cache_path: *const c_char,
}

struct CsvWriteArguments {
    path: *const c_char,
    data: *const u8,
    data_length: usize,
    delimiter: u8,
    encoding: u8,
    write_bom: u8,
    print_header: u8,
    overwrite_file: u8,
}

#[unsafe(no_mangle)]
pub extern "C" fn miniexcel_abi_version() -> u32 {
    ABI_VERSION
}

/// Opens a path-based XLSX query and returns an opaque native handle.
///
/// # Safety
///
/// String pointers must be null-terminated UTF-8. `path`, `start_cell`, and `out_handle` must be
/// non-null and valid for the duration of the call. `sheet_name` may be null.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn miniexcel_query_open(
    path: *const c_char,
    use_header_row: u8,
    sheet_name: *const c_char,
    start_cell: *const c_char,
    out_handle: *mut *mut QueryHandle,
) -> i32 {
    ffi_result(|| unsafe {
        open_query(
            QueryOpenOptions {
                path,
                use_header_row,
                sheet_name,
                start_cell,
                end_cell: ptr::null(),
                ignore_empty_rows: 0,
                fill_merged_cells: 0,
                trim_headers: 1,
                enable_shared_string_cache: 1,
                shared_string_cache_size: 5 * 1024 * 1024,
                shared_string_cache_path: ptr::null(),
            },
            out_handle,
        )
    })
}

/// Opens a bounded path-based XLSX query and returns an opaque native handle.
///
/// # Safety
///
/// String pointers must be null-terminated UTF-8. `path`, `start_cell`, and `out_handle` must be
/// non-null and valid for the duration of the call. `sheet_name` and `end_cell` may be null.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn miniexcel_query_range_open(
    path: *const c_char,
    use_header_row: u8,
    sheet_name: *const c_char,
    start_cell: *const c_char,
    end_cell: *const c_char,
    out_handle: *mut *mut QueryHandle,
) -> i32 {
    ffi_result(|| unsafe {
        open_query(
            QueryOpenOptions {
                path,
                use_header_row,
                sheet_name,
                start_cell,
                end_cell,
                ignore_empty_rows: 0,
                fill_merged_cells: 0,
                trim_headers: 1,
                enable_shared_string_cache: 1,
                shared_string_cache_size: 5 * 1024 * 1024,
                shared_string_cache_path: ptr::null(),
            },
            out_handle,
        )
    })
}

/// Opens a configured path-based XLSX query and returns an opaque native handle.
///
/// # Safety
///
/// Required string and output pointers must be non-null and valid for the duration of the call.
/// `sheet_name`, `end_cell`, and `shared_string_cache_path` may be null.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn miniexcel_query_options_open(
    path: *const c_char,
    use_header_row: u8,
    sheet_name: *const c_char,
    start_cell: *const c_char,
    end_cell: *const c_char,
    ignore_empty_rows: u8,
    fill_merged_cells: u8,
    trim_headers: u8,
    enable_shared_string_cache: u8,
    shared_string_cache_size: u64,
    shared_string_cache_path: *const c_char,
    out_handle: *mut *mut QueryHandle,
) -> i32 {
    ffi_result(|| unsafe {
        open_query(
            QueryOpenOptions {
                path,
                use_header_row,
                sheet_name,
                start_cell,
                end_cell,
                ignore_empty_rows,
                fill_merged_cells,
                trim_headers,
                enable_shared_string_cache,
                shared_string_cache_size,
                shared_string_cache_path,
            },
            out_handle,
        )
    })
}

/// Opens a path-based query over a named OpenXML table.
///
/// # Safety
///
/// `path`, `table_name`, and `out_handle` must be non-null and valid for the duration of the call.
/// `sheet_name` may be null.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn miniexcel_query_table_open(
    path: *const c_char,
    sheet_name: *const c_char,
    table_name: *const c_char,
    out_handle: *mut *mut QueryHandle,
) -> i32 {
    ffi_result(|| {
        if path.is_null() || table_name.is_null() || out_handle.is_null() {
            set_last_error("path, table_name, and out_handle are required");
            return Err(ERROR_INVALID_ARGUMENT);
        }

        unsafe { ptr::write(out_handle, ptr::null_mut()) };
        let path = unsafe { read_utf8(path) }?;
        let table_name = unsafe { read_utf8(table_name) }?;
        if table_name.is_empty() {
            set_last_error("table_name cannot be empty");
            return Err(ERROR_INVALID_ARGUMENT);
        }
        let sheet_name = if sheet_name.is_null() {
            None
        } else {
            let value = unsafe { read_utf8(sheet_name) }?;
            (!value.is_empty()).then_some(value)
        };

        let rows = MiniExcel::query_table(path, table_name, sheet_name).map_err(|error| {
            set_last_error(error.to_string());
            ERROR_QUERY
        })?;
        let handle = Box::new(QueryHandle {
            rows,
            frame: Vec::new(),
        });
        unsafe { ptr::write(out_handle, Box::into_raw(handle)) };
        Ok(RESULT_BATCH)
    })
}

/// Opens a path-based CSV query using explicit read options.
///
/// # Safety
///
/// `path` and `out_handle` must be non-null and valid for the duration of the call.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn miniexcel_query_csv_open(
    path: *const c_char,
    use_header_row: u8,
    delimiter: u8,
    encoding: u8,
    read_empty_as_null: u8,
    trim_headers: u8,
    out_handle: *mut *mut QueryHandle,
) -> i32 {
    ffi_result(|| {
        if path.is_null() || out_handle.is_null() {
            set_last_error("path and out_handle are required");
            return Err(ERROR_INVALID_ARGUMENT);
        }
        if delimiter == 0 {
            set_last_error("delimiter must be a single-byte character");
            return Err(ERROR_INVALID_ARGUMENT);
        }

        unsafe { ptr::write(out_handle, ptr::null_mut()) };
        let path = unsafe { read_utf8(path) }?;
        let options = csv_read_options(
            use_header_row,
            delimiter,
            encoding,
            read_empty_as_null,
            trim_headers,
        )?;
        let rows = MiniExcel::query_csv_with_options(path, &options).map_err(|error| {
            set_last_error(error.to_string());
            ERROR_QUERY
        })?;
        let handle = Box::new(QueryHandle {
            rows,
            frame: Vec::new(),
        });
        unsafe { ptr::write(out_handle, Box::into_raw(handle)) };
        Ok(RESULT_BATCH)
    })
}

/// Returns selected CSV column names through an owned metadata buffer.
///
/// # Safety
///
/// `path` and all output pointers must be non-null and valid for the duration of the call.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn miniexcel_get_csv_columns(
    path: *const c_char,
    use_header_row: u8,
    delimiter: u8,
    encoding: u8,
    read_empty_as_null: u8,
    trim_headers: u8,
    out_handle: *mut *mut BufferHandle,
    out_data: *mut *const u8,
    out_length: *mut usize,
) -> i32 {
    ffi_result(|| {
        if path.is_null() || out_handle.is_null() || out_data.is_null() || out_length.is_null() {
            set_last_error("path, out_handle, out_data, and out_length are required");
            return Err(ERROR_INVALID_ARGUMENT);
        }

        unsafe {
            ptr::write(out_handle, ptr::null_mut());
            ptr::write(out_data, ptr::null());
            ptr::write(out_length, 0);
        }

        let path = unsafe { read_utf8(path) }?;
        let options = csv_read_options(
            use_header_row,
            delimiter,
            encoding,
            read_empty_as_null,
            trim_headers,
        )?;
        let columns = MiniExcel::get_csv_columns(path, &options).map_err(|error| {
            set_last_error(error.to_string());
            ERROR_QUERY
        })?;
        let handle = Box::new(BufferHandle {
            frame: write_strings(columns)?,
        });
        unsafe {
            ptr::write(out_data, handle.frame.as_ptr());
            ptr::write(out_length, handle.frame.len());
            ptr::write(out_handle, Box::into_raw(handle));
        }
        Ok(RESULT_BATCH)
    })
}

/// Writes the next bounded batch into memory owned by the query handle.
///
/// # Safety
///
/// `handle` must have been returned by `miniexcel_query_open` and not yet closed. Output pointers
/// must be non-null and writable. Returned data is valid until the next call using the handle.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn miniexcel_query_next_batch(
    handle: *mut QueryHandle,
    max_rows: u32,
    out_data: *mut *const u8,
    out_length: *mut usize,
) -> i32 {
    ffi_result(|| {
        if handle.is_null() || max_rows == 0 || out_data.is_null() || out_length.is_null() {
            set_last_error("handle, max_rows, out_data, and out_length are required");
            return Err(ERROR_INVALID_ARGUMENT);
        }

        let handle = unsafe { &mut *handle };
        handle.frame.clear();
        write_u32(&mut handle.frame, 0);

        let mut row_count = 0_u32;
        while row_count < max_rows {
            let Some(row) = handle.rows.next() else {
                break;
            };
            let row = row.map_err(|error| {
                set_last_error(error.to_string());
                ERROR_QUERY
            })?;
            write_row(&mut handle.frame, &row)?;
            row_count += 1;
        }

        if row_count == 0 {
            unsafe {
                ptr::write(out_data, ptr::null());
                ptr::write(out_length, 0);
            }
            return Ok(RESULT_END);
        }

        handle.frame[0..4].copy_from_slice(&row_count.to_le_bytes());
        unsafe {
            ptr::write(out_data, handle.frame.as_ptr());
            ptr::write(out_length, handle.frame.len());
        }
        Ok(RESULT_BATCH)
    })
}

/// Closes a query handle and releases its worker and temporary resources.
///
/// # Safety
///
/// `handle` must be null or a handle returned by `miniexcel_query_open` that has not been closed.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn miniexcel_query_close(handle: *mut QueryHandle) {
    if !handle.is_null() {
        let _ = catch_unwind(AssertUnwindSafe(|| drop(unsafe { Box::from_raw(handle) })));
    }
}

/// Returns worksheet names in workbook order using memory owned by an opaque buffer handle.
///
/// # Safety
///
/// `path`, `out_handle`, `out_data`, and `out_length` must be non-null and valid for the duration
/// of the call. Returned data remains valid until `miniexcel_buffer_close` closes the handle.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn miniexcel_get_sheet_names(
    path: *const c_char,
    out_handle: *mut *mut BufferHandle,
    out_data: *mut *const u8,
    out_length: *mut usize,
) -> i32 {
    ffi_result(|| {
        if path.is_null() || out_handle.is_null() || out_data.is_null() || out_length.is_null() {
            set_last_error("path, out_handle, out_data, and out_length are required");
            return Err(ERROR_INVALID_ARGUMENT);
        }

        unsafe {
            ptr::write(out_handle, ptr::null_mut());
            ptr::write(out_data, ptr::null());
            ptr::write(out_length, 0);
        }

        let path = unsafe { read_utf8(path) }?;
        let names = MiniExcel::get_sheet_names(path).map_err(|error| {
            set_last_error(error.to_string());
            ERROR_QUERY
        })?;
        let mut frame = Vec::new();
        write_length(&mut frame, names.len())?;
        for name in names {
            write_string(&mut frame, name)?;
        }

        let handle = Box::new(BufferHandle { frame });
        unsafe {
            ptr::write(out_data, handle.frame.as_ptr());
            ptr::write(out_length, handle.frame.len());
            ptr::write(out_handle, Box::into_raw(handle));
        }
        Ok(RESULT_BATCH)
    })
}

/// Returns selected column names using memory owned by an opaque buffer handle.
///
/// # Safety
///
/// `path`, `start_cell`, and all output pointers must be non-null and valid for the duration of
/// the call. `sheet_name` may be null. Returned data remains valid until the handle is closed.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn miniexcel_get_columns(
    path: *const c_char,
    use_header_row: u8,
    sheet_name: *const c_char,
    start_cell: *const c_char,
    out_handle: *mut *mut BufferHandle,
    out_data: *mut *const u8,
    out_length: *mut usize,
) -> i32 {
    ffi_result(|| {
        if path.is_null()
            || start_cell.is_null()
            || out_handle.is_null()
            || out_data.is_null()
            || out_length.is_null()
        {
            set_last_error("path, start_cell, out_handle, out_data, and out_length are required");
            return Err(ERROR_INVALID_ARGUMENT);
        }

        unsafe {
            ptr::write(out_handle, ptr::null_mut());
            ptr::write(out_data, ptr::null());
            ptr::write(out_length, 0);
        }

        let path = unsafe { read_utf8(path) }?;
        let start_cell = unsafe { read_utf8(start_cell) }?;
        let start_cell = CellReference::from_str(start_cell).map_err(|error| {
            set_last_error(error.to_string());
            ERROR_INVALID_ARGUMENT
        })?;
        let mut options = ReadOptions::new()
            .with_header_mode(if use_header_row == 0 {
                HeaderMode::None
            } else {
                HeaderMode::FirstRow
            })
            .with_start_cell(start_cell);

        if !sheet_name.is_null() {
            let sheet_name = unsafe { read_utf8(sheet_name) }?;
            if !sheet_name.is_empty() {
                options = options.with_sheet_name(sheet_name);
            }
        }

        let columns = MiniExcel::get_columns(path, &options).map_err(|error| {
            set_last_error(error.to_string());
            ERROR_QUERY
        })?;
        let handle = Box::new(BufferHandle {
            frame: write_strings(columns)?,
        });
        unsafe {
            ptr::write(out_data, handle.frame.as_ptr());
            ptr::write(out_length, handle.frame.len());
            ptr::write(out_handle, Box::into_raw(handle));
        }
        Ok(RESULT_BATCH)
    })
}

/// Returns worksheet dimensions as optional A1 start/end address pairs.
///
/// # Safety
///
/// `path` and all output pointers must be non-null and valid for the duration of the call.
/// Returned data remains valid until the buffer handle is closed.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn miniexcel_get_sheet_dimensions(
    path: *const c_char,
    out_handle: *mut *mut BufferHandle,
    out_data: *mut *const u8,
    out_length: *mut usize,
) -> i32 {
    ffi_result(|| {
        if path.is_null() || out_handle.is_null() || out_data.is_null() || out_length.is_null() {
            set_last_error("path, out_handle, out_data, and out_length are required");
            return Err(ERROR_INVALID_ARGUMENT);
        }

        unsafe {
            ptr::write(out_handle, ptr::null_mut());
            ptr::write(out_data, ptr::null());
            ptr::write(out_length, 0);
        }

        let path = unsafe { read_utf8(path) }?;
        let dimensions = MiniExcel::get_sheet_dimensions(path).map_err(|error| {
            set_last_error(error.to_string());
            ERROR_QUERY
        })?;
        let mut frame = Vec::new();
        write_length(&mut frame, dimensions.len())?;
        for dimension in dimensions {
            write_string(
                &mut frame,
                dimension
                    .start_cell()
                    .map(|cell| cell.to_string())
                    .unwrap_or_default(),
            )?;
            write_string(
                &mut frame,
                dimension
                    .end_cell()
                    .map(|cell| cell.to_string())
                    .unwrap_or_default(),
            )?;
        }

        let handle = Box::new(BufferHandle { frame });
        unsafe {
            ptr::write(out_data, handle.frame.as_ptr());
            ptr::write(out_length, handle.frame.len());
            ptr::write(out_handle, Box::into_raw(handle));
        }
        Ok(RESULT_BATCH)
    })
}

/// Returns worksheet metadata in workbook order.
///
/// # Safety
///
/// `path` and all output pointers must be non-null and valid for the duration of the call.
/// Returned data remains valid until the buffer handle is closed.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn miniexcel_get_sheet_info(
    path: *const c_char,
    out_handle: *mut *mut BufferHandle,
    out_data: *mut *const u8,
    out_length: *mut usize,
) -> i32 {
    ffi_result(|| {
        if path.is_null() || out_handle.is_null() || out_data.is_null() || out_length.is_null() {
            set_last_error("path, out_handle, out_data, and out_length are required");
            return Err(ERROR_INVALID_ARGUMENT);
        }

        unsafe {
            ptr::write(out_handle, ptr::null_mut());
            ptr::write(out_data, ptr::null());
            ptr::write(out_length, 0);
        }

        let path = unsafe { read_utf8(path) }?;
        let sheets = MiniExcel::get_sheet_info(path).map_err(|error| {
            set_last_error(error.to_string());
            ERROR_QUERY
        })?;
        let mut frame = Vec::new();
        write_length(&mut frame, sheets.len())?;
        for sheet in sheets {
            write_u32(&mut frame, sheet.id());
            write_length(&mut frame, sheet.index())?;
            write_string(&mut frame, sheet.name())?;
            frame.push(match sheet.sheet_type() {
                SheetType::Worksheet => 0,
                SheetType::DialogSheet => 1,
                SheetType::MacroSheet => 2,
                SheetType::ChartSheet => 3,
                SheetType::Vba => 4,
            });
            frame.push(match sheet.visibility() {
                SheetVisibility::Visible => 0,
                SheetVisibility::Hidden => 1,
                SheetVisibility::VeryHidden => 2,
            });
            frame.push(u8::from(sheet.is_active()));
        }

        let handle = Box::new(BufferHandle { frame });
        unsafe {
            ptr::write(out_data, handle.frame.as_ptr());
            ptr::write(out_length, handle.frame.len());
            ptr::write(out_handle, Box::into_raw(handle));
        }
        Ok(RESULT_BATCH)
    })
}

/// Returns threaded comments, replies, and legacy notes for a worksheet.
///
/// # Safety
///
/// `path` and all output pointers must be non-null and valid for the duration of the call.
/// `sheet_name` may be null. Returned data remains valid until the buffer handle is closed.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn miniexcel_get_comments(
    path: *const c_char,
    sheet_name: *const c_char,
    out_handle: *mut *mut BufferHandle,
    out_data: *mut *const u8,
    out_length: *mut usize,
) -> i32 {
    ffi_result(|| {
        if path.is_null() || out_handle.is_null() || out_data.is_null() || out_length.is_null() {
            set_last_error("path, out_handle, out_data, and out_length are required");
            return Err(ERROR_INVALID_ARGUMENT);
        }

        unsafe {
            ptr::write(out_handle, ptr::null_mut());
            ptr::write(out_data, ptr::null());
            ptr::write(out_length, 0);
        }

        let path = unsafe { read_utf8(path) }?;
        let sheet_name = if sheet_name.is_null() {
            None
        } else {
            let value = unsafe { read_utf8(sheet_name) }?;
            (!value.is_empty()).then_some(value)
        };
        let comments = MiniExcel::get_comments(path, sheet_name).map_err(|error| {
            set_last_error(error.to_string());
            ERROR_QUERY
        })?;
        let mut frame = Vec::new();
        write_string(&mut frame, comments.sheet_name())?;
        write_length(&mut frame, comments.threaded_comments().len())?;
        for comment in comments.threaded_comments() {
            write_string(&mut frame, comment.id().to_string())?;
            write_string(&mut frame, comment.cell().to_string())?;
            write_person(&mut frame, comment.person())?;
            write_timestamp(&mut frame, comment.created_at())?;
            frame.push(u8::from(comment.resolved()));
            write_string(&mut frame, comment.text())?;
            write_length(&mut frame, comment.replies().len())?;
            for reply in comment.replies() {
                write_string(&mut frame, reply.id().to_string())?;
                write_string(&mut frame, reply.parent_id().to_string())?;
                write_person(&mut frame, reply.person())?;
                write_timestamp(&mut frame, reply.created_at())?;
                write_string(&mut frame, reply.text())?;
            }
        }
        write_length(&mut frame, comments.notes().len())?;
        for note in comments.notes() {
            write_optional_string(&mut frame, note.id().map(|id| id.to_string()).as_deref())?;
            write_string(&mut frame, note.cell().to_string())?;
            write_optional_string(&mut frame, note.author())?;
            write_string(&mut frame, note.text())?;
        }

        let handle = Box::new(BufferHandle { frame });
        unsafe {
            ptr::write(out_data, handle.frame.as_ptr());
            ptr::write(out_length, handle.frame.len());
            ptr::write(out_handle, Box::into_raw(handle));
        }
        Ok(RESULT_BATCH)
    })
}

/// Creates a single-sheet XLSX workbook from encoded dynamic rows.
///
/// # Safety
///
/// `path`, `data`, and `out_row_count` must be non-null and valid for the supplied lengths.
/// `sheet_name` may be null.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn miniexcel_save_as(
    path: *const c_char,
    data: *const u8,
    data_length: usize,
    print_header: u8,
    sheet_name: *const c_char,
    overwrite_file: u8,
    out_row_count: *mut u32,
) -> i32 {
    ffi_result(|| {
        if path.is_null() || data.is_null() || out_row_count.is_null() {
            set_last_error("path, data, and out_row_count are required");
            return Err(ERROR_INVALID_ARGUMENT);
        }

        unsafe { ptr::write(out_row_count, 0) };
        let path = unsafe { read_utf8(path) }?;
        let bytes = unsafe { std::slice::from_raw_parts(data, data_length) };
        let rows = decode_rows(bytes)?;
        let mut options = WriteOptions::new()
            .with_print_header(print_header != 0)
            .with_overwrite_file(overwrite_file != 0);
        if !sheet_name.is_null() {
            let sheet_name = unsafe { read_utf8(sheet_name) }?;
            if !sheet_name.is_empty() {
                options = options.with_sheet_name(sheet_name);
            }
        }
        MiniExcel::save_as_with_options(path, &rows, &options).map_err(|error| {
            set_last_error(error.to_string());
            ERROR_WRITE
        })?;
        let row_count = u32::try_from(rows.len()).map_err(|_| {
            set_last_error("row count exceeds the ABI limit");
            ERROR_WRITE
        })?;
        unsafe { ptr::write(out_row_count, row_count) };
        Ok(RESULT_BATCH)
    })
}

/// Creates a multi-sheet XLSX workbook from an ordered encoded sheet collection.
///
/// # Safety
///
/// `path`, `data`, and all output pointers must be valid for the supplied lengths.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn miniexcel_save_as_sheets(
    path: *const c_char,
    data: *const u8,
    data_length: usize,
    print_header: u8,
    overwrite_file: u8,
    out_handle: *mut *mut BufferHandle,
    out_data: *mut *const u8,
    out_length: *mut usize,
) -> i32 {
    ffi_result(|| {
        if path.is_null()
            || data.is_null()
            || out_handle.is_null()
            || out_data.is_null()
            || out_length.is_null()
        {
            set_last_error("path, data, out_handle, out_data, and out_length are required");
            return Err(ERROR_INVALID_ARGUMENT);
        }
        unsafe {
            ptr::write(out_handle, ptr::null_mut());
            ptr::write(out_data, ptr::null());
            ptr::write(out_length, 0);
        }
        let path = unsafe { read_utf8(path) }?;
        let sheets = decode_sheets(unsafe { std::slice::from_raw_parts(data, data_length) })?;
        let options = WriteOptions::new()
            .with_print_header(print_header != 0)
            .with_overwrite_file(overwrite_file != 0);
        let counts = MiniExcel::save_as_sheets(
            path,
            sheets.iter().map(|(name, rows)| (name, rows.as_slice())),
            &options,
        )
        .map_err(|error| {
            set_last_error(error.to_string());
            ERROR_WRITE
        })?;
        let mut frame = Vec::new();
        write_length(&mut frame, counts.len())?;
        for count in counts {
            write_length(&mut frame, count)?;
        }
        let handle = Box::new(BufferHandle { frame });
        unsafe {
            ptr::write(out_data, handle.frame.as_ptr());
            ptr::write(out_length, handle.frame.len());
            ptr::write(out_handle, Box::into_raw(handle));
        }
        Ok(RESULT_BATCH)
    })
}

/// Creates a CSV file from encoded dynamic rows.
///
/// # Safety
///
/// `path`, `data`, and `out_row_count` must be non-null and valid for the supplied lengths.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn miniexcel_save_csv(
    path: *const c_char,
    data: *const u8,
    data_length: usize,
    delimiter: u8,
    encoding: u8,
    write_bom: u8,
    print_header: u8,
    overwrite_file: u8,
    out_row_count: *mut u32,
) -> i32 {
    ffi_result(|| unsafe {
        write_csv(
            CsvWriteArguments {
                path,
                data,
                data_length,
                delimiter,
                encoding,
                write_bom,
                print_header,
                overwrite_file,
            },
            false,
            out_row_count,
        )
    })
}

/// Appends encoded dynamic rows to a CSV file.
///
/// # Safety
///
/// `path`, `data`, and `out_row_count` must be non-null and valid for the supplied lengths.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn miniexcel_append_csv(
    path: *const c_char,
    data: *const u8,
    data_length: usize,
    delimiter: u8,
    encoding: u8,
    write_bom: u8,
    print_header: u8,
    out_row_count: *mut u32,
) -> i32 {
    ffi_result(|| unsafe {
        write_csv(
            CsvWriteArguments {
                path,
                data,
                data_length,
                delimiter,
                encoding,
                write_bom,
                print_header,
                overwrite_file: 0,
            },
            true,
            out_row_count,
        )
    })
}

/// Inserts or replaces a worksheet in an XLSX workbook.
///
/// # Safety
///
/// `path`, `data`, `sheet_name`, and `out_row_count` must be valid for the supplied lengths.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn miniexcel_insert_sheet(
    path: *const c_char,
    data: *const u8,
    data_length: usize,
    sheet_name: *const c_char,
    print_header: u8,
    replace_existing: u8,
    remove_supported_relationships: u8,
    out_row_count: *mut u32,
) -> i32 {
    ffi_result(|| {
        if path.is_null() || data.is_null() || sheet_name.is_null() || out_row_count.is_null() {
            set_last_error("path, data, sheet_name, and out_row_count are required");
            return Err(ERROR_INVALID_ARGUMENT);
        }
        unsafe { ptr::write(out_row_count, 0) };
        let path = unsafe { read_utf8(path) }?;
        let sheet_name = unsafe { read_utf8(sheet_name) }?;
        let rows = decode_rows(unsafe { std::slice::from_raw_parts(data, data_length) })?;
        let options = insert_options(
            sheet_name,
            print_header,
            replace_existing,
            remove_supported_relationships,
            false,
        );
        let count = MiniExcel::insert(path, &rows, &options).map_err(|error| {
            set_last_error(error.to_string());
            ERROR_WRITE
        })?;
        write_row_count(count, out_row_count)
    })
}

/// Copies an XLSX workbook and adds or replaces one worksheet in the destination.
///
/// # Safety
///
/// Both paths, `data`, `sheet_name`, and `out_row_count` must be valid for supplied lengths.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn miniexcel_copy_and_add_sheet(
    source_path: *const c_char,
    destination_path: *const c_char,
    data: *const u8,
    data_length: usize,
    sheet_name: *const c_char,
    print_header: u8,
    replace_existing: u8,
    remove_supported_relationships: u8,
    overwrite_destination: u8,
    out_row_count: *mut u32,
) -> i32 {
    ffi_result(|| {
        if source_path.is_null()
            || destination_path.is_null()
            || data.is_null()
            || sheet_name.is_null()
            || out_row_count.is_null()
        {
            set_last_error(
                "source_path, destination_path, data, sheet_name, and out_row_count are required",
            );
            return Err(ERROR_INVALID_ARGUMENT);
        }
        unsafe { ptr::write(out_row_count, 0) };
        let source_path = unsafe { read_utf8(source_path) }?;
        let destination_path = unsafe { read_utf8(destination_path) }?;
        let sheet_name = unsafe { read_utf8(sheet_name) }?;
        let rows = decode_rows(unsafe { std::slice::from_raw_parts(data, data_length) })?;
        let options = insert_options(
            sheet_name,
            print_header,
            replace_existing,
            remove_supported_relationships,
            overwrite_destination != 0,
        );
        let count = MiniExcel::copy_and_add_sheet(source_path, destination_path, &rows, &options)
            .map_err(|error| {
            set_last_error(error.to_string());
            ERROR_WRITE
        })?;
        write_row_count(count, out_row_count)
    })
}

/// Fills an XLSX template from a UTF-8 JSON value and atomically writes the destination.
///
/// # Safety
///
/// All string pointers must be non-null, valid, null-terminated UTF-8 for the duration of the call.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn miniexcel_fill_template(
    destination_path: *const c_char,
    template_path: *const c_char,
    json_data: *const u8,
    json_length: usize,
    overwrite_file: u8,
    ignore_missing_variables: u8,
) -> i32 {
    ffi_result(|| {
        if destination_path.is_null() || template_path.is_null() || json_data.is_null() {
            set_last_error("destination_path, template_path, and json_data are required");
            return Err(ERROR_INVALID_ARGUMENT);
        }

        let destination_path = unsafe { read_utf8(destination_path) }?;
        let template_path = unsafe { read_utf8(template_path) }?;
        let json = unsafe { std::slice::from_raw_parts(json_data, json_length) };
        let value: serde_json::Value = serde_json::from_slice(json).map_err(|error| {
            set_last_error(format!("invalid template JSON: {error}"));
            ERROR_INVALID_ARGUMENT
        })?;
        let options = TemplateOptions::new()
            .with_overwrite_file(overwrite_file != 0)
            .with_ignore_missing_variables(ignore_missing_variables != 0);
        MiniExcel::save_as_template(destination_path, template_path, &value, &options).map_err(
            |error| {
                set_last_error(error.to_string());
                ERROR_WRITE
            },
        )?;
        Ok(RESULT_BATCH)
    })
}

/// Merges tagged same-value cells into a separate XLSX destination.
///
/// # Safety
///
/// Both path pointers must be non-null, valid, null-terminated UTF-8 for the duration of the call.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn miniexcel_merge_same_cells(
    destination_path: *const c_char,
    source_path: *const c_char,
    overwrite_file: u8,
) -> i32 {
    ffi_result(|| {
        if destination_path.is_null() || source_path.is_null() {
            set_last_error("destination_path and source_path are required");
            return Err(ERROR_INVALID_ARGUMENT);
        }
        let destination_path = unsafe { read_utf8(destination_path) }?;
        let source_path = unsafe { read_utf8(source_path) }?;
        let options = MergeSameCellsOptions::new().with_overwrite_file(overwrite_file != 0);
        MiniExcel::merge_same_cells(source_path, destination_path, &options).map_err(|error| {
            set_last_error(error.to_string());
            ERROR_WRITE
        })?;
        Ok(RESULT_BATCH)
    })
}

/// Atomically renames a worksheet in an existing XLSX workbook.
///
/// # Safety
///
/// All string pointers must be non-null, valid, null-terminated UTF-8 for the duration of the call.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn miniexcel_rename_sheet(
    path: *const c_char,
    sheet_name: *const c_char,
    new_sheet_name: *const c_char,
) -> i32 {
    ffi_result(|| {
        if path.is_null() || sheet_name.is_null() || new_sheet_name.is_null() {
            set_last_error("path, sheet_name, and new_sheet_name are required");
            return Err(ERROR_INVALID_ARGUMENT);
        }
        let path = unsafe { read_utf8(path) }?;
        let sheet_name = unsafe { read_utf8(sheet_name) }?;
        let new_sheet_name = unsafe { read_utf8(new_sheet_name) }?;
        MiniExcel::rename_sheet(path, sheet_name, new_sheet_name).map_err(|error| {
            set_last_error(error.to_string());
            ERROR_WRITE
        })?;
        Ok(RESULT_BATCH)
    })
}

/// Atomically moves a worksheet to a zero-based index.
///
/// # Safety
///
/// Both string pointers must be non-null, valid, null-terminated UTF-8 for the duration of the call.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn miniexcel_reorder_sheet(
    path: *const c_char,
    sheet_name: *const c_char,
    new_sheet_index: i32,
) -> i32 {
    ffi_result(|| {
        if path.is_null() || sheet_name.is_null() {
            set_last_error("path and sheet_name are required");
            return Err(ERROR_INVALID_ARGUMENT);
        }
        let path = unsafe { read_utf8(path) }?;
        let sheet_name = unsafe { read_utf8(sheet_name) }?;
        MiniExcel::reorder_sheet(path, sheet_name, new_sheet_index).map_err(|error| {
            set_last_error(error.to_string());
            ERROR_WRITE
        })?;
        Ok(RESULT_BATCH)
    })
}

/// Atomically changes a worksheet visibility state.
///
/// # Safety
///
/// Both string pointers must be non-null, valid, null-terminated UTF-8 for the duration of the call.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn miniexcel_set_sheet_visibility(
    path: *const c_char,
    sheet_name: *const c_char,
    visibility: u8,
) -> i32 {
    ffi_result(|| {
        if path.is_null() || sheet_name.is_null() {
            set_last_error("path and sheet_name are required");
            return Err(ERROR_INVALID_ARGUMENT);
        }
        let visibility = match visibility {
            0 => SheetVisibility::Visible,
            1 => SheetVisibility::Hidden,
            2 => SheetVisibility::VeryHidden,
            _ => {
                set_last_error("visibility is not supported");
                return Err(ERROR_INVALID_ARGUMENT);
            }
        };
        let path = unsafe { read_utf8(path) }?;
        let sheet_name = unsafe { read_utf8(sheet_name) }?;
        MiniExcel::set_sheet_visibility(path, sheet_name, visibility).map_err(|error| {
            set_last_error(error.to_string());
            ERROR_WRITE
        })?;
        Ok(RESULT_BATCH)
    })
}

/// Releases a buffer returned by a metadata operation.
///
/// # Safety
///
/// `handle` must be null or a handle returned by this library that has not already been closed.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn miniexcel_buffer_close(handle: *mut BufferHandle) {
    if !handle.is_null() {
        let _ = catch_unwind(AssertUnwindSafe(|| drop(unsafe { Box::from_raw(handle) })));
    }
}

/// Returns the last error recorded on the current native thread.
///
/// # Safety
///
/// `out_length` may be null; otherwise it must be writable. The returned data remains valid until
/// the next MiniExcel FFI error on this thread.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn miniexcel_last_error(out_length: *mut usize) -> *const u8 {
    LAST_ERROR.with(|error| {
        let error = error.borrow();
        if !out_length.is_null() {
            unsafe { ptr::write(out_length, error.len()) };
        }
        error.as_ptr()
    })
}

fn ffi_result(operation: impl FnOnce() -> Result<i32, i32>) -> i32 {
    match catch_unwind(AssertUnwindSafe(operation)) {
        Ok(Ok(result)) => result,
        Ok(Err(code)) => code,
        Err(_) => {
            set_last_error("Rust panic crossed the MiniExcel FFI boundary");
            ERROR_PANIC
        }
    }
}

unsafe fn read_utf8<'a>(value: *const c_char) -> Result<&'a str, i32> {
    unsafe { CStr::from_ptr(value) }.to_str().map_err(|error| {
        set_last_error(error.to_string());
        ERROR_INVALID_ARGUMENT
    })
}

unsafe fn open_query(
    arguments: QueryOpenOptions,
    out_handle: *mut *mut QueryHandle,
) -> Result<i32, i32> {
    let QueryOpenOptions {
        path,
        use_header_row,
        sheet_name,
        start_cell,
        end_cell,
        ignore_empty_rows,
        fill_merged_cells,
        trim_headers,
        enable_shared_string_cache,
        shared_string_cache_size,
        shared_string_cache_path,
    } = arguments;
    if path.is_null() || start_cell.is_null() || out_handle.is_null() {
        set_last_error("path, start_cell, and out_handle are required");
        return Err(ERROR_INVALID_ARGUMENT);
    }

    unsafe { ptr::write(out_handle, ptr::null_mut()) };
    let path = unsafe { read_utf8(path) }?;
    let start_cell = unsafe { read_utf8(start_cell) }?;
    let start_cell = CellReference::from_str(start_cell).map_err(|error| {
        set_last_error(error.to_string());
        ERROR_INVALID_ARGUMENT
    })?;

    let mut options = ReadOptions::new()
        .with_header_mode(if use_header_row == 0 {
            HeaderMode::None
        } else {
            HeaderMode::FirstRow
        })
        .with_start_cell(start_cell)
        .with_ignore_empty_rows(ignore_empty_rows != 0)
        .with_fill_merged_cells(fill_merged_cells != 0)
        .with_trim_headers(trim_headers != 0)
        .with_shared_string_disk_cache(enable_shared_string_cache != 0)
        .with_shared_string_cache_size(shared_string_cache_size);

    if !end_cell.is_null() {
        let end_cell = unsafe { read_utf8(end_cell) }?;
        let end_cell = CellReference::from_str(end_cell).map_err(|error| {
            set_last_error(error.to_string());
            ERROR_INVALID_ARGUMENT
        })?;
        options = options.with_end_cell(end_cell);
    }

    if !sheet_name.is_null() {
        let sheet_name = unsafe { read_utf8(sheet_name) }?;
        if !sheet_name.is_empty() {
            options = options.with_sheet_name(sheet_name);
        }
    }

    if !shared_string_cache_path.is_null() {
        let cache_path = unsafe { read_utf8(shared_string_cache_path) }?;
        if !cache_path.is_empty() {
            options = options.with_shared_string_cache_path(cache_path);
        }
    }

    let rows = MiniExcel::query_with_options(path, &options).map_err(|error| {
        set_last_error(error.to_string());
        ERROR_QUERY
    })?;
    let handle = Box::new(QueryHandle {
        rows,
        frame: Vec::new(),
    });
    unsafe { ptr::write(out_handle, Box::into_raw(handle)) };
    Ok(RESULT_BATCH)
}

fn set_last_error(message: impl AsRef<str>) {
    LAST_ERROR.with(|error| {
        let mut error = error.borrow_mut();
        error.clear();
        error.extend_from_slice(message.as_ref().as_bytes());
    });
}

fn write_row(frame: &mut Vec<u8>, row: &DynamicRow) -> Result<(), i32> {
    write_length(frame, row.len())?;
    for (name, value) in row {
        write_string(frame, name)?;
        match value {
            CellValue::Empty => frame.push(0),
            CellValue::Bool(value) => {
                frame.push(1);
                frame.push(u8::from(*value));
            }
            CellValue::Int(value) => {
                frame.push(2);
                frame.extend_from_slice(&value.to_le_bytes());
            }
            CellValue::Float(value) => {
                frame.push(3);
                frame.extend_from_slice(&value.to_le_bytes());
            }
            CellValue::String(value) => {
                frame.push(4);
                write_string(frame, value)?;
            }
            CellValue::Date(value) => {
                frame.push(5);
                write_string(frame, value.format("%Y-%m-%d").to_string())?;
            }
            CellValue::Time(value) => {
                frame.push(6);
                write_string(frame, value.format("%H:%M:%S%.f").to_string())?;
            }
            CellValue::DateTime(value) => {
                frame.push(7);
                write_string(frame, value.format("%Y-%m-%dT%H:%M:%S%.f").to_string())?;
            }
            CellValue::Duration(value) => {
                frame.push(8);
                frame.extend_from_slice(&value.num_milliseconds().to_le_bytes());
            }
            CellValue::Error(value) => {
                frame.push(9);
                write_string(frame, value)?;
            }
        }
    }
    Ok(())
}

fn write_string(frame: &mut Vec<u8>, value: impl AsRef<str>) -> Result<(), i32> {
    let bytes = value.as_ref().as_bytes();
    write_length(frame, bytes.len())?;
    frame.extend_from_slice(bytes);
    Ok(())
}

fn write_strings(values: Vec<String>) -> Result<Vec<u8>, i32> {
    let mut frame = Vec::new();
    write_length(&mut frame, values.len())?;
    for value in values {
        write_string(&mut frame, value)?;
    }
    Ok(frame)
}

fn write_optional_string(frame: &mut Vec<u8>, value: Option<&str>) -> Result<(), i32> {
    frame.push(u8::from(value.is_some()));
    if let Some(value) = value {
        write_string(frame, value)?;
    }
    Ok(())
}

fn write_person(frame: &mut Vec<u8>, person: Option<&CommentPerson>) -> Result<(), i32> {
    frame.push(u8::from(person.is_some()));
    if let Some(person) = person {
        write_string(frame, person.id().to_string())?;
        write_string(frame, person.display_name())?;
        write_optional_string(frame, person.provider_id())?;
    }
    Ok(())
}

fn write_timestamp(frame: &mut Vec<u8>, timestamp: Option<&CommentTimestamp>) -> Result<(), i32> {
    let value = timestamp.map(|value| match value {
        CommentTimestamp::Local(value) => value.format("%Y-%m-%dT%H:%M:%S%.f").to_string(),
        CommentTimestamp::Offset(value) => value.to_rfc3339(),
    });
    write_optional_string(frame, value.as_deref())
}

fn csv_read_options(
    use_header_row: u8,
    delimiter: u8,
    encoding: u8,
    read_empty_as_null: u8,
    trim_headers: u8,
) -> Result<CsvReadOptions, i32> {
    if delimiter == 0 {
        set_last_error("delimiter must be a single-byte character");
        return Err(ERROR_INVALID_ARGUMENT);
    }
    let encoding = parse_csv_encoding(encoding)?;
    let configuration = CsvConfiguration::new()
        .with_delimiter(delimiter)
        .with_encoding(encoding)
        .with_read_empty_as_null(read_empty_as_null != 0);
    Ok(CsvReadOptions::new()
        .with_configuration(configuration)
        .with_header_mode(if use_header_row == 0 {
            HeaderMode::None
        } else {
            HeaderMode::FirstRow
        })
        .with_trim_headers(trim_headers != 0))
}

fn parse_csv_encoding(encoding: u8) -> Result<CsvEncoding, i32> {
    match encoding {
        0 => Ok(CsvEncoding::Utf8),
        1 => Ok(CsvEncoding::Utf16Le),
        2 => Ok(CsvEncoding::Utf16Be),
        3 => Ok(CsvEncoding::Gbk),
        4 => Ok(CsvEncoding::Windows1252),
        _ => {
            set_last_error("encoding is not supported");
            Err(ERROR_INVALID_ARGUMENT)
        }
    }
}

unsafe fn write_csv(
    arguments: CsvWriteArguments,
    append: bool,
    out_row_count: *mut u32,
) -> Result<i32, i32> {
    let CsvWriteArguments {
        path,
        data,
        data_length,
        delimiter,
        encoding,
        write_bom,
        print_header,
        overwrite_file,
    } = arguments;
    if path.is_null() || data.is_null() || out_row_count.is_null() {
        set_last_error("path, data, and out_row_count are required");
        return Err(ERROR_INVALID_ARGUMENT);
    }
    if delimiter == 0 {
        set_last_error("delimiter must be a single-byte character");
        return Err(ERROR_INVALID_ARGUMENT);
    }

    unsafe { ptr::write(out_row_count, 0) };
    let path = unsafe { read_utf8(path) }?;
    let rows = decode_rows(unsafe { std::slice::from_raw_parts(data, data_length) })?;
    let configuration = CsvConfiguration::new()
        .with_delimiter(delimiter)
        .with_encoding(parse_csv_encoding(encoding)?)
        .with_write_bom(write_bom != 0);
    let options = CsvWriteOptions::new()
        .with_configuration(configuration)
        .with_print_header(print_header != 0)
        .with_overwrite_file(overwrite_file != 0);
    let count = if append {
        MiniExcel::append_csv(path, &rows, &options)
    } else {
        MiniExcel::save_csv(path, &rows, &options)
    }
    .map_err(|error| {
        set_last_error(error.to_string());
        ERROR_WRITE
    })?;
    let count = u32::try_from(count).map_err(|_| {
        set_last_error("row count exceeds the ABI limit");
        ERROR_WRITE
    })?;
    unsafe { ptr::write(out_row_count, count) };
    Ok(RESULT_BATCH)
}

fn decode_rows(bytes: &[u8]) -> Result<Vec<DynamicRow>, i32> {
    let mut reader = FrameInput::new(bytes);
    let rows = read_rows(&mut reader)?;
    reader.ensure_complete()?;
    Ok(rows)
}

fn decode_sheets(bytes: &[u8]) -> Result<Vec<(String, Vec<DynamicRow>)>, i32> {
    let mut reader = FrameInput::new(bytes);
    let sheet_count = reader.read_length()?;
    let mut sheets = Vec::with_capacity(sheet_count);
    for _ in 0..sheet_count {
        sheets.push((reader.read_string()?, read_rows(&mut reader)?));
    }
    reader.ensure_complete()?;
    Ok(sheets)
}

fn read_rows(reader: &mut FrameInput<'_>) -> Result<Vec<DynamicRow>, i32> {
    let row_count = reader.read_length()?;
    let mut rows = Vec::with_capacity(row_count);
    for _ in 0..row_count {
        let cell_count = reader.read_length()?;
        let mut row = DynamicRow::with_capacity(cell_count);
        for _ in 0..cell_count {
            let name = reader.read_string()?;
            let value = match reader.read_byte()? {
                0 => CellValue::Empty,
                1 => CellValue::Bool(reader.read_byte()? != 0),
                2 => CellValue::Int(reader.read_i64()?),
                3 => CellValue::Float(f64::from_bits(reader.read_u64()?)),
                4 => CellValue::String(reader.read_string()?),
                tag => {
                    set_last_error(format!("input frame contains unsupported value tag {tag}"));
                    return Err(ERROR_INVALID_ARGUMENT);
                }
            };
            row.insert(name, value);
        }
        rows.push(row);
    }
    Ok(rows)
}

fn insert_options(
    sheet_name: &str,
    print_header: u8,
    replace_existing: u8,
    remove_supported_relationships: u8,
    overwrite_file: bool,
) -> InsertOptions {
    InsertOptions::new()
        .with_sheet_name(sheet_name)
        .with_print_header(print_header != 0)
        .with_existing_sheet_policy(if replace_existing == 0 {
            ExistingSheetPolicy::Reject
        } else {
            ExistingSheetPolicy::Replace
        })
        .with_target_relationship_policy(if remove_supported_relationships == 0 {
            TargetRelationshipPolicy::Reject
        } else {
            TargetRelationshipPolicy::RemoveSupported
        })
        .with_overwrite_file(overwrite_file)
}

fn write_row_count(count: usize, out_row_count: *mut u32) -> Result<i32, i32> {
    let count = u32::try_from(count).map_err(|_| {
        set_last_error("row count exceeds the ABI limit");
        ERROR_WRITE
    })?;
    unsafe { ptr::write(out_row_count, count) };
    Ok(RESULT_BATCH)
}

struct FrameInput<'a> {
    bytes: &'a [u8],
    offset: usize,
}

impl<'a> FrameInput<'a> {
    const fn new(bytes: &'a [u8]) -> Self {
        Self { bytes, offset: 0 }
    }

    fn read_byte(&mut self) -> Result<u8, i32> {
        self.ensure_available(1)?;
        let value = self.bytes[self.offset];
        self.offset += 1;
        Ok(value)
    }

    fn read_u32(&mut self) -> Result<u32, i32> {
        self.ensure_available(4)?;
        let mut value = [0_u8; 4];
        value.copy_from_slice(&self.bytes[self.offset..self.offset + 4]);
        self.offset += 4;
        Ok(u32::from_le_bytes(value))
    }

    fn read_u64(&mut self) -> Result<u64, i32> {
        self.ensure_available(8)?;
        let mut value = [0_u8; 8];
        value.copy_from_slice(&self.bytes[self.offset..self.offset + 8]);
        self.offset += 8;
        Ok(u64::from_le_bytes(value))
    }

    fn read_i64(&mut self) -> Result<i64, i32> {
        self.read_u64()
            .map(|value| i64::from_le_bytes(value.to_le_bytes()))
    }

    fn read_length(&mut self) -> Result<usize, i32> {
        self.read_u32().map(|value| value as usize)
    }

    fn read_string(&mut self) -> Result<String, i32> {
        let length = self.read_length()?;
        self.ensure_available(length)?;
        let value = std::str::from_utf8(&self.bytes[self.offset..self.offset + length])
            .map_err(|error| {
                set_last_error(error.to_string());
                ERROR_INVALID_ARGUMENT
            })?
            .to_owned();
        self.offset += length;
        Ok(value)
    }

    fn ensure_complete(&self) -> Result<(), i32> {
        if self.offset == self.bytes.len() {
            Ok(())
        } else {
            set_last_error("input frame contains trailing data");
            Err(ERROR_INVALID_ARGUMENT)
        }
    }

    fn ensure_available(&self, length: usize) -> Result<(), i32> {
        if self.offset <= self.bytes.len().saturating_sub(length) {
            Ok(())
        } else {
            set_last_error("input frame is truncated");
            Err(ERROR_INVALID_ARGUMENT)
        }
    }
}

fn write_length(frame: &mut Vec<u8>, length: usize) -> Result<(), i32> {
    let length = u32::try_from(length).map_err(|_| {
        set_last_error("FFI frame value exceeds the 4 GiB format limit");
        ERROR_QUERY
    })?;
    write_u32(frame, length);
    Ok(())
}

fn write_u32(frame: &mut Vec<u8>, value: u32) {
    frame.extend_from_slice(&value.to_le_bytes());
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn reports_the_supported_abi_version() {
        assert_eq!(miniexcel_abi_version(), 1);
    }

    #[test]
    fn rejects_missing_required_query_arguments() {
        let result = unsafe {
            miniexcel_query_open(ptr::null(), 0, ptr::null(), ptr::null(), ptr::null_mut())
        };

        assert_eq!(result, ERROR_INVALID_ARGUMENT);

        let mut length = 0;
        let error = unsafe { miniexcel_last_error(&mut length) };
        let message = unsafe { std::slice::from_raw_parts(error, length) };
        assert_eq!(message, b"path, start_cell, and out_handle are required");
    }

    #[test]
    fn rejects_missing_required_sheet_name_arguments() {
        let result = unsafe {
            miniexcel_get_sheet_names(
                ptr::null(),
                ptr::null_mut(),
                ptr::null_mut(),
                ptr::null_mut(),
            )
        };

        assert_eq!(result, ERROR_INVALID_ARGUMENT);

        let mut length = 0;
        let error = unsafe { miniexcel_last_error(&mut length) };
        let message = unsafe { std::slice::from_raw_parts(error, length) };
        assert_eq!(
            message,
            b"path, out_handle, out_data, and out_length are required"
        );
    }
}
