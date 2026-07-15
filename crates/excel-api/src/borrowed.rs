//! Read-only views over values supplied by Excel for one callback invocation.
//!
//! The only entry point from the raw ABI is [`RawExcelValue::from_callback`].
//! It is unsafe because Rust cannot prove the validity or lifetime of the
//! pointer-bearing members of an `XLOPER12`. Once that boundary has been
//! audited by the callback thunk, decoding and observation are safe.
//!
//! Callback views cannot be sent to another thread:
//!
//! ```compile_fail
//! use excel_api::RawExcelValue;
//!
//! fn require_send<T: Send>() {}
//! require_send::<RawExcelValue<'static>>();
//! ```
//!
//! Nor can their callback lifetime be extended:
//!
//! ```compile_fail
//! use excel_api::{ExcelValueRef, RawExcelValue};
//!
//! fn escape(raw: &RawExcelValue<'_>) -> ExcelValueRef<'static> {
//!     raw.decode().unwrap()
//! }
//! ```

use core::{
    marker::PhantomData,
    mem::size_of,
    ptr::{self, NonNull},
    slice,
};

use excel_api_sys::{
    COL, EXCEL12_MAX_COLUMN, EXCEL12_MAX_COLUMNS, EXCEL12_MAX_ROW, EXCEL12_MAX_ROWS,
    EXCEL12_MAX_STRING_CODE_UNITS, IDSHEET, RW, XCHAR, XLBIT_DLL_FREE, XLBIT_XL_FREE, XLOPER12,
    XLREF12, XLTYPE_BIG_DATA, XLTYPE_BOOL, XLTYPE_ERR, XLTYPE_FLOW, XLTYPE_INT, XLTYPE_MISSING,
    XLTYPE_MULTI, XLTYPE_NIL, XLTYPE_NUM, XLTYPE_REF, XLTYPE_SREF, XLTYPE_STR, XlError,
};

use crate::ExcelError;

type CallbackMarker<'call> = PhantomData<(&'call XLOPER12, *mut ())>;

/// Why a raw callback value could not be represented as a safe borrowed view.
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum DecodeError {
    MalformedType(u32),
    UnsupportedType(u32),
    InvalidError(i32),
    NullStringPointer,
    StringTooLong(usize),
    UnterminatedString,
    MalformedArrayDimensions { rows: RW, columns: COL },
    ArrayTooLarge,
    NullArrayPointer,
    NestedArray,
    ReferenceInArray,
    InvalidSingleReferenceCount(u16),
    NullReferencePointer,
    EmptyMultiReference,
    MalformedReferenceArea,
}

/// An `XLOPER12` whose storage is owned by Excel for one callback invocation.
///
/// This wrapper is deliberately neither `Clone`, `Copy`, `Send`, nor `Sync`.
///
/// ```compile_fail
/// use excel_api::RawExcelValue;
/// fn require_sync<T: Sync>() {}
/// require_sync::<RawExcelValue<'static>>();
/// ```
///
/// ```compile_fail
/// use excel_api::RawExcelValue;
/// fn require_clone<T: Clone>() {}
/// require_clone::<RawExcelValue<'static>>();
/// ```
#[derive(Debug)]
pub struct RawExcelValue<'call> {
    raw: NonNull<XLOPER12>,
    _callback: CallbackMarker<'call>,
}

impl<'call> RawExcelValue<'call> {
    /// Establishes the single unsafe boundary for a callback argument.
    ///
    /// # Safety
    ///
    /// `raw` and every pointer-bearing member reachable from it must describe
    /// initialized, properly aligned SDK storage that remains readable and is
    /// not mutated for all of `'call`. Counted strings must include their
    /// prefix and payload, null-terminated direct strings have separate
    /// constructors, multis must contain `rows * columns` initialized elements,
    /// and multi-references must contain `count` initialized areas.
    pub unsafe fn from_callback(raw: &'call XLOPER12) -> Self {
        Self {
            raw: NonNull::from(raw),
            _callback: PhantomData,
        }
    }

    /// Decodes the raw value after masking Excel's ownership bits.
    pub fn decode(&self) -> Result<ExcelValueRef<'call>, DecodeError> {
        // SAFETY: `from_callback` establishes readability and immutability for
        // the complete callback lifetime.
        unsafe { decode_xloper(self.raw, DecodeContext::TopLevel) }
    }
}

/// A decoded, read-only Excel callback value.
///
/// ```compile_fail
/// use excel_api::ExcelValueRef;
/// fn require_send<T: Send>() {}
/// require_send::<ExcelValueRef<'static>>();
/// ```
///
/// ```compile_fail
/// use excel_api::ExcelValueRef;
/// fn require_sync<T: Sync>() {}
/// require_sync::<ExcelValueRef<'static>>();
/// ```
///
/// ```compile_fail
/// use excel_api::ExcelValueRef;
/// fn require_clone<T: Clone>() {}
/// require_clone::<ExcelValueRef<'static>>();
/// ```
#[derive(Debug)]
pub enum ExcelValueRef<'call> {
    Number(f64),
    Text(ExcelStr<'call>),
    Boolean(bool),
    Reference(ExcelReference<'call>),
    Error(ExcelError),
    Array(ExcelArrayView<'call>),
    Missing(ExcelMissing),
    Nil(ExcelNil),
    Integer(i32),
}

impl ExcelValueRef<'_> {
    pub const fn kind_name(&self) -> &'static str {
        match self {
            Self::Number(_) => "number",
            Self::Text(_) => "text",
            Self::Boolean(_) => "boolean",
            Self::Reference(_) => "reference",
            Self::Error(_) => "error",
            Self::Array(_) => "array",
            Self::Missing(_) => "missing",
            Self::Nil(_) => "nil",
            Self::Integer(_) => "integer",
        }
    }
}

/// Borrowed UTF-16 code units. No Unicode conversion is performed.
///
/// ```compile_fail
/// use excel_api::ExcelStr;
/// fn require_send<T: Send>() {}
/// require_send::<ExcelStr<'static>>();
/// ```
///
/// ```compile_fail
/// use excel_api::ExcelStr;
/// fn require_sync<T: Sync>() {}
/// require_sync::<ExcelStr<'static>>();
/// ```
///
/// ```compile_fail
/// use excel_api::ExcelStr;
/// fn require_clone<T: Clone>() {}
/// require_clone::<ExcelStr<'static>>();
/// ```
#[derive(Debug)]
pub struct ExcelStr<'call> {
    payload: NonNull<XCHAR>,
    len: usize,
    _callback: CallbackMarker<'call>,
}

impl<'call> ExcelStr<'call> {
    /// Parses a registered counted direct UTF-16 argument.
    ///
    /// # Safety
    ///
    /// `counted` must be aligned, readable, and immutable for `'call`. It must
    /// contain a readable length prefix followed by that many code units.
    pub unsafe fn from_counted_direct(counted: &'call XCHAR) -> Result<Self, DecodeError> {
        // SAFETY: guaranteed by this function's caller.
        unsafe { parse_direct_counted(NonNull::from(counted)) }
    }

    /// Parses a registered null-terminated direct UTF-16 argument.
    ///
    /// Scanning is bounded by Excel's 32,767-code-unit string limit.
    ///
    /// # Safety
    ///
    /// `terminated` must be aligned, readable, and immutable for `'call`
    /// through the first NUL or through index 32,767 inclusive.
    pub unsafe fn from_null_terminated_direct(
        terminated: &'call XCHAR,
    ) -> Result<Self, DecodeError> {
        // SAFETY: guaranteed by this function's caller.
        unsafe { parse_direct_null_terminated(NonNull::from(terminated)) }
    }

    pub const fn len(&self) -> usize {
        self.len
    }

    pub const fn is_empty(&self) -> bool {
        self.len == 0
    }

    /// Returns the exact UTF-16 payload, excluding a prefix or terminator.
    pub fn as_utf16(&self) -> &'call [u16] {
        // SAFETY: every constructor validates the payload extent, and the
        // callback contract keeps it readable and immutable for `'call`.
        unsafe { slice::from_raw_parts(self.payload.as_ptr(), self.len) }
    }
}

/// Marker for an omitted function argument (`xltypeMissing`).
#[derive(Debug, Eq, PartialEq)]
pub struct ExcelMissing {
    _private: (),
}

/// Marker for an empty Excel value (`xltypeNil`).
#[derive(Debug, Eq, PartialEq)]
pub struct ExcelNil {
    _private: (),
}

/// A flat row-major `xltypeMulti` view.
///
/// ```compile_fail
/// use excel_api::ExcelArrayView;
/// fn require_send<T: Send>() {}
/// require_send::<ExcelArrayView<'static>>();
/// ```
///
/// ```compile_fail
/// use excel_api::ExcelArrayView;
/// fn require_sync<T: Sync>() {}
/// require_sync::<ExcelArrayView<'static>>();
/// ```
///
/// ```compile_fail
/// use excel_api::ExcelArrayView;
/// fn require_clone<T: Clone>() {}
/// require_clone::<ExcelArrayView<'static>>();
/// ```
#[derive(Debug)]
pub struct ExcelArrayView<'call> {
    elements: NonNull<XLOPER12>,
    rows: usize,
    columns: usize,
    _callback: CallbackMarker<'call>,
}

impl<'call> ExcelArrayView<'call> {
    pub const fn dimensions(&self) -> (usize, usize) {
        (self.rows, self.columns)
    }

    pub const fn row_count(&self) -> usize {
        self.rows
    }

    pub const fn column_count(&self) -> usize {
        self.columns
    }

    pub fn get(
        &self,
        row: usize,
        column: usize,
    ) -> Result<Option<ExcelValueRef<'call>>, DecodeError> {
        if row >= self.rows || column >= self.columns {
            return Ok(None);
        }
        let index = row * self.columns + column;
        self.decode_index(index).map(Some)
    }

    pub fn row(&self, row: usize) -> Option<ExcelArrayElements<'_, 'call>> {
        (row < self.rows).then_some(ExcelArrayElements {
            array: self,
            next: row * self.columns,
            remaining: self.columns,
            stride: 1,
        })
    }

    pub fn column(&self, column: usize) -> Option<ExcelArrayElements<'_, 'call>> {
        (column < self.columns).then_some(ExcelArrayElements {
            array: self,
            next: column,
            remaining: self.rows,
            stride: self.columns,
        })
    }

    pub fn rows(&self) -> ExcelArrayRows<'_, 'call> {
        ExcelArrayRows {
            array: self,
            next: 0,
        }
    }

    pub fn columns(&self) -> ExcelArrayColumns<'_, 'call> {
        ExcelArrayColumns {
            array: self,
            next: 0,
        }
    }

    fn decode_index(&self, index: usize) -> Result<ExcelValueRef<'call>, DecodeError> {
        // SAFETY: array construction validates the allocation extent and this
        // method only receives an in-bounds linear index.
        let element = unsafe { NonNull::new_unchecked(self.elements.as_ptr().add(index)) };
        // SAFETY: the array callback contract covers every initialized element.
        unsafe { decode_xloper(element, DecodeContext::ArrayElement) }
    }

    fn validate_elements(&self) -> Result<(), DecodeError> {
        let count = self.rows * self.columns;
        for index in 0..count {
            self.decode_index(index)?;
        }
        Ok(())
    }
}

/// Elements in one row or column of an Excel array.
#[derive(Debug)]
pub struct ExcelArrayElements<'view, 'call> {
    array: &'view ExcelArrayView<'call>,
    next: usize,
    remaining: usize,
    stride: usize,
}

impl<'call> Iterator for ExcelArrayElements<'_, 'call> {
    type Item = Result<ExcelValueRef<'call>, DecodeError>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.remaining == 0 {
            return None;
        }
        let index = self.next;
        self.next += self.stride;
        self.remaining -= 1;
        Some(self.array.decode_index(index))
    }

    fn size_hint(&self) -> (usize, Option<usize>) {
        (self.remaining, Some(self.remaining))
    }
}

impl ExactSizeIterator for ExcelArrayElements<'_, '_> {}

/// Iterator over an array's rows.
#[derive(Debug)]
pub struct ExcelArrayRows<'view, 'call> {
    array: &'view ExcelArrayView<'call>,
    next: usize,
}

impl<'view, 'call> Iterator for ExcelArrayRows<'view, 'call> {
    type Item = ExcelArrayElements<'view, 'call>;

    fn next(&mut self) -> Option<Self::Item> {
        let row = self.next;
        let result = self.array.row(row)?;
        self.next += 1;
        Some(result)
    }

    fn size_hint(&self) -> (usize, Option<usize>) {
        let remaining = self.array.rows - self.next;
        (remaining, Some(remaining))
    }
}

impl ExactSizeIterator for ExcelArrayRows<'_, '_> {}

/// Iterator over an array's columns.
#[derive(Debug)]
pub struct ExcelArrayColumns<'view, 'call> {
    array: &'view ExcelArrayView<'call>,
    next: usize,
}

impl<'view, 'call> Iterator for ExcelArrayColumns<'view, 'call> {
    type Item = ExcelArrayElements<'view, 'call>;

    fn next(&mut self) -> Option<Self::Item> {
        let column = self.next;
        let result = self.array.column(column)?;
        self.next += 1;
        Some(result)
    }

    fn size_hint(&self) -> (usize, Option<usize>) {
        let remaining = self.array.columns - self.next;
        (remaining, Some(remaining))
    }
}

impl ExactSizeIterator for ExcelArrayColumns<'_, '_> {}

/// A reference-preserving callback value.
///
/// ```compile_fail
/// use excel_api::ExcelReference;
/// fn require_send<T: Send>() {}
/// require_send::<ExcelReference<'static>>();
/// ```
///
/// ```compile_fail
/// use excel_api::ExcelReference;
/// fn require_sync<T: Sync>() {}
/// require_sync::<ExcelReference<'static>>();
/// ```
///
/// ```compile_fail
/// use excel_api::ExcelReference;
/// fn require_clone<T: Clone>() {}
/// require_clone::<ExcelReference<'static>>();
/// ```
#[derive(Debug)]
pub enum ExcelReference<'call> {
    Single(ExcelSingleReference<'call>),
    Multiple(ExcelMultiReference<'call>),
}

/// An inline, current-sheet `xltypeSRef`.
#[derive(Debug)]
pub struct ExcelSingleReference<'call> {
    area: ExcelReferenceArea<'call>,
}

impl<'call> ExcelSingleReference<'call> {
    pub const fn area(&self) -> &ExcelReferenceArea<'call> {
        &self.area
    }
}

/// An explicit-sheet, one-or-more-area `xltypeRef`.
#[derive(Debug)]
pub struct ExcelMultiReference<'call> {
    areas: NonNull<XLREF12>,
    area_count: usize,
    sheet_id: IDSHEET,
    _callback: CallbackMarker<'call>,
}

impl<'call> ExcelMultiReference<'call> {
    pub const fn sheet_id(&self) -> IDSHEET {
        self.sheet_id
    }

    pub const fn area_count(&self) -> usize {
        self.area_count
    }

    pub fn area(&self, index: usize) -> Option<ExcelReferenceArea<'call>> {
        if index >= self.area_count {
            return None;
        }
        // SAFETY: the callback contract guarantees `area_count` initialized
        // entries and the bounds check keeps this pointer within them.
        let area = unsafe { NonNull::new_unchecked(self.areas.as_ptr().add(index)) };
        Some(ExcelReferenceArea {
            area,
            _callback: PhantomData,
        })
    }

    pub fn areas(&self) -> ExcelReferenceAreas<'_, 'call> {
        ExcelReferenceAreas {
            reference: self,
            next: 0,
        }
    }
}

/// One rectangular area within a reference.
#[derive(Debug)]
pub struct ExcelReferenceArea<'call> {
    area: NonNull<XLREF12>,
    _callback: CallbackMarker<'call>,
}

impl ExcelReferenceArea<'_> {
    pub fn first_row(&self) -> RW {
        self.read().rwFirst
    }

    pub fn last_row(&self) -> RW {
        self.read().rwLast
    }

    pub fn first_column(&self) -> COL {
        self.read().colFirst
    }

    pub fn last_column(&self) -> COL {
        self.read().colLast
    }

    fn read(&self) -> XLREF12 {
        // SAFETY: construction validates that this area is initialized and the
        // callback contract prevents mutation for its lifetime.
        unsafe { self.area.as_ptr().read() }
    }
}

/// Iterator over the areas of an `xltypeRef`.
#[derive(Debug)]
pub struct ExcelReferenceAreas<'view, 'call> {
    reference: &'view ExcelMultiReference<'call>,
    next: usize,
}

impl<'call> Iterator for ExcelReferenceAreas<'_, 'call> {
    type Item = ExcelReferenceArea<'call>;

    fn next(&mut self) -> Option<Self::Item> {
        let result = self.reference.area(self.next)?;
        self.next += 1;
        Some(result)
    }

    fn size_hint(&self) -> (usize, Option<usize>) {
        let remaining = self.reference.area_count - self.next;
        (remaining, Some(remaining))
    }
}

impl ExactSizeIterator for ExcelReferenceAreas<'_, '_> {}

#[derive(Clone, Copy)]
enum DecodeContext {
    TopLevel,
    ArrayElement,
}

unsafe fn decode_xloper<'call>(
    raw: NonNull<XLOPER12>,
    context: DecodeContext,
) -> Result<ExcelValueRef<'call>, DecodeError> {
    // SAFETY: the caller carries the callback allocation invariant.
    let xltype = unsafe { ptr::addr_of!((*raw.as_ptr()).xltype).read() };
    let base_type = xltype & !(XLBIT_XL_FREE | XLBIT_DLL_FREE);

    match base_type {
        XLTYPE_NUM => {
            // SAFETY: the validated tag selects the `num` union member.
            let value = unsafe { ptr::addr_of!((*raw.as_ptr()).val.num).read() };
            Ok(ExcelValueRef::Number(value))
        }
        XLTYPE_STR => {
            // SAFETY: the validated tag selects the `str` union member.
            let string = unsafe { ptr::addr_of!((*raw.as_ptr()).val.str).read() };
            let string = NonNull::new(string).ok_or(DecodeError::NullStringPointer)?;
            // SAFETY: the callback boundary guarantees the counted payload.
            unsafe { parse_xloper_counted(string).map(ExcelValueRef::Text) }
        }
        XLTYPE_BOOL => {
            // SAFETY: the validated tag selects the `xbool` union member.
            let value = unsafe { ptr::addr_of!((*raw.as_ptr()).val.xbool).read() };
            Ok(ExcelValueRef::Boolean(value != 0))
        }
        XLTYPE_REF => {
            if matches!(context, DecodeContext::ArrayElement) {
                return Err(DecodeError::ReferenceInArray);
            }
            // SAFETY: the validated tag selects the `mref` union member.
            let value = unsafe { ptr::addr_of!((*raw.as_ptr()).val.mref).read() };
            let reference = NonNull::new(value.lpmref).ok_or(DecodeError::NullReferencePointer)?;
            // SAFETY: the callback boundary guarantees a readable XLMREF12.
            let count = unsafe { ptr::addr_of!((*reference.as_ptr()).count).read() } as usize;
            if count == 0 {
                return Err(DecodeError::EmptyMultiReference);
            }
            // SAFETY: `reftbl` begins at this SDK-defined field address.
            let areas = unsafe {
                NonNull::new_unchecked(
                    ptr::addr_of!((*reference.as_ptr()).reftbl)
                        .cast::<XLREF12>()
                        .cast_mut(),
                )
            };
            validate_reference_areas(areas, count)?;
            Ok(ExcelValueRef::Reference(ExcelReference::Multiple(
                ExcelMultiReference {
                    areas,
                    area_count: count,
                    sheet_id: value.idSheet,
                    _callback: PhantomData,
                },
            )))
        }
        XLTYPE_ERR => {
            // SAFETY: the validated tag selects the `err` union member.
            let value = unsafe { ptr::addr_of!((*raw.as_ptr()).val.err).read() };
            decode_error(value).map(ExcelValueRef::Error)
        }
        XLTYPE_FLOW => Err(DecodeError::UnsupportedType(base_type)),
        XLTYPE_MULTI => {
            if matches!(context, DecodeContext::ArrayElement) {
                return Err(DecodeError::NestedArray);
            }
            // SAFETY: the validated tag selects the `array` union member.
            let value = unsafe { ptr::addr_of!((*raw.as_ptr()).val.array).read() };
            let (rows, columns) = validate_array_dimensions(value.rows, value.columns)?;
            let elements = NonNull::new(value.lparray).ok_or(DecodeError::NullArrayPointer)?;
            let array = ExcelArrayView {
                elements,
                rows,
                columns,
                _callback: PhantomData,
            };
            array.validate_elements()?;
            Ok(ExcelValueRef::Array(array))
        }
        XLTYPE_MISSING => Ok(ExcelValueRef::Missing(ExcelMissing { _private: () })),
        XLTYPE_NIL => Ok(ExcelValueRef::Nil(ExcelNil { _private: () })),
        XLTYPE_SREF => {
            if matches!(context, DecodeContext::ArrayElement) {
                return Err(DecodeError::ReferenceInArray);
            }
            // SAFETY: the validated tag selects the `sref` union member.
            let value = unsafe { ptr::addr_of!((*raw.as_ptr()).val.sref).read() };
            if value.count != 1 {
                return Err(DecodeError::InvalidSingleReferenceCount(value.count));
            }
            validate_reference_area(value.reference)?;
            // SAFETY: this is the address of the inline `reference` field and
            // the parent XLOPER12 remains fixed for the callback lifetime.
            let area = unsafe {
                NonNull::new_unchecked(ptr::addr_of!((*raw.as_ptr()).val.sref.reference).cast_mut())
            };
            Ok(ExcelValueRef::Reference(ExcelReference::Single(
                ExcelSingleReference {
                    area: ExcelReferenceArea {
                        area,
                        _callback: PhantomData,
                    },
                },
            )))
        }
        XLTYPE_INT => {
            // SAFETY: the validated tag selects the `w` union member.
            let value = unsafe { ptr::addr_of!((*raw.as_ptr()).val.w).read() };
            Ok(ExcelValueRef::Integer(value))
        }
        XLTYPE_BIG_DATA => Err(DecodeError::UnsupportedType(base_type)),
        _ => Err(DecodeError::MalformedType(base_type)),
    }
}

unsafe fn parse_xloper_counted<'call>(
    counted: NonNull<XCHAR>,
) -> Result<ExcelStr<'call>, DecodeError> {
    // SAFETY: this parser's caller guarantees a readable counted XLOPER string.
    unsafe { parse_counted_payload(counted) }
}

unsafe fn parse_direct_counted<'call>(
    counted: NonNull<XCHAR>,
) -> Result<ExcelStr<'call>, DecodeError> {
    // SAFETY: this parser's caller guarantees a readable direct counted string.
    unsafe { parse_counted_payload(counted) }
}

unsafe fn parse_counted_payload<'call>(
    counted: NonNull<XCHAR>,
) -> Result<ExcelStr<'call>, DecodeError> {
    // SAFETY: the prefix is initialized and readable by the parser contract.
    let len = unsafe { counted.as_ptr().read() } as usize;
    if len > EXCEL12_MAX_STRING_CODE_UNITS {
        return Err(DecodeError::StringTooLong(len));
    }
    // SAFETY: a counted string always has a prefix, so adding one stays at the
    // payload start even when the logical payload is empty.
    let payload = unsafe { NonNull::new_unchecked(counted.as_ptr().add(1)) };
    Ok(ExcelStr {
        payload,
        len,
        _callback: PhantomData,
    })
}

unsafe fn parse_direct_null_terminated<'call>(
    terminated: NonNull<XCHAR>,
) -> Result<ExcelStr<'call>, DecodeError> {
    for len in 0..=EXCEL12_MAX_STRING_CODE_UNITS {
        // SAFETY: the parser contract guarantees readability through this
        // bounded scan range or the first terminator.
        if unsafe { terminated.as_ptr().add(len).read() } == 0 {
            return Ok(ExcelStr {
                payload: terminated,
                len,
                _callback: PhantomData,
            });
        }
    }
    Err(DecodeError::UnterminatedString)
}

fn validate_array_dimensions(rows: RW, columns: COL) -> Result<(usize, usize), DecodeError> {
    if !(1..=EXCEL12_MAX_ROWS).contains(&rows) || !(1..=EXCEL12_MAX_COLUMNS).contains(&columns) {
        return Err(DecodeError::MalformedArrayDimensions { rows, columns });
    }
    let rows = usize::try_from(rows).map_err(|_| DecodeError::ArrayTooLarge)?;
    let columns = usize::try_from(columns).map_err(|_| DecodeError::ArrayTooLarge)?;
    let count = rows
        .checked_mul(columns)
        .ok_or(DecodeError::ArrayTooLarge)?;
    if count > isize::MAX as usize / size_of::<XLOPER12>() {
        return Err(DecodeError::ArrayTooLarge);
    }
    Ok((rows, columns))
}

fn validate_reference_areas(areas: NonNull<XLREF12>, count: usize) -> Result<(), DecodeError> {
    for index in 0..count {
        // SAFETY: the multi-reference callback contract guarantees `count`
        // initialized entries.
        let area = unsafe { areas.as_ptr().add(index).read() };
        validate_reference_area(area)?;
    }
    Ok(())
}

fn validate_reference_area(area: XLREF12) -> Result<(), DecodeError> {
    let rows_valid = (0..=EXCEL12_MAX_ROW).contains(&area.rwFirst)
        && (area.rwFirst..=EXCEL12_MAX_ROW).contains(&area.rwLast);
    let columns_valid = (0..=EXCEL12_MAX_COLUMN).contains(&area.colFirst)
        && (area.colFirst..=EXCEL12_MAX_COLUMN).contains(&area.colLast);
    if rows_valid && columns_valid {
        Ok(())
    } else {
        Err(DecodeError::MalformedReferenceArea)
    }
}

fn decode_error(value: i32) -> Result<ExcelError, DecodeError> {
    let value = match value {
        excel_api_sys::xlerrNull => XlError::Null,
        excel_api_sys::xlerrDiv0 => XlError::Div0,
        excel_api_sys::xlerrValue => XlError::Value,
        excel_api_sys::xlerrRef => XlError::Ref,
        excel_api_sys::xlerrName => XlError::Name,
        excel_api_sys::xlerrNum => XlError::Num,
        excel_api_sys::xlerrNA => XlError::Na,
        excel_api_sys::xlerrGettingData => XlError::GettingData,
        _ => return Err(DecodeError::InvalidError(value)),
    };
    Ok(value.into())
}

#[cfg(test)]
mod tests {
    use core::ffi::c_void;

    use excel_api_sys::{
        XLMREF12, XLOPER12Array, XLOPER12BigData, XLOPER12BigDataHandle, XLOPER12MRef,
        XLOPER12SRef, XLOPER12Value, xlbitDLLFree, xlbitXLFree, xltypeBigData, xltypeBool,
        xltypeErr, xltypeFlow, xltypeInt, xltypeMissing, xltypeMulti, xltypeNil, xltypeNum,
        xltypeRef, xltypeSRef, xltypeStr,
    };

    use super::*;

    fn raw(value: XLOPER12Value, xltype: u32) -> XLOPER12 {
        XLOPER12 { val: value, xltype }
    }

    fn decode(value: &XLOPER12) -> Result<ExcelValueRef<'_>, DecodeError> {
        // SAFETY: each test keeps all referenced fixtures alive and immutable
        // for the returned view's lifetime.
        unsafe { RawExcelValue::from_callback(value) }.decode()
    }

    fn zeroed_value() -> XLOPER12Value {
        // SAFETY: every union member consists only of integer or raw-pointer
        // fields, for which the all-zero bit pattern is valid.
        unsafe { core::mem::zeroed() }
    }

    #[test]
    fn decodes_scalar_tags_and_masks_ownership_bits() {
        let number = raw(
            XLOPER12Value { num: 2.5 },
            xltypeNum | xlbitXLFree | xlbitDLLFree,
        );
        assert!(matches!(decode(&number), Ok(ExcelValueRef::Number(2.5))));

        let boolean = raw(XLOPER12Value { xbool: -1 }, xltypeBool);
        assert!(matches!(decode(&boolean), Ok(ExcelValueRef::Boolean(true))));

        let integer = raw(XLOPER12Value { w: -42 }, xltypeInt);
        assert!(matches!(decode(&integer), Ok(ExcelValueRef::Integer(-42))));

        for (raw_error, expected) in [
            (excel_api_sys::xlerrNull, ExcelError::Null),
            (excel_api_sys::xlerrDiv0, ExcelError::Div0),
            (excel_api_sys::xlerrValue, ExcelError::Value),
            (excel_api_sys::xlerrRef, ExcelError::Ref),
            (excel_api_sys::xlerrName, ExcelError::Name),
            (excel_api_sys::xlerrNum, ExcelError::Num),
            (excel_api_sys::xlerrNA, ExcelError::Na),
            (excel_api_sys::xlerrGettingData, ExcelError::GettingData),
        ] {
            let error = raw(XLOPER12Value { err: raw_error }, xltypeErr);
            let Ok(ExcelValueRef::Error(actual)) = decode(&error) else {
                panic!("expected error value");
            };
            assert_eq!(actual, expected);
        }
    }

    #[test]
    fn missing_and_nil_are_independent() {
        let missing = raw(XLOPER12Value { w: 0 }, xltypeMissing);
        let nil = raw(XLOPER12Value { w: 0 }, xltypeNil);
        assert!(matches!(decode(&missing), Ok(ExcelValueRef::Missing(_))));
        assert!(matches!(decode(&nil), Ok(ExcelValueRef::Nil(_))));
    }

    #[test]
    fn parses_all_three_string_abis_without_conversion() {
        let mut xloper_text = [3, b'A' as u16, 0, 0xD800];
        let value = raw(
            XLOPER12Value {
                str: xloper_text.as_mut_ptr(),
            },
            xltypeStr,
        );
        let Ok(ExcelValueRef::Text(text)) = decode(&value) else {
            panic!("expected text");
        };
        assert_eq!(text.as_utf16(), &[b'A' as u16, 0, 0xD800]);

        let counted = [2, 0xD800, b'Z' as u16];
        // SAFETY: `counted` contains its prefix and complete payload.
        let direct = unsafe { ExcelStr::from_counted_direct(&counted[0]) }.unwrap();
        assert_eq!(direct.as_utf16(), &[0xD800, b'Z' as u16]);

        let terminated = [b'X' as u16, b'Y' as u16, 0, b'Z' as u16];
        // SAFETY: the fixture contains a terminator within the bounded scan.
        let direct = unsafe { ExcelStr::from_null_terminated_direct(&terminated[0]) }.unwrap();
        assert_eq!(direct.as_utf16(), &[b'X' as u16, b'Y' as u16]);

        let mut empty = [0_u16];
        let value = raw(
            XLOPER12Value {
                str: empty.as_mut_ptr(),
            },
            xltypeStr,
        );
        let Ok(ExcelValueRef::Text(empty)) = decode(&value) else {
            panic!("expected empty text");
        };
        assert!(empty.is_empty());
        assert_eq!(empty.as_utf16(), &[]);
    }

    #[test]
    fn rejects_malformed_strings() {
        let mut too_long = [u16::MAX, 0];
        let value = raw(
            XLOPER12Value {
                str: too_long.as_mut_ptr(),
            },
            xltypeStr,
        );
        assert!(matches!(
            decode(&value),
            Err(DecodeError::StringTooLong(65_535))
        ));

        let null = raw(
            XLOPER12Value {
                str: ptr::null_mut(),
            },
            xltypeStr,
        );
        assert!(matches!(decode(&null), Err(DecodeError::NullStringPointer)));

        let unterminated = [1_u16; EXCEL12_MAX_STRING_CODE_UNITS + 1];
        // SAFETY: the entire bounded scan range is initialized.
        let result = unsafe { ExcelStr::from_null_terminated_direct(&unterminated[0]) };
        assert!(matches!(result, Err(DecodeError::UnterminatedString)));
    }

    #[test]
    fn array_supports_indexing_rows_and_columns() {
        let mut elements = [
            raw(XLOPER12Value { num: 1.0 }, xltypeNum),
            raw(XLOPER12Value { num: 2.0 }, xltypeNum),
            raw(XLOPER12Value { num: 3.0 }, xltypeNum),
            raw(XLOPER12Value { num: 4.0 }, xltypeNum),
        ];
        let value = raw(
            XLOPER12Value {
                array: XLOPER12Array {
                    lparray: elements.as_mut_ptr(),
                    rows: 2,
                    columns: 2,
                },
            },
            xltypeMulti,
        );
        let Ok(ExcelValueRef::Array(array)) = decode(&value) else {
            panic!("expected array");
        };
        assert_eq!(array.dimensions(), (2, 2));
        assert!(matches!(
            array.get(1, 0),
            Ok(Some(ExcelValueRef::Number(3.0)))
        ));
        assert!(matches!(array.get(2, 0), Ok(None)));

        let second_row: Vec<_> = array.row(1).unwrap().collect();
        assert!(matches!(second_row[0], Ok(ExcelValueRef::Number(3.0))));
        assert!(matches!(second_row[1], Ok(ExcelValueRef::Number(4.0))));

        let second_column: Vec<_> = array.column(1).unwrap().collect();
        assert!(matches!(second_column[0], Ok(ExcelValueRef::Number(2.0))));
        assert!(matches!(second_column[1], Ok(ExcelValueRef::Number(4.0))));
        assert_eq!(array.rows().len(), 2);
        assert_eq!(array.columns().len(), 2);
    }

    #[test]
    fn rejects_malformed_array_dimensions_and_pointer() {
        for (rows, columns) in [(0, 1), (1, 0), (-1, 1), (1, -1), (EXCEL12_MAX_ROWS + 1, 1)] {
            let value = raw(
                XLOPER12Value {
                    array: XLOPER12Array {
                        lparray: ptr::null_mut(),
                        rows,
                        columns,
                    },
                },
                xltypeMulti,
            );
            assert!(matches!(
                decode(&value),
                Err(DecodeError::MalformedArrayDimensions {
                    rows: actual_rows,
                    columns: actual_columns,
                }) if actual_rows == rows && actual_columns == columns
            ));
        }

        let value = raw(
            XLOPER12Value {
                array: XLOPER12Array {
                    lparray: ptr::null_mut(),
                    rows: 1,
                    columns: 1,
                },
            },
            xltypeMulti,
        );
        assert!(matches!(decode(&value), Err(DecodeError::NullArrayPointer)));
    }

    #[test]
    fn rejects_nested_arrays_and_references_in_arrays() {
        let mut nested = raw(
            XLOPER12Value {
                array: XLOPER12Array {
                    lparray: ptr::null_mut(),
                    rows: 1,
                    columns: 1,
                },
            },
            xltypeMulti,
        );
        let outer = raw(
            XLOPER12Value {
                array: XLOPER12Array {
                    lparray: &mut nested,
                    rows: 1,
                    columns: 1,
                },
            },
            xltypeMulti,
        );
        assert!(matches!(decode(&outer), Err(DecodeError::NestedArray)));

        let mut reference = raw(
            XLOPER12Value {
                sref: XLOPER12SRef {
                    count: 1,
                    reference: XLREF12::default(),
                },
            },
            xltypeSRef,
        );
        let outer = raw(
            XLOPER12Value {
                array: XLOPER12Array {
                    lparray: &mut reference,
                    rows: 1,
                    columns: 1,
                },
            },
            xltypeMulti,
        );
        assert!(matches!(decode(&outer), Err(DecodeError::ReferenceInArray)));
    }

    #[test]
    fn represents_single_and_multi_references_distinctly() {
        let single = raw(
            XLOPER12Value {
                sref: XLOPER12SRef {
                    count: 1,
                    reference: XLREF12 {
                        rwFirst: 1,
                        rwLast: 2,
                        colFirst: 3,
                        colLast: 4,
                    },
                },
            },
            xltypeSRef,
        );
        let Ok(ExcelValueRef::Reference(ExcelReference::Single(single))) = decode(&single) else {
            panic!("expected single reference");
        };
        assert_eq!(single.area().first_row(), 1);
        assert_eq!(single.area().last_column(), 4);

        let mut mref = XLMREF12 {
            count: 1,
            reftbl: [XLREF12 {
                rwFirst: 5,
                rwLast: 6,
                colFirst: 7,
                colLast: 8,
            }],
        };
        let multiple = raw(
            XLOPER12Value {
                mref: XLOPER12MRef {
                    lpmref: &mut mref,
                    idSheet: 99,
                },
            },
            xltypeRef,
        );
        let Ok(ExcelValueRef::Reference(ExcelReference::Multiple(multiple))) = decode(&multiple)
        else {
            panic!("expected multi reference");
        };
        assert_eq!(multiple.sheet_id(), 99);
        assert_eq!(multiple.area_count(), 1);
        assert_eq!(multiple.areas().next().unwrap().first_column(), 7);
    }

    #[test]
    fn multi_reference_exposes_every_sdk_tail_area() {
        #[repr(C)]
        struct TwoAreaReference {
            count: u16,
            padding: u16,
            areas: [XLREF12; 2],
        }

        let mut storage = TwoAreaReference {
            count: 2,
            padding: 0,
            areas: [
                XLREF12 {
                    rwFirst: 1,
                    rwLast: 1,
                    colFirst: 2,
                    colLast: 2,
                },
                XLREF12 {
                    rwFirst: 3,
                    rwLast: 4,
                    colFirst: 5,
                    colLast: 6,
                },
            ],
        };
        let value = raw(
            XLOPER12Value {
                mref: XLOPER12MRef {
                    lpmref: (&mut storage as *mut TwoAreaReference).cast::<XLMREF12>(),
                    idSheet: 7,
                },
            },
            xltypeRef,
        );
        let Ok(ExcelValueRef::Reference(ExcelReference::Multiple(reference))) = decode(&value)
        else {
            panic!("expected multi reference");
        };
        assert_eq!(reference.area_count(), 2);
        assert_eq!(reference.area(0).unwrap().first_row(), 1);
        assert_eq!(reference.area(1).unwrap().last_column(), 6);
        assert!(reference.area(2).is_none());
    }

    #[test]
    fn validates_reference_shapes() {
        let invalid_count = raw(
            XLOPER12Value {
                sref: XLOPER12SRef {
                    count: 2,
                    reference: XLREF12::default(),
                },
            },
            xltypeSRef,
        );
        assert!(matches!(
            decode(&invalid_count),
            Err(DecodeError::InvalidSingleReferenceCount(2))
        ));

        let invalid_area = raw(
            XLOPER12Value {
                sref: XLOPER12SRef {
                    count: 1,
                    reference: XLREF12 {
                        rwFirst: 2,
                        rwLast: 1,
                        colFirst: 0,
                        colLast: 0,
                    },
                },
            },
            xltypeSRef,
        );
        assert!(matches!(
            decode(&invalid_area),
            Err(DecodeError::MalformedReferenceArea)
        ));
    }

    #[test]
    fn every_sdk_type_is_decoded_or_explicitly_rejected() {
        let flow = raw(XLOPER12Value { w: 0 }, xltypeFlow);
        assert!(matches!(
            decode(&flow),
            Err(DecodeError::UnsupportedType(value)) if value == xltypeFlow
        ));

        let big_data = raw(
            XLOPER12Value {
                bigdata: XLOPER12BigData {
                    h: XLOPER12BigDataHandle {
                        hdata: ptr::null_mut::<c_void>(),
                    },
                    cbData: 0,
                },
            },
            xltypeBigData,
        );
        assert!(matches!(
            decode(&big_data),
            Err(DecodeError::UnsupportedType(value)) if value == xltypeBigData
        ));

        let invalid_error = raw(XLOPER12Value { err: -1 }, xltypeErr);
        assert!(matches!(
            decode(&invalid_error),
            Err(DecodeError::InvalidError(-1))
        ));
    }

    #[test]
    fn malformed_tags_are_not_interpreted_as_union_members() {
        for xltype in [
            0,
            xltypeNum | xltypeBool,
            xlbitXLFree,
            xlbitDLLFree,
            0x2000,
            0x8000_0000 | xltypeNum,
        ] {
            let value = raw(zeroed_value(), xltype);
            assert!(matches!(decode(&value), Err(DecodeError::MalformedType(_))));
        }
    }

    #[test]
    fn malformed_type_fuzz_never_panics() {
        for xltype in 0_u32..=u16::MAX as u32 {
            // A zeroed union is a valid raw-pointer/integer bit pattern. Any
            // pointer-bearing recognized tag is rejected before dereference.
            let value = raw(zeroed_value(), xltype);
            let _ = decode(&value);
        }
        for xltype in [u32::MAX, 0x8000_0001, 0x0002_0000, 0xDEAD_BEEF] {
            let value = raw(zeroed_value(), xltype);
            let _ = decode(&value);
        }
    }
}
