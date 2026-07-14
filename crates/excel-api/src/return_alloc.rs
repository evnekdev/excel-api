//! Stable Excel return storage and its consuming handoff to Excel.
//!
//! Materialization creates a locally owned ABI tree with base type bits only.
//! [`ExcelReturn::into_raw_for_excel`] consumes that local owner, marks only
//! the root with `xlbitDLLFree`, and transfers the allocation to Excel. The
//! matching [`xl_auto_free12`] callback reconstructs and drops the exact
//! root-first Rust allocation.

use core::mem::size_of;

use excel_api_sys::{
    XCHAR, XLOPER12, XLOPER12Array, XLOPER12Value, XLTYPE_MASK, xlbitDLLFree, xlbitXLFree,
    xltypeBool, xltypeErr, xltypeInt, xltypeMissing, xltypeMulti, xltypeNil, xltypeNum, xltypeStr,
};

use crate::{
    ExcelError, PlannedArrayElement, PlannedText, PlannedValue, ReturnMaterializationError,
    ReturnOwnershipStrategy, ReturnPlan, ReturnStorageTotals, ReturnText,
};

/// Opaque local owner of one stable Excel ABI return tree.
///
/// Dropping this value while it is local releases the root and every nested
/// backing allocation through ordinary Rust field drops. Consuming handoff
/// leaves no Rust owner behind; Excel must then invoke the matching callback
/// exactly once.
///
/// The root is deliberately read-only:
///
/// ```compile_fail
/// use excel_api::ExcelReturn;
/// fn cannot_mutate(mut value: ExcelReturn) {
///     value.as_xloper().xltype = 0;
/// }
/// ```
///
/// Handoff consumes the owner, so safe code cannot hand it off twice:
///
/// ```compile_fail
/// fn cannot_handoff_twice(value: excel_api::ExcelReturn) {
///     let _first = value.into_raw_for_excel();
///     let _second = value.into_raw_for_excel();
/// }
/// ```
///
/// Raw reclamation is not a safe ownership constructor:
///
/// ```compile_fail
/// fn cannot_reclaim_safely(pointer: *mut excel_api_sys::XLOPER12) {
///     excel_api::xl_auto_free12(pointer);
/// }
/// ```
pub struct ExcelReturn {
    allocation: Box<ReturnAllocation>,
}

impl ExcelReturn {
    pub fn from_plan(plan: ReturnPlan) -> Result<Self, ReturnMaterializationError> {
        Materializer::new().materialize(plan)
    }

    /// Read-only access to the stable root for inspection and future handoff.
    pub fn as_xloper(&self) -> &XLOPER12 {
        &self.allocation.root
    }

    /// Consumes this local owner and transfers its complete allocation to Excel.
    ///
    /// The returned pointer is the root at byte offset zero of the same
    /// `ReturnAllocation`; it is not a copy. Only the root receives
    /// `xlbitDLLFree`. Every nested value retains base type bits only and its
    /// storage remains owned by the leaked top-level allocation until the
    /// matching [`xl_auto_free12`] call.
    pub fn into_raw_for_excel(self) -> *mut XLOPER12 {
        let mut allocation = self.allocation;
        let original_type = allocation.root.xltype;
        let base_type = original_type & XLTYPE_MASK;

        // Prompt 05 materialization establishes these invariants. Keep every
        // potentially panicking check while the Box still has a local owner.
        debug_assert_eq!(original_type & xlbitDLLFree, 0);
        debug_assert_eq!(original_type & xlbitXLFree, 0);
        debug_assert_eq!(original_type, base_type);
        debug_assert!(is_supported_root_type(base_type));

        allocation.root.xltype = original_type | xlbitDLLFree;

        #[cfg(test)]
        HANDED_OFF_ROOTS.fetch_add(1, core::sync::atomic::Ordering::SeqCst);

        let allocation_pointer = allocation.as_mut() as *mut ReturnAllocation;
        let root_pointer = &mut allocation.root as *mut XLOPER12;
        debug_assert_eq!(root_pointer.cast::<ReturnAllocation>(), allocation_pointer);

        let _allocation_pointer = Box::into_raw(allocation);
        root_pointer
    }
}

const fn is_supported_root_type(xltype: u32) -> bool {
    xltype == xltypeNum
        || xltype == xltypeInt
        || xltype == xltypeBool
        || xltype == xltypeErr
        || xltype == xltypeMissing
        || xltype == xltypeNil
        || xltype == xltypeStr
        || xltype == xltypeMulti
}

/// Reusable implementation of the Excel 12 DLL-owned return cleanup callback.
///
/// A final XLL should export an exact `xlAutoFree12` symbol with the SDK's
/// `WINAPI` calling convention and delegate directly to this function. Null is
/// accepted as a defensive no-op, although Excel is not expected to supply it
/// during correct operation. Panics are contained and discarded; no panic can
/// unwind through this FFI boundary when unwinding is enabled. A build using
/// `panic = "abort"` still aborts, because aborting panics cannot be caught.
///
/// # Safety
///
/// A non-null `pointer` must be the pointer returned exactly once by
/// [`ExcelReturn::into_raw_for_excel`] from this same loaded binary, must not
/// have been reclaimed already, and must no longer be accessed by Excel after
/// this callback begins. Arbitrary non-null pointers and duplicate callbacks
/// cannot be validated safely and violate the Excel/XLL ownership contract.
pub unsafe extern "system" fn xl_auto_free12(pointer: *mut XLOPER12) {
    if pointer.is_null() {
        return;
    }

    #[cfg(test)]
    contain_callback_panic(|| {
        if PANIC_BEFORE_RECLAIM.swap(false, core::sync::atomic::Ordering::SeqCst) {
            panic!("injected callback-boundary panic");
        }
    });

    contain_callback_panic(|| {
        // SAFETY: upheld by this callback's non-null handed-off-pointer
        // contract. This closure is invoked exactly once for this call.
        unsafe { reclaim_excel_return(pointer) }
    });
}

fn contain_callback_panic(action: impl FnOnce()) {
    let _ = std::panic::catch_unwind(std::panic::AssertUnwindSafe(action));
}

/// Reconstructs the exact root-first owner and drops the complete return tree.
///
/// # Safety
///
/// `pointer` must be non-null, must have been returned exactly once by
/// `ExcelReturn::into_raw_for_excel`, must not have been reclaimed before, and
/// must point at a `ReturnAllocation` with the current binary's compatible
/// layout and offset-zero root. Excel must perform no further access after
/// entering the callback.
unsafe fn reclaim_excel_return(pointer: *mut XLOPER12) {
    let allocation = pointer.cast::<ReturnAllocation>();
    // SAFETY: the caller guarantees this is the unique pointer produced by
    // Box::into_raw for this exact ReturnAllocation. Dropping its Rust fields
    // releases all nested storage without inspecting raw ABI tags.
    drop(unsafe { Box::from_raw(allocation) });

    #[cfg(test)]
    {
        let previous = HANDED_OFF_ROOTS.fetch_sub(1, core::sync::atomic::Ordering::SeqCst);
        assert!(previous > 0, "callback reclaimed a non-handed-off root");
        CALLBACK_FREES.fetch_add(1, core::sync::atomic::Ordering::SeqCst);
    }
}

impl ReturnPlan {
    pub fn materialize(self) -> Result<ExcelReturn, ReturnMaterializationError> {
        ExcelReturn::from_plan(self)
    }
}

/// Root-first owner. Raw ABI pointers target only the boxed storage owned by
/// later fields; local cleanup uses these Rust fields and never walks unions.
#[repr(C)]
struct ReturnAllocation {
    root: XLOPER12,
    array_elements: Option<ReturnArrayBuffer>,
    string_buffers: Box<[ReturnUtf16Buffer]>,
    #[cfg(test)]
    _tracker: RootTracker,
}

struct ReturnUtf16Buffer {
    storage: Box<[XCHAR]>,
}

impl ReturnUtf16Buffer {
    fn from_planned(text: PlannedText) -> Result<Self, ReturnMaterializationError> {
        let PlannedText {
            source,
            payload_code_units,
            storage_code_units,
        } = text;

        let mut storage = Vec::new();
        storage.try_reserve_exact(storage_code_units).map_err(|_| {
            ReturnMaterializationError::AllocationFailure {
                storage: "counted UTF-16 string",
            }
        })?;
        storage.push(0);

        let utf8_source = matches!(&source, ReturnText::Utf8(_));
        match source {
            ReturnText::Utf8(value) => storage.extend(value.encode_utf16()),
            ReturnText::Utf16(value) => storage.extend_from_slice(value.as_utf16()),
        }

        let actual_payload = storage.len() - 1;
        if utf8_source && actual_payload != payload_code_units {
            return Err(ReturnMaterializationError::Utf8EncodedLengthMismatch {
                planned: payload_code_units,
                actual: actual_payload,
            });
        }
        if actual_payload != payload_code_units || storage.len() != storage_code_units {
            return Err(ReturnMaterializationError::StringBufferLengthMismatch {
                planned: storage_code_units,
                actual: storage.len(),
            });
        }
        let prefix = XCHAR::try_from(actual_payload).map_err(|_| {
            ReturnMaterializationError::StringBufferLengthMismatch {
                planned: storage_code_units,
                actual: storage.len(),
            }
        })?;
        storage[0] = prefix;

        #[cfg(test)]
        LIVE_STRING_BUFFERS.fetch_add(1, core::sync::atomic::Ordering::SeqCst);
        Ok(Self {
            storage: storage.into_boxed_slice(),
        })
    }

    fn abi_ptr(&self) -> *mut XCHAR {
        // The Excel ABI uses a mutable pointer type, but the materialized
        // counted string is immutable until the later consuming handoff.
        self.storage.as_ptr().cast_mut()
    }

    fn payload_units(&self) -> usize {
        self.storage.len() - 1
    }

    fn storage_units(&self) -> usize {
        self.storage.len()
    }
}

#[cfg(test)]
impl Drop for ReturnUtf16Buffer {
    fn drop(&mut self) {
        let previous = LIVE_STRING_BUFFERS.fetch_sub(1, core::sync::atomic::Ordering::SeqCst);
        assert!(previous > 0, "return string buffer dropped more than once");
    }
}

struct ReturnArrayBuffer {
    storage: Box<[XLOPER12]>,
}

impl ReturnArrayBuffer {
    fn allocate(elements: usize) -> Result<Self, ReturnMaterializationError> {
        let mut storage = Vec::new();
        storage.try_reserve_exact(elements).map_err(|_| {
            ReturnMaterializationError::AllocationFailure {
                storage: "XLOPER12 array elements",
            }
        })?;
        storage.resize(elements, raw_empty());

        #[cfg(test)]
        LIVE_ARRAY_BUFFERS.fetch_add(1, core::sync::atomic::Ordering::SeqCst);
        Ok(Self {
            storage: storage.into_boxed_slice(),
        })
    }

    fn abi_ptr(&self) -> *mut XLOPER12 {
        // The allocation is boxed and never resized. The ABI pointer remains
        // stable while the owning ReturnAllocation is alive.
        self.storage.as_ptr().cast_mut()
    }
}

#[cfg(test)]
impl Drop for ReturnArrayBuffer {
    fn drop(&mut self) {
        let previous = LIVE_ARRAY_BUFFERS.fetch_sub(1, core::sync::atomic::Ordering::SeqCst);
        assert!(previous > 0, "return array buffer dropped more than once");
    }
}

#[derive(Clone, Debug)]
enum RootBlueprint {
    Number(f64),
    Integer(i32),
    Boolean(bool),
    Error(ExcelError),
    Missing,
    Empty,
    Text {
        buffer: usize,
    },
    Array {
        rows: usize,
        columns: usize,
        elements: Box<[ElementBlueprint]>,
    },
}

#[derive(Clone, Debug)]
enum ElementBlueprint {
    Number(f64),
    Integer(i32),
    Boolean(bool),
    Error(ExcelError),
    Missing,
    Empty,
    Text { buffer: usize },
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum FailurePoint {
    BeforeFirstString,
    AfterStringAllocations(usize),
    BeforeArrayAllocation,
    AfterArrayAllocation,
    DuringElementInitialization(usize),
    BeforeRootCreation,
}

impl FailurePoint {
    #[cfg(test)]
    const fn label(self) -> &'static str {
        match self {
            Self::BeforeFirstString => "before first string allocation",
            Self::AfterStringAllocations(_) => "after string allocation",
            Self::BeforeArrayAllocation => "before array allocation",
            Self::AfterArrayAllocation => "after array allocation",
            Self::DuringElementInitialization(_) => "during array element initialization",
            Self::BeforeRootCreation => "before final root creation",
        }
    }
}

struct Materializer {
    #[cfg(test)]
    failure: Option<FailurePoint>,
    strings_allocated: usize,
}

impl Materializer {
    fn new() -> Self {
        Self {
            #[cfg(test)]
            failure: None,
            strings_allocated: 0,
        }
    }

    #[cfg(test)]
    fn with_failure(failure: FailurePoint) -> Self {
        Self {
            failure: Some(failure),
            strings_allocated: 0,
        }
    }

    fn materialize(mut self, plan: ReturnPlan) -> Result<ExcelReturn, ReturnMaterializationError> {
        let ReturnPlan {
            root,
            totals,
            strategy,
        } = plan;
        if strategy != ReturnOwnershipStrategy::DllOwnedXloper12 {
            return Err(ReturnMaterializationError::UnsupportedPlannedValue {
                variant: "ownership strategy",
            });
        }

        let text_count = count_texts(&root);
        let mut strings = Vec::new();
        strings.try_reserve_exact(text_count).map_err(|_| {
            ReturnMaterializationError::AllocationFailure {
                storage: "string-buffer owner table",
            }
        })?;
        let blueprint = self.prepare_root(root, &mut strings)?;
        let string_buffers = strings.into_boxed_slice();

        let mut array_elements = self.materialize_array(&blueprint, &string_buffers)?;
        verify_totals(&totals, &string_buffers, array_elements.as_ref())?;
        self.fail_if(FailurePoint::BeforeRootCreation)?;

        let root = raw_root(&blueprint, &string_buffers, array_elements.as_ref())?;
        let allocation = ReturnAllocation {
            root,
            array_elements: array_elements.take(),
            string_buffers,
            #[cfg(test)]
            _tracker: RootTracker::new(),
        };
        Ok(ExcelReturn {
            allocation: Box::new(allocation),
        })
    }

    fn prepare_root(
        &mut self,
        root: PlannedValue,
        strings: &mut Vec<ReturnUtf16Buffer>,
    ) -> Result<RootBlueprint, ReturnMaterializationError> {
        Ok(match root {
            PlannedValue::Number(value) => RootBlueprint::Number(value),
            PlannedValue::Integer(value) => RootBlueprint::Integer(value),
            PlannedValue::Boolean(value) => RootBlueprint::Boolean(value),
            PlannedValue::Error(value) => RootBlueprint::Error(value),
            PlannedValue::Missing => RootBlueprint::Missing,
            PlannedValue::Empty => RootBlueprint::Empty,
            PlannedValue::Text(text) => RootBlueprint::Text {
                buffer: self.push_string(text, strings)?,
            },
            PlannedValue::Array(array) => {
                let expected = array.rows.checked_mul(array.columns).ok_or(
                    ReturnMaterializationError::ArrayShapeMismatch {
                        rows: array.rows,
                        columns: array.columns,
                        elements: array.elements.len(),
                    },
                )?;
                if expected != array.elements.len() {
                    return Err(ReturnMaterializationError::ArrayShapeMismatch {
                        rows: array.rows,
                        columns: array.columns,
                        elements: array.elements.len(),
                    });
                }
                let mut elements = Vec::new();
                elements.try_reserve_exact(expected).map_err(|_| {
                    ReturnMaterializationError::AllocationFailure {
                        storage: "array element blueprints",
                    }
                })?;
                for element in array.elements.into_vec() {
                    elements.push(self.prepare_element(element, strings)?);
                }
                RootBlueprint::Array {
                    rows: array.rows,
                    columns: array.columns,
                    elements: elements.into_boxed_slice(),
                }
            }
        })
    }

    fn prepare_element(
        &mut self,
        element: PlannedArrayElement,
        strings: &mut Vec<ReturnUtf16Buffer>,
    ) -> Result<ElementBlueprint, ReturnMaterializationError> {
        Ok(match element {
            PlannedArrayElement::Number(value) => ElementBlueprint::Number(value),
            PlannedArrayElement::Integer(value) => ElementBlueprint::Integer(value),
            PlannedArrayElement::Boolean(value) => ElementBlueprint::Boolean(value),
            PlannedArrayElement::Error(value) => ElementBlueprint::Error(value),
            PlannedArrayElement::Missing => ElementBlueprint::Missing,
            PlannedArrayElement::Empty => ElementBlueprint::Empty,
            PlannedArrayElement::Text(text) => ElementBlueprint::Text {
                buffer: self.push_string(text, strings)?,
            },
        })
    }

    fn push_string(
        &mut self,
        text: PlannedText,
        strings: &mut Vec<ReturnUtf16Buffer>,
    ) -> Result<usize, ReturnMaterializationError> {
        if self.strings_allocated == 0 {
            self.fail_if(FailurePoint::BeforeFirstString)?;
        }
        let buffer = ReturnUtf16Buffer::from_planned(text)?;
        strings.push(buffer);
        self.strings_allocated += 1;
        self.fail_if(FailurePoint::AfterStringAllocations(self.strings_allocated))?;
        Ok(strings.len() - 1)
    }

    fn materialize_array(
        &self,
        blueprint: &RootBlueprint,
        strings: &[ReturnUtf16Buffer],
    ) -> Result<Option<ReturnArrayBuffer>, ReturnMaterializationError> {
        let RootBlueprint::Array { elements, .. } = blueprint else {
            return Ok(None);
        };
        self.fail_if(FailurePoint::BeforeArrayAllocation)?;
        let mut storage = ReturnArrayBuffer::allocate(elements.len())?;
        self.fail_if(FailurePoint::AfterArrayAllocation)?;
        for (index, (destination, element)) in
            storage.storage.iter_mut().zip(elements.iter()).enumerate()
        {
            self.fail_if(FailurePoint::DuringElementInitialization(index))?;
            *destination = raw_element(element, strings)?;
        }
        Ok(Some(storage))
    }

    fn fail_if(&self, point: FailurePoint) -> Result<(), ReturnMaterializationError> {
        #[cfg(test)]
        if self.failure == Some(point) {
            return Err(ReturnMaterializationError::InjectedTestFailure {
                stage: point.label(),
            });
        }
        let _ = point;
        Ok(())
    }
}

fn count_texts(root: &PlannedValue) -> usize {
    match root {
        PlannedValue::Text(_) => 1,
        PlannedValue::Array(array) => array
            .elements
            .iter()
            .filter(|element| matches!(element, PlannedArrayElement::Text(_)))
            .count(),
        _ => 0,
    }
}

fn verify_totals(
    planned: &ReturnStorageTotals,
    strings: &[ReturnUtf16Buffer],
    array: Option<&ReturnArrayBuffer>,
) -> Result<(), ReturnMaterializationError> {
    let root_bytes = size_of::<XLOPER12>();
    let array_element_bytes = array.map_or(0, |value| value.storage.len() * root_bytes);
    let string_payload_code_units = strings.iter().map(ReturnUtf16Buffer::payload_units).sum();
    let string_storage_code_units = strings.iter().map(ReturnUtf16Buffer::storage_units).sum();
    let string_storage_bytes = string_storage_code_units * size_of::<XCHAR>();
    let total_bytes = root_bytes + array_element_bytes + string_storage_bytes;
    let allocation_count = 1 + usize::from(array.is_some()) + strings.len();

    for (field, expected, actual) in [
        ("root_bytes", planned.root_bytes, root_bytes),
        (
            "array_element_bytes",
            planned.array_element_bytes,
            array_element_bytes,
        ),
        (
            "string_payload_code_units",
            planned.string_payload_code_units,
            string_payload_code_units,
        ),
        (
            "string_storage_code_units",
            planned.string_storage_code_units,
            string_storage_code_units,
        ),
        (
            "string_storage_bytes",
            planned.string_storage_bytes,
            string_storage_bytes,
        ),
        ("total_bytes", planned.total_bytes, total_bytes),
        (
            "allocation_count",
            planned.allocation_count,
            allocation_count,
        ),
    ] {
        if expected != actual {
            return Err(ReturnMaterializationError::PlanStorageInvariantMismatch {
                field,
                planned: expected,
                actual,
            });
        }
    }
    Ok(())
}

fn raw_root(
    blueprint: &RootBlueprint,
    strings: &[ReturnUtf16Buffer],
    array: Option<&ReturnArrayBuffer>,
) -> Result<XLOPER12, ReturnMaterializationError> {
    Ok(match blueprint {
        RootBlueprint::Number(value) => raw_number(*value),
        RootBlueprint::Integer(value) => raw_integer(*value),
        RootBlueprint::Boolean(value) => raw_boolean(*value),
        RootBlueprint::Error(value) => raw_error(*value),
        RootBlueprint::Missing => raw_missing(),
        RootBlueprint::Empty => raw_empty(),
        RootBlueprint::Text { buffer } => {
            let buffer = strings.get(*buffer).ok_or(
                ReturnMaterializationError::UnsupportedPlannedValue {
                    variant: "missing root string buffer",
                },
            )?;
            raw_text(buffer.abi_ptr())
        }
        RootBlueprint::Array {
            rows,
            columns,
            elements,
        } => {
            let storage = array.ok_or(ReturnMaterializationError::UnsupportedPlannedValue {
                variant: "missing array element buffer",
            })?;
            if storage.storage.len() != elements.len() {
                return Err(ReturnMaterializationError::ArrayShapeMismatch {
                    rows: *rows,
                    columns: *columns,
                    elements: storage.storage.len(),
                });
            }
            let planned_rows = *rows;
            let planned_columns = *columns;
            let rows = i32::try_from(planned_rows).map_err(|_| {
                ReturnMaterializationError::ArrayShapeMismatch {
                    rows: planned_rows,
                    columns: planned_columns,
                    elements: elements.len(),
                }
            })?;
            let columns = i32::try_from(planned_columns).map_err(|_| {
                ReturnMaterializationError::ArrayShapeMismatch {
                    rows: planned_rows,
                    columns: planned_columns,
                    elements: elements.len(),
                }
            })?;
            raw_array(storage.abi_ptr(), rows, columns)
        }
    })
}

fn raw_element(
    element: &ElementBlueprint,
    strings: &[ReturnUtf16Buffer],
) -> Result<XLOPER12, ReturnMaterializationError> {
    Ok(match element {
        ElementBlueprint::Number(value) => raw_number(*value),
        ElementBlueprint::Integer(value) => raw_integer(*value),
        ElementBlueprint::Boolean(value) => raw_boolean(*value),
        ElementBlueprint::Error(value) => raw_error(*value),
        ElementBlueprint::Missing => raw_missing(),
        ElementBlueprint::Empty => raw_empty(),
        ElementBlueprint::Text { buffer } => {
            let buffer = strings.get(*buffer).ok_or(
                ReturnMaterializationError::UnsupportedPlannedValue {
                    variant: "missing array string buffer",
                },
            )?;
            raw_text(buffer.abi_ptr())
        }
    })
}

fn zeroed_value() -> XLOPER12Value {
    // SAFETY: every member of XLOPER12Value consists solely of integers,
    // floating-point values, or raw pointers. The all-zero bit pattern is
    // valid for each member. Zeroing the complete union also defines all bytes
    // not occupied by the subsequently selected active member.
    unsafe { core::mem::zeroed() }
}

fn raw_number(value: f64) -> XLOPER12 {
    let mut raw = zeroed_value();
    raw.num = value;
    XLOPER12 {
        val: raw,
        xltype: xltypeNum,
    }
}

fn raw_integer(value: i32) -> XLOPER12 {
    let mut raw = zeroed_value();
    raw.w = value;
    XLOPER12 {
        val: raw,
        xltype: xltypeInt,
    }
}

fn raw_boolean(value: bool) -> XLOPER12 {
    let mut raw = zeroed_value();
    raw.xbool = i32::from(value);
    XLOPER12 {
        val: raw,
        xltype: xltypeBool,
    }
}

fn raw_error(value: ExcelError) -> XLOPER12 {
    let mut raw = zeroed_value();
    raw.err = match value {
        ExcelError::Null => excel_api_sys::xlerrNull,
        ExcelError::Div0 => excel_api_sys::xlerrDiv0,
        ExcelError::Value => excel_api_sys::xlerrValue,
        ExcelError::Ref => excel_api_sys::xlerrRef,
        ExcelError::Name => excel_api_sys::xlerrName,
        ExcelError::Num => excel_api_sys::xlerrNum,
        ExcelError::Na => excel_api_sys::xlerrNA,
        ExcelError::GettingData => excel_api_sys::xlerrGettingData,
    };
    XLOPER12 {
        val: raw,
        xltype: xltypeErr,
    }
}

fn raw_missing() -> XLOPER12 {
    XLOPER12 {
        val: zeroed_value(),
        xltype: xltypeMissing,
    }
}

fn raw_empty() -> XLOPER12 {
    XLOPER12 {
        val: zeroed_value(),
        xltype: xltypeNil,
    }
}

fn raw_text(pointer: *mut XCHAR) -> XLOPER12 {
    let mut raw = zeroed_value();
    raw.str = pointer;
    XLOPER12 {
        val: raw,
        xltype: xltypeStr,
    }
}

fn raw_array(pointer: *mut XLOPER12, rows: i32, columns: i32) -> XLOPER12 {
    let mut raw = zeroed_value();
    raw.array = XLOPER12Array {
        lparray: pointer,
        rows,
        columns,
    };
    XLOPER12 {
        val: raw,
        xltype: xltypeMulti,
    }
}

#[cfg(test)]
use core::sync::atomic::AtomicUsize;

#[cfg(test)]
static LIVE_ROOTS: AtomicUsize = AtomicUsize::new(0);
#[cfg(test)]
static LIVE_STRING_BUFFERS: AtomicUsize = AtomicUsize::new(0);
#[cfg(test)]
static LIVE_ARRAY_BUFFERS: AtomicUsize = AtomicUsize::new(0);
#[cfg(test)]
static HANDED_OFF_ROOTS: AtomicUsize = AtomicUsize::new(0);
#[cfg(test)]
static CALLBACK_FREES: AtomicUsize = AtomicUsize::new(0);
#[cfg(test)]
static PANIC_BEFORE_RECLAIM: core::sync::atomic::AtomicBool =
    core::sync::atomic::AtomicBool::new(false);

#[cfg(test)]
struct RootTracker;

#[cfg(test)]
impl RootTracker {
    fn new() -> Self {
        LIVE_ROOTS.fetch_add(1, core::sync::atomic::Ordering::SeqCst);
        Self
    }
}

#[cfg(test)]
impl Drop for RootTracker {
    fn drop(&mut self) {
        let previous = LIVE_ROOTS.fetch_sub(1, core::sync::atomic::Ordering::SeqCst);
        assert!(previous > 0, "return root dropped more than once");
    }
}

#[cfg(test)]
mod tests {
    use core::{mem::offset_of, slice};
    use std::sync::{Mutex, MutexGuard};

    use excel_api_sys::{xlbitDLLFree, xlbitXLFree};

    use super::*;
    use crate::{ExcelReturnArray, ExcelReturnValue, ExcelString, ReturnText};

    static TEST_LOCK: Mutex<()> = Mutex::new(());

    fn lock() -> MutexGuard<'static, ()> {
        TEST_LOCK.lock().unwrap_or_else(|error| error.into_inner())
    }

    fn live_counts() -> (usize, usize, usize) {
        use core::sync::atomic::Ordering;

        (
            LIVE_ROOTS.load(Ordering::SeqCst),
            LIVE_STRING_BUFFERS.load(Ordering::SeqCst),
            LIVE_ARRAY_BUFFERS.load(Ordering::SeqCst),
        )
    }

    fn handoff_counts() -> (usize, usize) {
        use core::sync::atomic::Ordering;

        (
            HANDED_OFF_ROOTS.load(Ordering::SeqCst),
            CALLBACK_FREES.load(Ordering::SeqCst),
        )
    }

    fn assert_no_live_or_handed_off_allocations() {
        assert_eq!(live_counts(), (0, 0, 0));
        assert_eq!(handoff_counts().0, 0);
    }

    fn assert_no_ownership_bits(value: &XLOPER12) {
        assert_eq!(value.xltype & (xlbitDLLFree | xlbitXLFree), 0);
    }

    fn number(value: &XLOPER12) -> f64 {
        assert_eq!(value.xltype, xltypeNum);
        // SAFETY: the asserted base tag selects `val.num`; the complete union
        // was initialized by the centralized raw constructor.
        unsafe { value.val.num }
    }

    fn integer(value: &XLOPER12) -> i32 {
        assert_eq!(value.xltype, xltypeInt);
        // SAFETY: the asserted base tag selects `val.w`; the complete union was
        // initialized by the centralized raw constructor.
        unsafe { value.val.w }
    }

    fn boolean(value: &XLOPER12) -> i32 {
        assert_eq!(value.xltype, xltypeBool);
        // SAFETY: the asserted base tag selects `val.xbool`; the complete union
        // was initialized by the centralized raw constructor.
        unsafe { value.val.xbool }
    }

    fn error_code(value: &XLOPER12) -> i32 {
        assert_eq!(value.xltype, xltypeErr);
        // SAFETY: the asserted base tag selects `val.err`; the complete union
        // was initialized by the centralized raw constructor.
        unsafe { value.val.err }
    }

    fn counted_storage<'owner>(_owner: &'owner ExcelReturn, value: &XLOPER12) -> &'owner [XCHAR] {
        assert_eq!(value.xltype, xltypeStr);
        // SAFETY: the tag selects `val.str`. Every string XLOPER produced by
        // this module points to a nonempty, aligned Box<[XCHAR]> owned by
        // `_owner`. The prefix is readable and gives the exact remaining
        // initialized extent. `_owner` keeps the allocation alive and fixed.
        unsafe {
            let pointer = value.val.str;
            let payload = usize::from(pointer.read());
            slice::from_raw_parts(pointer, payload + 1)
        }
    }

    fn array_elements(owner: &ExcelReturn) -> &[XLOPER12] {
        let root = owner.as_xloper();
        assert_eq!(root.xltype, xltypeMulti);
        // SAFETY: the tag selects `val.array`. Its pointer targets the stable
        // boxed element buffer owned by `owner`; rows and columns were checked
        // and their product is the exact initialized element count.
        unsafe {
            let array = root.val.array;
            let count =
                usize::try_from(array.rows).unwrap() * usize::try_from(array.columns).unwrap();
            slice::from_raw_parts(array.lparray, count)
        }
    }

    fn array_pointer(value: &XLOPER12) -> *mut XLOPER12 {
        assert_eq!(value.xltype, xltypeMulti);
        // SAFETY: the asserted base tag selects `val.array`; only its pointer
        // value is copied and no pointee is accessed here.
        unsafe { value.val.array.lparray }
    }

    fn string_pointer(value: &XLOPER12) -> *mut XCHAR {
        assert_eq!(value.xltype, xltypeStr);
        // SAFETY: the asserted base tag selects `val.str`; only its pointer
        // value is copied and no pointee is accessed here.
        unsafe { value.val.str }
    }

    fn string_pointer_with_flags(value: &XLOPER12) -> *mut XCHAR {
        assert_eq!(value.xltype & XLTYPE_MASK, xltypeStr);
        // SAFETY: the masked base tag selects `val.str`; only the pointer is
        // copied while the handed-off allocation remains alive.
        unsafe { value.val.str }
    }

    fn array_pointer_with_flags(value: &XLOPER12) -> *mut XLOPER12 {
        assert_eq!(value.xltype & XLTYPE_MASK, xltypeMulti);
        // SAFETY: the masked base tag selects `val.array`; only its pointer is
        // copied while the handed-off allocation remains alive.
        unsafe { value.val.array.lparray }
    }

    unsafe fn raw_counted_storage(value: &XLOPER12) -> &[XCHAR] {
        assert_eq!(value.xltype & XLTYPE_MASK, xltypeStr);
        // SAFETY: the caller keeps the handed-off ReturnAllocation alive, and
        // this module created a nonempty counted buffer whose prefix gives its
        // exact initialized extent.
        unsafe {
            let pointer = value.val.str;
            let payload = usize::from(pointer.read());
            slice::from_raw_parts(pointer, payload + 1)
        }
    }

    unsafe fn raw_array_elements(value: &XLOPER12) -> &[XLOPER12] {
        assert_eq!(value.xltype & XLTYPE_MASK, xltypeMulti);
        // SAFETY: the caller keeps the handed-off ReturnAllocation alive. The
        // materialized dimensions are positive and their product is the exact
        // initialized length of the boxed element block.
        unsafe {
            let array = value.val.array;
            let count =
                usize::try_from(array.rows).unwrap() * usize::try_from(array.columns).unwrap();
            slice::from_raw_parts(array.lparray, count)
        }
    }

    fn handoff_and_reclaim(returned: ExcelReturn) {
        let base_type = returned.as_xloper().xltype;
        let root_address = returned.as_xloper() as *const XLOPER12 as *mut XLOPER12;
        let callback_frees_before = handoff_counts().1;

        let pointer = returned.into_raw_for_excel();
        assert_eq!(pointer, root_address);
        // SAFETY: `pointer` has just been handed off and remains live until the
        // one callback invocation below.
        let root = unsafe { &*pointer };
        assert_eq!(root.xltype & XLTYPE_MASK, base_type);
        assert_ne!(root.xltype & xlbitDLLFree, 0);
        assert_eq!(root.xltype & xlbitXLFree, 0);
        assert_eq!(handoff_counts(), (1, callback_frees_before));

        // SAFETY: this is the unique pointer just produced above and this is
        // its sole callback reclamation.
        unsafe { xl_auto_free12(pointer) };
        assert_no_live_or_handed_off_allocations();
        assert_eq!(handoff_counts().1, callback_frees_before + 1);
    }

    fn move_return(value: ExcelReturn) -> ExcelReturn {
        value
    }

    fn materialize(value: ExcelReturnValue) -> ExcelReturn {
        value.plan().unwrap().materialize().unwrap()
    }

    fn mixed_array_plan() -> ReturnPlan {
        ExcelReturnValue::Array(
            ExcelReturnArray::new(
                2,
                4,
                vec![
                    ExcelReturnValue::Number(1.5),
                    ExcelReturnValue::Integer(-2),
                    ExcelReturnValue::Boolean(true),
                    ExcelReturnValue::Error(ExcelError::Ref),
                    ExcelReturnValue::Missing,
                    ExcelReturnValue::Empty,
                    ExcelReturnValue::Text(ReturnText::from("A\0B")),
                    ExcelReturnValue::Text(ReturnText::Utf16(ExcelString::from_utf16_units([
                        0xD800,
                    ]))),
                ],
            )
            .unwrap(),
        )
        .plan()
        .unwrap()
    }

    #[test]
    fn root_is_offset_zero_and_scalars_use_exact_sdk_members() {
        let _guard = lock();
        assert_eq!(offset_of!(ReturnAllocation, root), 0);

        let returned = materialize(ExcelReturnValue::Number(2.5));
        assert_eq!(
            returned.as_xloper() as *const XLOPER12,
            returned.allocation.as_ref() as *const ReturnAllocation as *const XLOPER12
        );
        assert_eq!(number(returned.as_xloper()), 2.5);
        assert_no_ownership_bits(returned.as_xloper());
        drop(returned);

        let returned = materialize(ExcelReturnValue::Integer(-42));
        assert_eq!(integer(returned.as_xloper()), -42);
        assert_no_ownership_bits(returned.as_xloper());
        drop(returned);

        for (input, expected) in [(false, 0), (true, 1)] {
            let returned = materialize(ExcelReturnValue::Boolean(input));
            assert_eq!(boolean(returned.as_xloper()), expected);
            assert_no_ownership_bits(returned.as_xloper());
        }

        let missing = materialize(ExcelReturnValue::Missing);
        assert_eq!(missing.as_xloper().xltype, xltypeMissing);
        assert_no_ownership_bits(missing.as_xloper());
        let empty = materialize(ExcelReturnValue::Empty);
        assert_eq!(empty.as_xloper().xltype, xltypeNil);
        assert_no_ownership_bits(empty.as_xloper());
        drop((missing, empty));
        assert_eq!(live_counts(), (0, 0, 0));
    }

    #[test]
    fn every_excel_error_uses_the_exact_sdk_code() {
        let _guard = lock();
        for (error, expected) in [
            (ExcelError::Null, excel_api_sys::xlerrNull),
            (ExcelError::Div0, excel_api_sys::xlerrDiv0),
            (ExcelError::Value, excel_api_sys::xlerrValue),
            (ExcelError::Ref, excel_api_sys::xlerrRef),
            (ExcelError::Name, excel_api_sys::xlerrName),
            (ExcelError::Num, excel_api_sys::xlerrNum),
            (ExcelError::Na, excel_api_sys::xlerrNA),
            (ExcelError::GettingData, excel_api_sys::xlerrGettingData),
        ] {
            let returned = materialize(ExcelReturnValue::Error(error));
            assert_eq!(error_code(returned.as_xloper()), expected);
            assert_no_ownership_bits(returned.as_xloper());
        }
        assert_eq!(live_counts(), (0, 0, 0));
    }

    #[test]
    fn counted_strings_preserve_every_utf16_form_without_terminator() {
        let _guard = lock();
        let cases: Vec<ReturnText> = vec![
            ReturnText::from(""),
            ReturnText::from("ASCII"),
            ReturnText::from("é水"),
            ReturnText::from("😃"),
            ReturnText::Utf16(ExcelString::from_utf16_units([0xD83D, 0xDE03])),
            ReturnText::Utf16(ExcelString::from_utf16_units([0xD800])),
            ReturnText::Utf16(ExcelString::from_utf16_units([0xDC00])),
            ReturnText::Utf16(ExcelString::from_utf16_units([b'A' as u16, 0, b'B' as u16])),
        ];

        for source in cases {
            let expected: Vec<u16> = match &source {
                ReturnText::Utf8(value) => value.encode_utf16().collect(),
                ReturnText::Utf16(value) => value.as_utf16().to_vec(),
            };
            let returned = materialize(ExcelReturnValue::Text(source));
            let storage = counted_storage(&returned, returned.as_xloper());
            assert_eq!(
                string_pointer(returned.as_xloper()),
                returned.allocation.string_buffers[0]
                    .storage
                    .as_ptr()
                    .cast_mut()
            );
            assert_eq!(usize::from(storage[0]), expected.len());
            assert_eq!(&storage[1..], expected);
            assert_eq!(
                returned.allocation.string_buffers[0].storage.len(),
                expected.len() + 1
            );
            assert_no_ownership_bits(returned.as_xloper());
        }
        assert_eq!(live_counts(), (0, 0, 0));
    }

    #[test]
    fn maximum_counted_string_uses_exact_prefix_and_payload() {
        let _guard = lock();
        let maximum = excel_api_sys::EXCEL12_MAX_STRING_CODE_UNITS;
        let payload = vec![0xD800; maximum];
        let returned = materialize(ExcelReturnValue::Text(ReturnText::Utf16(
            ExcelString::from_utf16_units(payload.clone()),
        )));
        let storage = counted_storage(&returned, returned.as_xloper());
        assert_eq!(usize::from(storage[0]), maximum);
        assert_eq!(&storage[1..], payload);
        assert_eq!(storage.len(), maximum + 1);
        drop(returned);
        assert_eq!(live_counts(), (0, 0, 0));
    }

    #[test]
    fn arrays_are_one_stable_row_major_block_with_deep_strings() {
        let _guard = lock();
        let returned = mixed_array_plan().materialize().unwrap();
        let root = returned.as_xloper();
        assert_no_ownership_bits(root);
        // SAFETY: the root tag is xltypeMulti and therefore selects val.array.
        let raw_array = unsafe { root.val.array };
        assert_eq!((raw_array.rows, raw_array.columns), (2, 4));

        let elements = array_elements(&returned);
        assert_eq!(elements.as_ptr(), raw_array.lparray);
        assert_eq!(number(&elements[0]), 1.5);
        assert_eq!(integer(&elements[1]), -2);
        assert_eq!(boolean(&elements[2]), 1);
        assert_eq!(error_code(&elements[3]), excel_api_sys::xlerrRef);
        assert_eq!(elements[4].xltype, xltypeMissing);
        assert_eq!(elements[5].xltype, xltypeNil);
        assert_eq!(&counted_storage(&returned, &elements[6])[1..], &[65, 0, 66]);
        assert_eq!(&counted_storage(&returned, &elements[7])[1..], &[0xD800]);
        assert_eq!(
            string_pointer(&elements[6]),
            returned.allocation.string_buffers[0]
                .storage
                .as_ptr()
                .cast_mut()
        );
        assert_eq!(
            string_pointer(&elements[7]),
            returned.allocation.string_buffers[1]
                .storage
                .as_ptr()
                .cast_mut()
        );
        for element in elements {
            assert_no_ownership_bits(element);
        }
        assert_eq!(
            returned
                .allocation
                .array_elements
                .as_ref()
                .unwrap()
                .storage
                .len(),
            8
        );
        assert_eq!(returned.allocation.string_buffers.len(), 2);
        drop(returned);
        assert_eq!(live_counts(), (0, 0, 0));
    }

    #[test]
    fn one_by_one_array_and_empty_nested_text_are_supported() {
        let _guard = lock();
        let plan = ExcelReturnValue::Array(
            ExcelReturnArray::new(1, 1, vec![ExcelReturnValue::Text(ReturnText::from(""))])
                .unwrap(),
        )
        .plan()
        .unwrap();
        let returned = plan.materialize().unwrap();
        let elements = array_elements(&returned);
        assert_eq!(elements.len(), 1);
        assert_eq!(counted_storage(&returned, &elements[0]), &[0]);
        drop(returned);
        assert_eq!(live_counts(), (0, 0, 0));
    }

    #[test]
    fn moving_excel_return_preserves_every_published_address() {
        let _guard = lock();
        let returned = mixed_array_plan().materialize().unwrap();
        let root_address = returned.as_xloper() as *const XLOPER12;
        let element_address = array_pointer(returned.as_xloper());
        let elements = array_elements(&returned);
        let string_address = string_pointer(&elements[6]);

        let returned = move_return(returned);
        assert_eq!(returned.as_xloper() as *const XLOPER12, root_address);
        assert_eq!(array_pointer(returned.as_xloper()), element_address);
        assert_eq!(array_elements(&returned).as_ptr(), element_address);
        assert_eq!(
            string_pointer(&array_elements(&returned)[6]),
            string_address
        );
        drop(returned);
        assert_eq!(live_counts(), (0, 0, 0));
    }

    #[test]
    fn materialization_verifies_every_planned_storage_total() {
        let _guard = lock();
        let plan = mixed_array_plan();
        let totals = plan.totals;
        let returned = plan.materialize().unwrap();
        let allocation = &returned.allocation;
        assert_eq!(totals.root_bytes, size_of::<XLOPER12>());
        assert_eq!(
            totals.array_element_bytes,
            allocation.array_elements.as_ref().unwrap().storage.len() * size_of::<XLOPER12>()
        );
        assert_eq!(
            totals.string_payload_code_units,
            allocation
                .string_buffers
                .iter()
                .map(ReturnUtf16Buffer::payload_units)
                .sum()
        );
        assert_eq!(
            totals.string_storage_code_units,
            allocation
                .string_buffers
                .iter()
                .map(ReturnUtf16Buffer::storage_units)
                .sum()
        );
        assert_eq!(
            totals.string_storage_bytes,
            totals.string_storage_code_units * size_of::<XCHAR>()
        );
        assert_eq!(
            totals.total_bytes,
            totals.root_bytes + totals.array_element_bytes + totals.string_storage_bytes
        );
        assert_eq!(totals.allocation_count, 4);
        drop(returned);
        assert_eq!(live_counts(), (0, 0, 0));

        let mut corrupted = mixed_array_plan();
        corrupted.totals.root_bytes += 1;
        assert!(matches!(
            corrupted.materialize(),
            Err(ReturnMaterializationError::PlanStorageInvariantMismatch {
                field: "root_bytes",
                ..
            })
        ));
        assert_eq!(live_counts(), (0, 0, 0));
    }

    #[test]
    fn corrupted_text_and_array_plans_fail_before_publication() {
        let _guard = lock();
        let mut utf8 = ExcelReturnValue::from("abc").plan().unwrap();
        let PlannedValue::Text(text) = &mut utf8.root else {
            panic!("expected text");
        };
        text.payload_code_units += 1;
        assert!(matches!(
            utf8.materialize(),
            Err(ReturnMaterializationError::Utf8EncodedLengthMismatch { .. })
        ));

        let mut utf16 =
            ExcelReturnValue::Text(ReturnText::Utf16(ExcelString::from_utf16_units([0xD800])))
                .plan()
                .unwrap();
        let PlannedValue::Text(text) = &mut utf16.root else {
            panic!("expected text");
        };
        text.storage_code_units += 1;
        assert!(matches!(
            utf16.materialize(),
            Err(ReturnMaterializationError::StringBufferLengthMismatch { .. })
        ));

        let mut array = ExcelReturnValue::Array(
            ExcelReturnArray::new(1, 1, vec![ExcelReturnValue::Empty]).unwrap(),
        )
        .plan()
        .unwrap();
        let PlannedValue::Array(planned_array) = &mut array.root else {
            panic!("expected array");
        };
        planned_array.rows = 2;
        assert!(matches!(
            array.materialize(),
            Err(ReturnMaterializationError::ArrayShapeMismatch { .. })
        ));
        assert_eq!(live_counts(), (0, 0, 0));
    }

    #[test]
    fn every_injected_failure_releases_partial_storage() {
        let _guard = lock();
        for point in [
            FailurePoint::BeforeFirstString,
            FailurePoint::AfterStringAllocations(1),
            FailurePoint::AfterStringAllocations(2),
            FailurePoint::BeforeArrayAllocation,
            FailurePoint::AfterArrayAllocation,
            FailurePoint::DuringElementInitialization(0),
            FailurePoint::DuringElementInitialization(4),
            FailurePoint::BeforeRootCreation,
        ] {
            assert_eq!(live_counts(), (0, 0, 0));
            let result = Materializer::with_failure(point).materialize(mixed_array_plan());
            assert!(matches!(
                result,
                Err(ReturnMaterializationError::InjectedTestFailure { .. })
            ));
            assert_eq!(live_counts(), (0, 0, 0), "failed at {point:?}");
        }
    }

    #[test]
    fn local_drop_and_repeated_construction_release_exactly_once() {
        let _guard = lock();
        let callback_frees_before = handoff_counts().1;
        let returned = mixed_array_plan().materialize().unwrap();
        assert_eq!(live_counts(), (1, 2, 1));
        drop(returned);
        assert_eq!(live_counts(), (0, 0, 0));

        for _ in 0..1_000 {
            let returned = mixed_array_plan().materialize().unwrap();
            assert_eq!(live_counts(), (1, 2, 1));
            drop(returned);
            assert_eq!(live_counts(), (0, 0, 0));
        }
        assert_eq!(handoff_counts().0, 0);
        assert_eq!(handoff_counts().1, callback_frees_before);
    }

    #[test]
    fn panic_before_handoff_keeps_raii_cleanup_local() {
        let _guard = lock();
        let callback_frees_before = handoff_counts().1;
        let result = std::panic::catch_unwind(|| {
            let _returned = mixed_array_plan().materialize().unwrap();
            panic!("injected panic before handoff");
        });
        assert!(result.is_err());
        assert_no_live_or_handed_off_allocations();
        assert_eq!(handoff_counts().1, callback_frees_before);
    }

    #[test]
    fn handoff_and_callback_cover_every_scalar_root_and_error() {
        let _guard = lock();
        let callback_frees_before = handoff_counts().1;
        let mut values = vec![
            ExcelReturnValue::Number(2.5),
            ExcelReturnValue::Integer(-42),
            ExcelReturnValue::Boolean(false),
            ExcelReturnValue::Boolean(true),
            ExcelReturnValue::Missing,
            ExcelReturnValue::Empty,
        ];
        values.extend(
            [
                ExcelError::Null,
                ExcelError::Div0,
                ExcelError::Value,
                ExcelError::Ref,
                ExcelError::Name,
                ExcelError::Num,
                ExcelError::Na,
                ExcelError::GettingData,
            ]
            .map(ExcelReturnValue::Error),
        );

        for value in values {
            handoff_and_reclaim(materialize(value));
        }
        assert_eq!(handoff_counts().1, callback_frees_before + 14);
    }

    #[test]
    fn handoff_preserves_counted_string_pointers_and_arbitrary_utf16() {
        let _guard = lock();
        let maximum = excel_api_sys::EXCEL12_MAX_STRING_CODE_UNITS;
        for payload in [
            vec![b'A' as u16, 0, b'B' as u16],
            vec![0xD800, 0xDC00],
            vec![0xD800],
            vec![0xDC00],
            vec![0xD800; maximum],
        ] {
            let returned = materialize(ExcelReturnValue::Text(ReturnText::Utf16(
                ExcelString::from_utf16_units(payload.clone()),
            )));
            let root_address = returned.as_xloper() as *const XLOPER12 as *mut XLOPER12;
            let string_address = string_pointer(returned.as_xloper());
            let callback_frees_before = handoff_counts().1;

            let pointer = returned.into_raw_for_excel();
            assert_eq!(pointer, root_address);
            // SAFETY: this pointer and its nested string remain owned by the
            // handed-off allocation until the callback below.
            let root = unsafe { &*pointer };
            assert_eq!(root.xltype & XLTYPE_MASK, xltypeStr);
            assert_ne!(root.xltype & xlbitDLLFree, 0);
            assert_eq!(root.xltype & xlbitXLFree, 0);
            // SAFETY: same live handed-off allocation contract.
            let storage = unsafe { raw_counted_storage(root) };
            assert_eq!(string_pointer_with_flags(root), string_address);
            assert_eq!(usize::from(storage[0]), payload.len());
            assert_eq!(&storage[1..], payload);

            // SAFETY: unique, live handoff pointer; called exactly once.
            unsafe { xl_auto_free12(pointer) };
            assert_no_live_or_handed_off_allocations();
            assert_eq!(handoff_counts().1, callback_frees_before + 1);
        }
    }

    #[test]
    fn handoff_preserves_multi_storage_and_marks_only_the_root() {
        let _guard = lock();
        let returned = mixed_array_plan().materialize().unwrap();
        let root_address = returned.as_xloper() as *const XLOPER12 as *mut XLOPER12;
        let element_address = array_pointer(returned.as_xloper());
        let nested_string_addresses = {
            let elements = array_elements(&returned);
            [string_pointer(&elements[6]), string_pointer(&elements[7])]
        };
        let callback_frees_before = handoff_counts().1;

        let pointer = move_return(returned).into_raw_for_excel();
        assert_eq!(pointer, root_address);
        // SAFETY: the handed-off root remains live until the callback below.
        let root = unsafe { &*pointer };
        assert_eq!(root.xltype & XLTYPE_MASK, xltypeMulti);
        assert_ne!(root.xltype & xlbitDLLFree, 0);
        assert_eq!(root.xltype & xlbitXLFree, 0);
        assert_eq!(array_pointer_with_flags(root), element_address);
        // SAFETY: the handed-off element block remains live with its root.
        let elements = unsafe { raw_array_elements(root) };
        assert_eq!(elements.as_ptr(), element_address);
        assert_eq!(elements.len(), 8);
        assert_eq!(number(&elements[0]), 1.5);
        assert_eq!(integer(&elements[1]), -2);
        assert_eq!(boolean(&elements[2]), 1);
        assert_eq!(error_code(&elements[3]), excel_api_sys::xlerrRef);
        assert_eq!(elements[4].xltype, xltypeMissing);
        assert_eq!(elements[5].xltype, xltypeNil);
        assert_eq!(string_pointer(&elements[6]), nested_string_addresses[0]);
        assert_eq!(string_pointer(&elements[7]), nested_string_addresses[1]);
        // SAFETY: nested counted buffers share the live top-level owner.
        let first_nested_string = unsafe { raw_counted_storage(&elements[6]) };
        // SAFETY: nested counted buffers share the live top-level owner.
        let second_nested_string = unsafe { raw_counted_storage(&elements[7]) };
        assert_eq!(first_nested_string, &[3, 65, 0, 66]);
        assert_eq!(second_nested_string, &[1, 0xD800]);
        for element in elements {
            assert_no_ownership_bits(element);
        }

        // SAFETY: unique, live handoff pointer; called exactly once.
        unsafe { xl_auto_free12(pointer) };
        assert_no_live_or_handed_off_allocations();
        assert_eq!(handoff_counts().1, callback_frees_before + 1);
    }

    #[test]
    fn repeated_handoff_callback_cycles_reclaim_exactly_once() {
        let _guard = lock();
        let callback_frees_before = handoff_counts().1;
        for _ in 0..1_000 {
            handoff_and_reclaim(mixed_array_plan().materialize().unwrap());
        }
        assert_eq!(handoff_counts().1, callback_frees_before + 1_000);
    }

    #[test]
    fn callback_is_null_tolerant_and_abi_exact() {
        let _guard = lock();
        let callback: excel_api_sys::XlAutoFree12Fn = xl_auto_free12;
        let before = handoff_counts();
        // SAFETY: null is explicitly accepted as a defensive no-op.
        unsafe { callback(core::ptr::null_mut()) };
        assert_eq!(handoff_counts(), before);
        assert_no_live_or_handed_off_allocations();
    }

    #[test]
    fn callback_contains_test_panic_and_still_reclaims_valid_handoff() {
        let _guard = lock();
        let callback_frees_before = handoff_counts().1;
        let pointer = materialize(ExcelReturnValue::from("panic-safe")).into_raw_for_excel();
        PANIC_BEFORE_RECLAIM.store(true, core::sync::atomic::Ordering::SeqCst);

        let result = std::panic::catch_unwind(|| {
            // SAFETY: unique, live handoff pointer; called exactly once.
            unsafe { xl_auto_free12(pointer) };
        });
        assert!(result.is_ok(), "panic escaped the extern callback boundary");
        assert_no_live_or_handed_off_allocations();
        assert_eq!(handoff_counts().1, callback_frees_before + 1);
    }

    #[test]
    fn handed_off_allocation_can_be_reclaimed_on_another_thread() {
        let _guard = lock();
        let callback_frees_before = handoff_counts().1;
        let pointer = mixed_array_plan()
            .materialize()
            .unwrap()
            .into_raw_for_excel();
        let address = pointer as usize;

        std::thread::spawn(move || {
            let pointer = address as *mut XLOPER12;
            // SAFETY: ownership of this unique handed-off pointer is moved
            // into this thread, which performs its sole callback reclamation.
            unsafe { xl_auto_free12(pointer) };
        })
        .join()
        .unwrap();

        assert_no_live_or_handed_off_allocations();
        assert_eq!(handoff_counts().1, callback_frees_before + 1);
    }
}
