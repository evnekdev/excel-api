//! Callback-scoped ownership of results created by the Excel C API.
//!
//! Excel owns any auxiliary storage reachable from the root, while Rust owns
//! the boxed root itself.  A required release is attempted exactly once.

use core::{fmt, marker::PhantomData, panic::AssertUnwindSafe};

use excel_api_sys::XLOPER12;

use crate::{
    ConversionError, ConversionLimits, DecodeError, ExcelValue, ExcelValueRef, RawExcelValue,
};

/// Whether an Excel C API result has an outstanding `xlFree` obligation.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum ExcelReleasePolicy {
    /// Excel allocated auxiliary storage and the root must be passed to
    /// `xlFree` once when it is no longer needed.
    XlFreeRequired,
    /// The call result has no auxiliary storage that needs releasing.
    NoReleaseRequired,
}

/// A failure reported while attempting the one permitted `xlFree` call.
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum ExcelReleaseError {
    /// Excel returned its exact (possibly bit-combined) C API failure code.
    ExcelCallFailure { code: i32 },
    /// The callback capability does not permit a call into Excel.
    InvalidContext,
    /// The operation was rejected in a multithreaded-calculation context.
    NotThreadSafe,
    /// The Excel callback entry point is unavailable or has been unlinked.
    BackendUnavailable,
    /// An injected backend panicked. Production backends must never panic.
    BackendPanicked,
}

impl fmt::Display for ExcelReleaseError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::ExcelCallFailure { code } => {
                write!(formatter, "xlFree returned Excel C API code {code}")
            }
            Self::InvalidContext => formatter.write_str("xlFree is unavailable in this context"),
            Self::NotThreadSafe => formatter.write_str("Excel rejected xlFree as not thread safe"),
            Self::BackendUnavailable => {
                formatter.write_str("the Excel callback backend is unavailable")
            }
            Self::BackendPanicked => formatter.write_str("the xlFree backend panicked"),
        }
    }
}

impl std::error::Error for ExcelReleaseError {}

/// The result of a consuming deep copy followed by the mandatory release.
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum ExcelOwnedConversionError {
    Conversion(ConversionError),
    Release(ExcelReleaseError),
    ConversionAndRelease {
        conversion: ConversionError,
        release: ExcelReleaseError,
    },
}

impl fmt::Display for ExcelOwnedConversionError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Conversion(error) => write!(formatter, "Excel result conversion failed: {error}"),
            Self::Release(error) => write!(formatter, "Excel result release failed: {error}"),
            Self::ConversionAndRelease {
                conversion,
                release,
            } => write!(
                formatter,
                "Excel result conversion failed ({conversion}) and release failed ({release})"
            ),
        }
    }
}

impl std::error::Error for ExcelOwnedConversionError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Conversion(error) => Some(error),
            Self::Release(error) => Some(error),
            Self::ConversionAndRelease { conversion, .. } => Some(conversion),
        }
    }
}

/// Narrow callback capability used only to release Excel-created results.
///
/// Implementations must be valid for the entire callback lifetime represented
/// by the borrow and must not panic. A production implementation will call
/// `Excel12v(xlFree, null, 1, [&mut root])` after runtime linking exists.
pub(crate) trait ExcelReleaseBackend {
    fn xl_free(&self, value: *mut XLOPER12) -> Result<(), ExcelReleaseError>;
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum ReleaseState {
    Active,
    Attempted,
}

/// A callback-scoped owner for one result root written by Excel.
///
/// The lifetime prevents the owner from outliving the backend/callback
/// capability. The raw-pointer marker deliberately makes this type neither
/// `Send` nor `Sync`; copy to [`ExcelValue`] before crossing threads.
///
/// ```compile_fail
/// use excel_api::ExcelOwnedValue;
/// fn require_send<T: Send>() {}
/// require_send::<ExcelOwnedValue<'static>>();
/// ```
pub struct ExcelOwnedValue<'call> {
    root: Box<XLOPER12>,
    policy: ExcelReleasePolicy,
    state: ReleaseState,
    backend: &'call dyn ExcelReleaseBackend,
    _not_send_or_sync: PhantomData<*mut ()>,
}

/// A pre-commit XLFree transfer token.
///
/// This token consumes the active owner but intentionally exposes neither a
/// raw root pointer nor a way to set `xlbitXLFree`. Until Prompt 08 commits it
/// at the final return boundary, dropping the token safely falls back to the
/// owner's normal exactly-once release.
///
/// ```compile_fail
/// use excel_api::ExcelXlFreeTransfer;
/// fn require_clone<T: Clone>() {}
/// require_clone::<ExcelXlFreeTransfer<'static>>();
/// ```
#[derive(Debug)]
pub struct ExcelXlFreeTransfer<'call> {
    owner: ExcelOwnedValue<'call>,
}

impl ExcelXlFreeTransfer<'_> {
    /// Borrows the still-uncommitted value without exposing its raw root.
    pub fn as_value_ref(&self) -> Result<ExcelValueRef<'_>, DecodeError> {
        self.owner.as_value_ref()
    }
}

impl fmt::Debug for ExcelOwnedValue<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        let root = self.root.as_ref() as *const XLOPER12;
        formatter
            .debug_struct("ExcelOwnedValue")
            .field("root", &root)
            .field("xltype", &self.root.xltype)
            .field("policy", &self.policy)
            .field("state", &self.state)
            .finish_non_exhaustive()
    }
}

impl<'call> ExcelOwnedValue<'call> {
    /// Constructs an owner from a root initialized by an Excel C API call.
    ///
    /// # Safety
    ///
    /// `root` must be the unique, not-yet-released result of a specified
    /// `Excel12`/`Excel12v` call. Every reachable pointer must remain valid
    /// until `backend` releases the top-level root. `policy` must come from the
    /// call/result ownership metadata, and `backend` must remain a legal Excel
    /// callback capability for `'call`.
    #[allow(dead_code, reason = "the Excel12v call layer is Prompt 08")]
    pub(crate) unsafe fn from_call_result(
        root: XLOPER12,
        policy: ExcelReleasePolicy,
        backend: &'call dyn ExcelReleaseBackend,
    ) -> Self {
        debug_assert_eq!(root.xltype & excel_api_sys::xlbitDLLFree, 0);
        Self {
            root: Box::new(root),
            policy,
            state: ReleaseState::Active,
            backend,
            _not_send_or_sync: PhantomData,
        }
    }

    /// Borrows the active result through the central ownership-bit-masking
    /// decoder.
    pub fn as_value_ref(&self) -> Result<ExcelValueRef<'_>, DecodeError> {
        debug_assert_eq!(self.state, ReleaseState::Active);
        // SAFETY: the constructor contract keeps the boxed root and all
        // reachable Excel storage readable and immutable until release.
        unsafe { RawExcelValue::from_callback(self.root.as_ref()) }.decode()
    }

    /// Borrows the stable result root for a typed, crate-internal follow-up
    /// C API call. It cannot escape the callback owner.
    pub(crate) fn raw_root(&self) -> &XLOPER12 {
        self.root.as_ref()
    }

    /// Deep-copies without consuming or releasing this result.
    pub fn to_owned_value(&self, limits: &ConversionLimits) -> Result<ExcelValue, ConversionError> {
        ExcelValue::from_borrowed_with_limits(self.as_value_ref()?, limits)
    }

    /// Deep-copies and then consumes the Excel release obligation.
    pub fn into_owned_value(
        mut self,
        limits: &ConversionLimits,
    ) -> Result<ExcelValue, ExcelOwnedConversionError> {
        let conversion = self.to_owned_value(limits);
        let release = self.attempt_release();
        match (conversion, release) {
            (Ok(value), Ok(())) => Ok(value),
            (Err(conversion), Ok(())) => Err(ExcelOwnedConversionError::Conversion(conversion)),
            (Ok(_), Err(release)) => Err(ExcelOwnedConversionError::Release(release)),
            (Err(conversion), Err(release)) => {
                Err(ExcelOwnedConversionError::ConversionAndRelease {
                    conversion,
                    release,
                })
            }
        }
    }

    /// Explicitly consumes this owner and attempts any required release once.
    pub fn release(mut self) -> Result<(), ExcelReleaseError> {
        self.attempt_release()
    }

    /// Consumes this owner into a non-duplicable pre-commit transfer token.
    ///
    /// No ownership bit is changed and no raw pointer is exposed here. The
    /// final thunk integration remains deliberately crate-internal work for
    /// Prompt 08.
    pub fn into_xlfree_transfer(self) -> ExcelXlFreeTransfer<'call> {
        ExcelXlFreeTransfer { owner: self }
    }

    fn attempt_release(&mut self) -> Result<(), ExcelReleaseError> {
        if self.state != ReleaseState::Active {
            return Ok(());
        }
        // Commit before the call: a reported failure does not prove that Excel
        // retained the payload, so retrying could double-free it.
        self.state = ReleaseState::Attempted;
        if self.policy == ExcelReleasePolicy::NoReleaseRequired {
            return Ok(());
        }

        let root = self.root.as_mut() as *mut XLOPER12;
        std::panic::catch_unwind(AssertUnwindSafe(|| self.backend.xl_free(root)))
            .unwrap_or(Err(ExcelReleaseError::BackendPanicked))
    }
}

impl Drop for ExcelOwnedValue<'_> {
    fn drop(&mut self) {
        let _ = self.attempt_release();
    }
}

#[cfg(test)]
mod tests {
    use std::{cell::RefCell, panic::catch_unwind, ptr};

    use excel_api_sys::{
        XLOPER12Array, XLOPER12Value, xlbitDLLFree, xlbitXLFree, xltypeErr, xltypeMulti, xltypeNum,
        xltypeStr,
    };

    use super::*;
    use crate::{ExcelError, ExcelString};

    #[derive(Clone, Copy, Debug, Eq, PartialEq)]
    struct ReleaseRecord {
        address: usize,
        xltype: u32,
    }

    #[derive(Default)]
    struct MockBackend {
        records: RefCell<Vec<ReleaseRecord>>,
        failure: RefCell<Option<ExcelReleaseError>>,
        poison_root: bool,
    }

    impl MockBackend {
        fn failing(error: ExcelReleaseError) -> Self {
            Self {
                failure: RefCell::new(Some(error)),
                ..Self::default()
            }
        }

        fn count(&self) -> usize {
            self.records.borrow().len()
        }
    }

    impl ExcelReleaseBackend for MockBackend {
        fn xl_free(&self, value: *mut XLOPER12) -> Result<(), ExcelReleaseError> {
            // SAFETY: the owner supplies its live boxed root for this call.
            let xltype = unsafe { (*value).xltype };
            self.records.borrow_mut().push(ReleaseRecord {
                address: value as usize,
                xltype,
            });
            let result = self.failure.borrow_mut().take().map_or(Ok(()), Err);
            if self.poison_root {
                // SAFETY: xlFree is documented as the sole C API operation
                // that mutates its argument, and no read follows an attempt.
                unsafe {
                    (*value).val = XLOPER12Value { w: 0 };
                    (*value).xltype = excel_api_sys::xltypeNil;
                }
            }
            result
        }
    }

    fn raw(value: XLOPER12Value, xltype: u32) -> XLOPER12 {
        XLOPER12 { val: value, xltype }
    }

    unsafe fn owner<'a>(
        value: XLOPER12,
        policy: ExcelReleasePolicy,
        backend: &'a MockBackend,
    ) -> ExcelOwnedValue<'a> {
        // SAFETY: each test keeps every fixture allocation alive through the
        // owner and uses a backend that never dereferences nested storage.
        unsafe { ExcelOwnedValue::from_call_result(value, policy, backend) }
    }

    #[test]
    fn explicit_drop_and_no_release_policies_are_exact() {
        let backend = MockBackend::default();
        // SAFETY: this scalar fixture is initialized and the backend remains live.
        let value = unsafe {
            owner(
                raw(XLOPER12Value { num: 1.5 }, xltypeNum),
                ExcelReleasePolicy::XlFreeRequired,
                &backend,
            )
        };
        value.release().unwrap();
        assert_eq!(backend.count(), 1);

        {
            // SAFETY: this scalar fixture is initialized and the backend remains live.
            let _fallback = unsafe {
                owner(
                    raw(XLOPER12Value { num: 2.5 }, xltypeNum),
                    ExcelReleasePolicy::XlFreeRequired,
                    &backend,
                )
            };
        }
        assert_eq!(backend.count(), 2);

        // SAFETY: this scalar fixture is initialized and has no nested payload.
        let no_release = unsafe {
            owner(
                raw(XLOPER12Value { num: 3.5 }, xltypeNum),
                ExcelReleasePolicy::NoReleaseRequired,
                &backend,
            )
        };
        no_release.release().unwrap();
        assert_eq!(backend.count(), 2);
    }

    #[test]
    fn ownership_bits_are_masked_and_utf16_copy_is_lossless() {
        let backend = MockBackend::default();
        let mut counted = vec![4, b'A' as u16, 0, 0xD800, 0xDC00];
        // SAFETY: the counted buffer remains initialized through owner release.
        let value = unsafe {
            owner(
                raw(
                    XLOPER12Value {
                        str: counted.as_mut_ptr(),
                    },
                    xltypeStr | xlbitXLFree,
                ),
                ExcelReleasePolicy::XlFreeRequired,
                &backend,
            )
        };
        assert_eq!(value.as_value_ref().unwrap().kind_name(), "text");
        let copied = value
            .into_owned_value(&ConversionLimits::default())
            .unwrap();
        assert_eq!(
            copied,
            ExcelValue::Text(ExcelString::from_utf16_units(counted[1..].to_vec()))
        );
        counted[1..].fill(b'Z' as u16);
        let ExcelValue::Text(copied) = copied else {
            panic!("expected text")
        };
        assert_eq!(copied.as_utf16(), &[b'A' as u16, 0, 0xD800, 0xDC00]);
        assert_eq!(backend.count(), 1);
    }

    #[test]
    fn multi_is_copied_and_only_the_top_level_root_is_released() {
        let backend = MockBackend::default();
        let mut first = [2, b'a' as u16, 0];
        let mut second = [2, 0xD800, b'b' as u16];
        let mut elements = vec![
            raw(XLOPER12Value { num: 4.0 }, xltypeNum),
            raw(
                XLOPER12Value {
                    err: excel_api_sys::xlerrNA,
                },
                xltypeErr,
            ),
            raw(XLOPER12Value { w: 0 }, excel_api_sys::xltypeNil),
            raw(
                XLOPER12Value {
                    str: first.as_mut_ptr(),
                },
                xltypeStr,
            ),
            raw(
                XLOPER12Value {
                    str: second.as_mut_ptr(),
                },
                xltypeStr,
            ),
            raw(XLOPER12Value { num: -1.0 }, xltypeNum),
        ];
        let root = raw(
            XLOPER12Value {
                array: XLOPER12Array {
                    lparray: elements.as_mut_ptr(),
                    rows: 2,
                    columns: 3,
                },
            },
            xltypeMulti,
        );
        // SAFETY: the element and string buffers remain live through release.
        let value = unsafe { owner(root, ExcelReleasePolicy::XlFreeRequired, &backend) };
        let copied = value
            .into_owned_value(&ConversionLimits::default())
            .unwrap();
        let ExcelValue::Array(copied) = copied else {
            panic!("expected array")
        };
        assert_eq!((copied.rows(), copied.columns()), (2, 3));
        assert_eq!(copied.get(0, 1), Some(&ExcelValue::Error(ExcelError::Na)));
        assert_eq!(backend.count(), 1);
        let record = backend.records.borrow()[0];
        assert_eq!(record.xltype, xltypeMulti);
        assert_ne!(record.address, elements.as_ptr() as usize);
    }

    #[test]
    fn conversion_and_release_failures_are_composed_without_retry() {
        let release = ExcelReleaseError::InvalidContext;
        let backend = MockBackend::failing(release.clone());
        let malformed = raw(
            XLOPER12Value {
                str: ptr::null_mut(),
            },
            xltypeStr,
        );
        // SAFETY: malformed nested storage is intentional; the initialized
        // root is rejected before the null pointer can be dereferenced.
        let error = unsafe { owner(malformed, ExcelReleasePolicy::XlFreeRequired, &backend) }
            .into_owned_value(&ConversionLimits::default())
            .unwrap_err();
        assert_eq!(
            error,
            ExcelOwnedConversionError::ConversionAndRelease {
                conversion: ConversionError::BorrowedValueDecode(DecodeError::NullStringPointer),
                release,
            }
        );
        assert_eq!(backend.count(), 1);
    }

    #[test]
    fn conversion_limits_and_references_fail_but_still_release() {
        let backend = MockBackend::default();
        let mut counted = [2, b'a' as u16, b'b' as u16];
        // SAFETY: the counted buffer remains initialized through release.
        let string = unsafe {
            owner(
                raw(
                    XLOPER12Value {
                        str: counted.as_mut_ptr(),
                    },
                    xltypeStr,
                ),
                ExcelReleasePolicy::XlFreeRequired,
                &backend,
            )
        };
        let limits = ConversionLimits {
            max_string_code_units: 1,
            ..ConversionLimits::default()
        };
        assert_eq!(
            string.into_owned_value(&limits),
            Err(ExcelOwnedConversionError::Conversion(
                ConversionError::StringLimitExceeded {
                    actual: 2,
                    maximum: 1,
                }
            ))
        );

        // SAFETY: the inline reference is initialized and contains no pointer.
        let reference = unsafe {
            owner(
                raw(
                    XLOPER12Value {
                        sref: excel_api_sys::XLOPER12SRef {
                            count: 1,
                            reference: excel_api_sys::XLREF12 {
                                rwFirst: 0,
                                rwLast: 0,
                                colFirst: 0,
                                colLast: 0,
                            },
                        },
                    },
                    excel_api_sys::xltypeSRef,
                ),
                ExcelReleasePolicy::XlFreeRequired,
                &backend,
            )
        };
        assert_eq!(
            reference.into_owned_value(&ConversionLimits::default()),
            Err(ExcelOwnedConversionError::Conversion(
                ConversionError::UnsupportedReference
            ))
        );
        assert_eq!(backend.count(), 2);
    }

    #[test]
    fn successful_conversion_preserves_release_failure_and_does_not_retry() {
        for failure in [
            ExcelReleaseError::ExcelCallFailure {
                code: excel_api_sys::xlretAbort | excel_api_sys::xlretUncalced,
            },
            ExcelReleaseError::NotThreadSafe,
            ExcelReleaseError::BackendUnavailable,
        ] {
            let backend = MockBackend::failing(failure.clone());
            // SAFETY: this scalar fixture is initialized and the backend remains live.
            let value = unsafe {
                owner(
                    raw(XLOPER12Value { num: 9.0 }, xltypeNum),
                    ExcelReleasePolicy::XlFreeRequired,
                    &backend,
                )
            };
            assert_eq!(
                value.into_owned_value(&ConversionLimits::default()),
                Err(ExcelOwnedConversionError::Release(failure))
            );
            assert_eq!(backend.count(), 1);
        }
    }

    #[test]
    fn panic_during_conversion_still_releases_once() {
        let backend = MockBackend::default();
        // SAFETY: this scalar fixture is initialized and the backend remains live.
        let value = unsafe {
            owner(
                raw(XLOPER12Value { num: 1.0 }, xltypeNum),
                ExcelReleasePolicy::XlFreeRequired,
                &backend,
            )
        };
        let result = catch_unwind(AssertUnwindSafe(move || {
            let _view = value.as_value_ref().unwrap();
            panic!("injected conversion panic");
        }));
        assert!(result.is_err());
        assert_eq!(backend.count(), 1);
    }

    #[test]
    fn backend_panic_is_contained_and_consumes_the_obligation() {
        struct PanicBackend(RefCell<usize>);
        impl ExcelReleaseBackend for PanicBackend {
            fn xl_free(&self, _: *mut XLOPER12) -> Result<(), ExcelReleaseError> {
                *self.0.borrow_mut() += 1;
                panic!("backend panic")
            }
        }
        let backend = PanicBackend(RefCell::new(0));
        // SAFETY: this scalar fixture is initialized and the backend remains live.
        let value = unsafe {
            ExcelOwnedValue::from_call_result(
                raw(XLOPER12Value { num: 1.0 }, xltypeNum),
                ExcelReleasePolicy::XlFreeRequired,
                &backend,
            )
        };
        assert_eq!(value.release(), Err(ExcelReleaseError::BackendPanicked));
        assert_eq!(*backend.0.borrow(), 1);
    }

    #[test]
    fn repeated_release_cycles_are_exactly_once_and_never_set_dllfree() {
        let backend = MockBackend::default();
        for index in 0..1_000 {
            // SAFETY: every scalar fixture is initialized and the backend remains live.
            let value = unsafe {
                owner(
                    raw(XLOPER12Value { num: index as f64 }, xltypeNum),
                    ExcelReleasePolicy::XlFreeRequired,
                    &backend,
                )
            };
            value.release().unwrap();
        }
        assert_eq!(backend.count(), 1_000);
        assert!(
            backend
                .records
                .borrow()
                .iter()
                .all(|record| record.xltype & xlbitDLLFree == 0)
        );
    }

    #[test]
    fn transfer_token_consumes_owner_without_premature_bit_or_release() {
        let backend = MockBackend::default();
        // SAFETY: this scalar fixture is initialized and the backend remains live.
        let value = unsafe {
            owner(
                raw(XLOPER12Value { num: 7.0 }, xltypeNum),
                ExcelReleasePolicy::XlFreeRequired,
                &backend,
            )
        };
        let token = value.into_xlfree_transfer();
        assert_eq!(backend.count(), 0);
        assert_eq!(token.as_value_ref().unwrap().kind_name(), "number");
        assert_eq!(token.owner.root.xltype & (xlbitXLFree | xlbitDLLFree), 0);
        drop(token);
        assert_eq!(backend.count(), 1);
    }
}
