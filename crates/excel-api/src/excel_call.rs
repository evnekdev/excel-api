//! Narrow Excel12v callback capability and the M8 call catalogue.

use core::{
    ffi::{c_char, c_int, c_void},
    fmt, ptr,
};
use std::sync::atomic::{AtomicUsize, Ordering};

use excel_api_sys::{LPXLOPER12, XLOPER12, XLOPER12Value};

use crate::excel_owned::ExcelReleaseBackend;
use crate::{
    CommandRegistration, ExcelOwnedValue, ExcelReleaseError, ExcelReleasePolicy, ExcelString,
    FunctionRegistration, LifecycleContext, RegistrationError,
};

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum CallPermission {
    Lifecycle,
    Worksheet,
    ThreadSafe,
    Macro,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum ResultRoot {
    None,
    Required,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct ExcelCallDescriptor {
    pub name: &'static str,
    pub function: i32,
    /// Callback capabilities from which this call is documented as legal.
    pub permissions: &'static [CallPermission],
    pub result: ResultRoot,
    pub release: ExcelReleasePolicy,
    pub thread_safe: bool,
    pub minimum_arguments: usize,
    pub maximum_arguments: usize,
}

pub const XL_GET_NAME: ExcelCallDescriptor = ExcelCallDescriptor {
    name: "xlGetName",
    function: excel_api_sys::xlGetName,
    permissions: &[CallPermission::Lifecycle],
    result: ResultRoot::Required,
    release: ExcelReleasePolicy::XlFreeRequired,
    thread_safe: false,
    minimum_arguments: 0,
    maximum_arguments: 0,
};
pub const XLF_REGISTER: ExcelCallDescriptor = ExcelCallDescriptor {
    name: "xlfRegister",
    function: excel_api_sys::xlfRegister,
    permissions: &[CallPermission::Lifecycle],
    result: ResultRoot::Required,
    release: ExcelReleasePolicy::XlFreeRequired,
    thread_safe: false,
    minimum_arguments: 10,
    maximum_arguments: 255,
};
pub const XLF_UNREGISTER: ExcelCallDescriptor = ExcelCallDescriptor {
    name: "xlfUnregister",
    function: excel_api_sys::xlfUnregister,
    permissions: &[CallPermission::Lifecycle],
    result: ResultRoot::Required,
    release: ExcelReleasePolicy::XlFreeRequired,
    thread_safe: false,
    minimum_arguments: 1,
    maximum_arguments: 1,
};
pub const XLF_SET_NAME: ExcelCallDescriptor = ExcelCallDescriptor {
    name: "xlfSetName",
    function: excel_api_sys::xlfSetName,
    permissions: &[CallPermission::Lifecycle],
    result: ResultRoot::Required,
    release: ExcelReleasePolicy::XlFreeRequired,
    thread_safe: false,
    minimum_arguments: 1,
    maximum_arguments: 2,
};
pub const XL_FREE: ExcelCallDescriptor = ExcelCallDescriptor {
    name: "xlFree",
    function: excel_api_sys::xlFree,
    permissions: &[CallPermission::ThreadSafe],
    result: ResultRoot::None,
    release: ExcelReleasePolicy::NoReleaseRequired,
    thread_safe: true,
    minimum_arguments: 1,
    maximum_arguments: 255,
};
/// Polls the Excel cancellation/break request. This is not a calculation-state query.
pub const XL_ABORT: ExcelCallDescriptor = ExcelCallDescriptor {
    name: "xlAbort",
    function: excel_api_sys::xlAbort,
    permissions: &[
        CallPermission::Worksheet,
        CallPermission::ThreadSafe,
        CallPermission::Macro,
    ],
    result: ResultRoot::Required,
    release: ExcelReleasePolicy::NoReleaseRequired,
    thread_safe: true,
    minimum_arguments: 0,
    maximum_arguments: 1,
};
/// Resolves a sheet name (or the active/front sheet when omitted) to its internal ID.
pub const XL_SHEET_ID: ExcelCallDescriptor = ExcelCallDescriptor {
    name: "xlSheetId",
    function: excel_api_sys::xlSheetId,
    permissions: &[
        CallPermission::Worksheet,
        CallPermission::ThreadSafe,
        CallPermission::Macro,
    ],
    result: ResultRoot::Required,
    release: ExcelReleasePolicy::NoReleaseRequired,
    thread_safe: true,
    minimum_arguments: 0,
    maximum_arguments: 1,
};
/// Resolves an internal/external reference to its sheet name.
pub const XL_SHEET_NM: ExcelCallDescriptor = ExcelCallDescriptor {
    name: "xlSheetNm",
    function: excel_api_sys::xlSheetNm,
    permissions: &[
        CallPermission::Worksheet,
        CallPermission::ThreadSafe,
        CallPermission::Macro,
    ],
    result: ResultRoot::Required,
    release: ExcelReleasePolicy::XlFreeRequired,
    thread_safe: true,
    minimum_arguments: 1,
    maximum_arguments: 1,
};
/// Returns information about the caller of a worksheet function or macro.
pub const XLF_CALLER: ExcelCallDescriptor = ExcelCallDescriptor {
    name: "xlfCaller",
    function: excel_api_sys::xlfCaller,
    permissions: &[CallPermission::Worksheet, CallPermission::Macro],
    result: ResultRoot::Required,
    release: ExcelReleasePolicy::XlFreeRequired,
    thread_safe: false,
    minimum_arguments: 0,
    maximum_arguments: 0,
};
/// Converts a C API value to an explicitly requested Excel root type.
pub const XL_COERCE: ExcelCallDescriptor = ExcelCallDescriptor {
    name: "xlCoerce",
    function: excel_api_sys::xlCoerce,
    permissions: &[
        CallPermission::Worksheet,
        CallPermission::ThreadSafe,
        CallPermission::Macro,
    ],
    result: ResultRoot::Required,
    release: ExcelReleasePolicy::XlFreeRequired,
    thread_safe: true,
    minimum_arguments: 1,
    maximum_arguments: 2,
};

/// Controls whether an `xlAbort` poll retains or clears a pending break.
///
/// `PreservePendingBreak` is represented by an omitted argument in
/// [`WorksheetContext::is_cancellation_requested`], or by an explicit TRUE in
/// [`WorksheetContext::is_cancellation_requested_with`]. `ClearPendingBreak`
/// is explicit FALSE. The returned Boolean reports a user break request only;
/// it never reports Excel's general calculation progress or state.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum AbortCheckMode {
    PreservePendingBreak,
    ClearPendingBreak,
}

/// The Excel root type requested from [`XL_COERCE`].
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum CoerceTarget {
    Number,
    Text,
    Boolean,
    Error,
}

impl CoerceTarget {
    const fn xltype(self) -> i32 {
        match self {
            Self::Number => excel_api_sys::xltypeNum as i32,
            Self::Text => excel_api_sys::xltypeStr as i32,
            Self::Boolean => excel_api_sys::xltypeBool as i32,
            Self::Error => excel_api_sys::xltypeErr as i32,
        }
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct ExcelReturnCode(pub i32);

impl ExcelReturnCode {
    pub const fn is_success(self) -> bool {
        self.0 == excel_api_sys::xlretSuccess
    }
    pub const fn has(self, flag: i32) -> bool {
        self.0 & flag != 0
    }
}

impl fmt::Display for ExcelReturnCode {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}", self.0)
    }
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub enum ExcelCallError {
    BackendUnavailable,
    InvalidArgumentCount {
        function: &'static str,
        count: usize,
    },
    IllegalContext {
        function: &'static str,
        context: CallPermission,
    },
    ExcelFailure {
        function: &'static str,
        code: ExcelReturnCode,
    },
    MalformedResult {
        function: &'static str,
        expected: &'static str,
        actual: &'static str,
    },
    ResultConversion(String),
    Registration(RegistrationError),
}

impl fmt::Display for ExcelCallError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::BackendUnavailable => f.write_str("Excel12v callback backend is unavailable"),
            Self::InvalidArgumentCount { function, count } => {
                write!(f, "{function} does not accept {count} arguments")
            }
            Self::IllegalContext { function, context } => {
                write!(
                    f,
                    "{function} is not legal in the {context:?} callback context"
                )
            }
            Self::ExcelFailure { function, code } => {
                write!(f, "{function} returned Excel C API code {code}")
            }
            Self::MalformedResult {
                function,
                expected,
                actual,
            } => write!(f, "{function} returned {actual}, expected {expected}"),
            Self::ResultConversion(message) => {
                write!(f, "Excel result conversion failed: {message}")
            }
            Self::Registration(error) => write!(f, "registration descriptor is invalid: {error}"),
        }
    }
}
impl std::error::Error for ExcelCallError {}

pub trait ExcelCallBackend: Send + Sync {
    fn link(&self) -> Result<(), ExcelCallError>;
    fn unlink(&self);
    fn is_linked(&self) -> bool;

    /// # Safety
    /// All argument pointers must remain valid and uniquely mutable where Excel permits mutation;
    /// `result` must be null only for calls documented without a result. The caller must be in an
    /// active Excel callback on a legal thread, and this backend must remain linked for the call.
    unsafe fn excel12v_raw(
        &self,
        function: i32,
        result: *mut XLOPER12,
        count: c_int,
        arguments: *mut LPXLOPER12,
    ) -> i32;
}

#[derive(Default)]
pub struct SdkExcel12vBackend {
    entry: AtomicUsize,
}

impl fmt::Debug for SdkExcel12vBackend {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("SdkExcel12vBackend")
            .field("linked", &self.is_linked())
            .finish()
    }
}

impl SdkExcel12vBackend {
    pub const fn new() -> Self {
        Self {
            entry: AtomicUsize::new(0),
        }
    }

    pub fn set_entry_point(&self, entry: excel_api_sys::Excel12EntryPtFn) {
        if self.entry.load(Ordering::Acquire) == 0 {
            self.entry.store(entry as usize, Ordering::Release);
        }
    }

    #[cfg(windows)]
    fn resolve_host_entry() -> Option<usize> {
        #[link(name = "kernel32")]
        unsafe extern "system" {
            fn GetModuleHandleW(name: *const u16) -> *mut c_void;
            fn GetProcAddress(module: *mut c_void, name: *const c_char) -> *mut c_void;
        }
        // SAFETY: null requests the current process executable; the symbol name is NUL-terminated.
        let module = unsafe { GetModuleHandleW(ptr::null()) };
        if module.is_null() {
            return None;
        }
        // SAFETY: `module` is the live host module returned above.
        let address = unsafe { GetProcAddress(module, c"MdCallBack12".as_ptr()) };
        (!address.is_null()).then_some(address as usize)
    }

    #[cfg(not(windows))]
    fn resolve_host_entry() -> Option<usize> {
        None
    }
}

impl ExcelCallBackend for SdkExcel12vBackend {
    fn link(&self) -> Result<(), ExcelCallError> {
        if self.is_linked() {
            return Ok(());
        }
        let entry = Self::resolve_host_entry().ok_or(ExcelCallError::BackendUnavailable)?;
        let _ = self
            .entry
            .compare_exchange(0, entry, Ordering::AcqRel, Ordering::Acquire);
        Ok(())
    }

    fn unlink(&self) {
        self.entry.store(0, Ordering::Release);
    }
    fn is_linked(&self) -> bool {
        self.entry.load(Ordering::Acquire) != 0
    }

    unsafe fn excel12v_raw(
        &self,
        function: i32,
        result: *mut XLOPER12,
        count: c_int,
        arguments: *mut LPXLOPER12,
    ) -> i32 {
        let address = self.entry.load(Ordering::Acquire);
        if address == 0 {
            return excel_api_sys::xlretFailed;
        }
        // SAFETY: a nonzero address is installed only from MdCallBack12 or SetExcel12EntryPt.
        let callback: excel_api_sys::Excel12EntryPtFn = unsafe { core::mem::transmute(address) };
        // SAFETY: forwarded from this method's contract while the atomic keeps the entry installed.
        unsafe { callback(function, count, arguments, result) }
    }
}

pub(crate) struct CallCapability<'call> {
    backend: &'call dyn ExcelCallBackend,
}

impl fmt::Debug for CallCapability<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("CallCapability")
            .field("linked", &self.backend.is_linked())
            .finish()
    }
}

impl<'call> CallCapability<'call> {
    pub(crate) const fn new(backend: &'call dyn ExcelCallBackend) -> Self {
        Self { backend }
    }

    fn call(
        &'call self,
        permission: CallPermission,
        descriptor: ExcelCallDescriptor,
        arguments: &mut [LPXLOPER12],
    ) -> Result<Option<ExcelOwnedValue<'call>>, ExcelCallError> {
        if !descriptor.permissions.contains(&permission) {
            return Err(ExcelCallError::IllegalContext {
                function: descriptor.name,
                context: permission,
            });
        }
        if arguments.len() < descriptor.minimum_arguments
            || arguments.len() > descriptor.maximum_arguments
        {
            return Err(ExcelCallError::InvalidArgumentCount {
                function: descriptor.name,
                count: arguments.len(),
            });
        }
        let count =
            c_int::try_from(arguments.len()).map_err(|_| ExcelCallError::InvalidArgumentCount {
                function: descriptor.name,
                count: arguments.len(),
            })?;
        if !self.backend.is_linked() {
            return Err(ExcelCallError::BackendUnavailable);
        }
        let mut root: XLOPER12 = XLOPER12 {
            val: XLOPER12Value { w: 0 },
            xltype: excel_api_sys::xltypeNil,
        };
        let result = if descriptor.result == ResultRoot::Required {
            &mut root
        } else {
            ptr::null_mut()
        };
        // SAFETY: roots and the mutable pointer vector remain stable through the synchronous call.
        let raw = unsafe {
            self.backend
                .excel12v_raw(descriptor.function, result, count, arguments.as_mut_ptr())
        };
        let code = ExcelReturnCode(raw);
        if !code.is_success() {
            crate::diagnostics::emit(crate::DiagnosticEvent::new(
                crate::DiagnosticCode::ExcelCall,
                crate::DiagnosticSeverity::Error,
                code.0,
            ));
            return Err(ExcelCallError::ExcelFailure {
                function: descriptor.name,
                code,
            });
        }
        if descriptor.result == ResultRoot::None {
            return Ok(None);
        }
        // SAFETY: successful Excel12v initialized this root; descriptor metadata supplies release policy.
        Ok(Some(unsafe {
            ExcelOwnedValue::from_call_result(root, descriptor.release, self)
        }))
    }

    fn cancellation_requested(
        &'call self,
        permission: CallPermission,
        mode: Option<AbortCheckMode>,
    ) -> Result<bool, ExcelCallError> {
        if !XL_ABORT.permissions.contains(&permission) {
            return Err(ExcelCallError::IllegalContext {
                function: XL_ABORT.name,
                context: permission,
            });
        }
        let mut retain = XLOPER12 {
            val: XLOPER12Value { xbool: 1 },
            xltype: excel_api_sys::xltypeBool,
        };
        if matches!(mode, Some(AbortCheckMode::ClearPendingBreak)) {
            retain.val = XLOPER12Value { xbool: 0 };
        }
        let mut arguments = mode
            .map(|_| vec![&mut retain as *mut XLOPER12])
            .unwrap_or_default();
        if !self.backend.is_linked() {
            return Err(ExcelCallError::BackendUnavailable);
        }
        let mut root = XLOPER12 {
            val: XLOPER12Value { w: 0 },
            xltype: excel_api_sys::xltypeNil,
        };
        // SAFETY: `root` and the optional Boolean argument remain initialized,
        // uniquely mutable, and live for this synchronous documented call.
        let raw = unsafe {
            self.backend.excel12v_raw(
                XL_ABORT.function,
                &mut root,
                arguments.len() as c_int,
                arguments.as_mut_ptr(),
            )
        };
        let code = ExcelReturnCode(raw);
        if !code.is_success() {
            crate::diagnostics::emit(crate::DiagnosticEvent::new(
                crate::DiagnosticCode::ExcelCall,
                crate::DiagnosticSeverity::Error,
                code.0,
            ));
            return Err(ExcelCallError::ExcelFailure {
                function: XL_ABORT.name,
                code,
            });
        }
        if root.xltype != excel_api_sys::xltypeBool {
            return Err(ExcelCallError::MalformedResult {
                function: XL_ABORT.name,
                expected: "Boolean cancellation flag",
                actual: "non-Boolean result",
            });
        }
        // `xlAbort` returns an immediate Boolean; its descriptor records that
        // no Excel-owned auxiliary allocation is expected, so no owner/xlFree
        // obligation is created.
        // SAFETY: the successful call's descriptor requires an immediate
        // `xltypeBool` result, which initializes this union member.
        Ok(unsafe { root.val.xbool != 0 })
    }
}

impl ExcelReleaseBackend for CallCapability<'_> {
    fn xl_free(&self, value: *mut XLOPER12) -> Result<(), ExcelReleaseError> {
        if !self.backend.is_linked() {
            return Err(ExcelReleaseError::BackendUnavailable);
        }
        let mut arguments = [value];
        // SAFETY: ExcelOwnedValue provides its unique live top-level C-API result root; xlFree has no result.
        let code = unsafe {
            self.backend.excel12v_raw(
                excel_api_sys::xlFree,
                ptr::null_mut(),
                1,
                arguments.as_mut_ptr(),
            )
        };
        if code == excel_api_sys::xlretSuccess {
            Ok(())
        } else if code & excel_api_sys::xlretNotThreadSafe != 0 {
            Err(ExcelReleaseError::NotThreadSafe)
        } else {
            Err(ExcelReleaseError::ExcelCallFailure { code })
        }
    }
}

fn counted(text: &str) -> Result<Box<[u16]>, RegistrationError> {
    let units: Vec<u16> = text.encode_utf16().collect();
    if units.len() > excel_api_sys::EXCEL12_MAX_STRING_CODE_UNITS {
        return Err(RegistrationError::StringTooLong);
    }
    let mut storage = Vec::with_capacity(units.len() + 1);
    storage.push(units.len() as u16);
    storage.extend(units);
    Ok(storage.into_boxed_slice())
}

struct CallArguments {
    strings: Vec<Box<[u16]>>,
    roots: Vec<XLOPER12>,
}

impl CallArguments {
    fn registration(
        module: &ExcelString,
        registration: &FunctionRegistration,
    ) -> Result<Self, ExcelCallError> {
        registration
            .validate()
            .map_err(ExcelCallError::Registration)?;
        let module_text = String::from_utf16(module.as_utf16()).map_err(|_| {
            ExcelCallError::ResultConversion("module name is not valid UTF-16".into())
        })?;
        let type_text = registration
            .type_text()
            .map_err(ExcelCallError::Registration)?;
        let argument_text = registration.argument_names.join(",");
        let mut values = vec![
            module_text,
            registration.rust_symbol.into(),
            type_text,
            registration.excel_name.into(),
            argument_text,
            registration.category.unwrap_or("Rust").into(),
            String::new(),
            String::new(),
            registration.description.unwrap_or("").into(),
        ];
        values.extend(
            registration
                .argument_descriptions
                .iter()
                .map(|value| (*value).to_owned()),
        );
        let mut strings = Vec::with_capacity(values.len());
        for value in values {
            strings.push(counted(&value).map_err(ExcelCallError::Registration)?);
        }
        let mut roots: Vec<XLOPER12> = strings
            .iter_mut()
            .map(|value| XLOPER12 {
                val: XLOPER12Value {
                    str: value.as_mut_ptr(),
                },
                xltype: excel_api_sys::xltypeStr,
            })
            .collect();
        roots.insert(
            5,
            XLOPER12 {
                val: XLOPER12Value { num: 1.0 },
                xltype: excel_api_sys::xltypeNum,
            },
        );
        Ok(Self { strings, roots })
    }

    fn command_registration(
        module: &ExcelString,
        registration: &CommandRegistration,
    ) -> Result<Self, ExcelCallError> {
        registration
            .validate()
            .map_err(ExcelCallError::Registration)?;
        let module_text = String::from_utf16(module.as_utf16()).map_err(|_| {
            ExcelCallError::ResultConversion("module name is not valid UTF-16".into())
        })?;
        let values = vec![
            module_text,
            registration.rust_symbol.into(),
            registration.type_text().into(),
            registration.excel_name.into(),
            String::new(),
            registration.shortcut.unwrap_or("").into(),
            String::new(),
            registration.description.unwrap_or("").into(),
        ];
        let mut strings = Vec::with_capacity(values.len());
        for value in values {
            strings.push(counted(&value).map_err(ExcelCallError::Registration)?);
        }
        let mut roots: Vec<XLOPER12> = strings
            .iter_mut()
            .map(|value| XLOPER12 {
                val: XLOPER12Value {
                    str: value.as_mut_ptr(),
                },
                xltype: excel_api_sys::xltypeStr,
            })
            .collect();
        // `pxMacroType = 2` is the documented command registration form.
        roots.insert(
            5,
            XLOPER12 {
                val: XLOPER12Value { num: 2.0 },
                xltype: excel_api_sys::xltypeNum,
            },
        );
        Ok(Self { strings, roots })
    }

    fn pointers(&mut self) -> Vec<LPXLOPER12> {
        let _keep_alive = &self.strings;
        self.roots
            .iter_mut()
            .map(|root| root as *mut XLOPER12)
            .collect()
    }

    fn one_string(value: &str) -> Result<Self, ExcelCallError> {
        let mut strings = vec![counted(value).map_err(ExcelCallError::Registration)?];
        let roots = vec![XLOPER12 {
            val: XLOPER12Value {
                str: strings[0].as_mut_ptr(),
            },
            xltype: excel_api_sys::xltypeStr,
        }];
        Ok(Self { strings, roots })
    }
}

impl LifecycleContext<'_> {
    pub(crate) fn get_module_name(&self) -> Result<ExcelString, ExcelCallError> {
        let mut arguments = [];
        let owner = self
            .capability()
            .call(CallPermission::Lifecycle, XL_GET_NAME, &mut arguments)?
            .expect("descriptor requires a result");
        let value = owner
            .into_owned_value(&crate::ConversionLimits::default())
            .map_err(|error| ExcelCallError::ResultConversion(error.to_string()))?;
        match value {
            crate::ExcelValue::Text(text) => Ok(text),
            other => Err(ExcelCallError::MalformedResult {
                function: XL_GET_NAME.name,
                expected: "text",
                actual: other.kind_name(),
            }),
        }
    }

    pub(crate) fn register(
        &self,
        module: &ExcelString,
        registration: &FunctionRegistration,
    ) -> Result<f64, ExcelCallError> {
        let mut storage = CallArguments::registration(module, registration)?;
        let mut arguments = storage.pointers();
        let owner = self
            .capability()
            .call(CallPermission::Lifecycle, XLF_REGISTER, &mut arguments)?
            .expect("descriptor requires a result");
        let value = owner
            .into_owned_value(&crate::ConversionLimits::default())
            .map_err(|error| ExcelCallError::ResultConversion(error.to_string()))?;
        match value {
            crate::ExcelValue::Number(id) if id.is_finite() => Ok(id),
            other => Err(ExcelCallError::MalformedResult {
                function: XLF_REGISTER.name,
                expected: "finite registration ID",
                actual: other.kind_name(),
            }),
        }
    }

    pub(crate) fn register_command(
        &self,
        module: &ExcelString,
        registration: &CommandRegistration,
    ) -> Result<f64, ExcelCallError> {
        let mut storage = CallArguments::command_registration(module, registration)?;
        let mut arguments = storage.pointers();
        let owner = self
            .capability()
            .call(CallPermission::Lifecycle, XLF_REGISTER, &mut arguments)?
            .expect("descriptor requires a result");
        let value = owner
            .into_owned_value(&crate::ConversionLimits::default())
            .map_err(|error| ExcelCallError::ResultConversion(error.to_string()))?;
        match value {
            crate::ExcelValue::Number(id) if id.is_finite() => Ok(id),
            other => Err(ExcelCallError::MalformedResult {
                function: XLF_REGISTER.name,
                expected: "finite command registration ID",
                actual: other.kind_name(),
            }),
        }
    }

    pub(crate) fn unregister(&self, registration_id: f64) -> Result<(), ExcelCallError> {
        let mut id = XLOPER12 {
            val: XLOPER12Value {
                num: registration_id,
            },
            xltype: excel_api_sys::xltypeNum,
        };
        let mut arguments = [&mut id as *mut XLOPER12];
        let owner = self
            .capability()
            .call(CallPermission::Lifecycle, XLF_UNREGISTER, &mut arguments)?
            .expect("descriptor requires a result");
        owner
            .release()
            .map_err(|error| ExcelCallError::ResultConversion(error.to_string()))
    }

    pub(crate) fn delete_defined_name(&self, name: &str) -> Result<(), ExcelCallError> {
        let mut storage = CallArguments::one_string(name)?;
        let mut arguments = storage.pointers();
        let owner = self
            .capability()
            .call(CallPermission::Lifecycle, XLF_SET_NAME, &mut arguments)?
            .expect("descriptor requires a result");
        owner
            .release()
            .map_err(|error| ExcelCallError::ResultConversion(error.to_string()))
    }
}

macro_rules! sheet_context {
    ($type:ty, $permission:expr) => {
        impl<'call> $type {
            /// Returns the internal ID of Excel's active/front sheet.
            ///
            /// This is deliberately not named "current": the zero-argument
            /// `xlSheetId` contract identifies the active/front sheet.
            pub fn active_sheet_id(&self) -> Result<u64, ExcelCallError> {
                let mut arguments = [];
                let owner = self
                    .capability()
                    .call($permission, XL_SHEET_ID, &mut arguments)?
                    .expect("descriptor requires a result");
                let root = owner.raw_root();
                if root.xltype != excel_api_sys::xltypeRef {
                    return Err(ExcelCallError::MalformedResult {
                        function: XL_SHEET_ID.name,
                        expected: "external reference containing a sheet ID",
                        actual: "non-reference result",
                    });
                }
                // SAFETY: `xlSheetId` documents the `mref.idSheet` member for
                // its successful `xltypeRef` result; `owner` remains live here.
                Ok(unsafe { root.val.mref.idSheet as u64 })
            }

            /// Returns the name of the *current* sheet, not the active sheet.
            pub fn current_sheet_name(&self) -> Result<ExcelString, ExcelCallError> {
                let mut current = XLOPER12 {
                    val: XLOPER12Value {
                        mref: excel_api_sys::XLOPER12MRef {
                            lpmref: core::ptr::null_mut(),
                            idSheet: 0,
                        },
                    },
                    xltype: excel_api_sys::xltypeRef,
                };
                let mut arguments = [&mut current as *mut XLOPER12];
                let owner = self
                    .capability()
                    .call($permission, XL_SHEET_NM, &mut arguments)?
                    .expect("descriptor requires a result");
                let value = owner
                    .into_owned_value(&crate::ConversionLimits::default())
                    .map_err(|error| ExcelCallError::ResultConversion(error.to_string()))?;
                match value {
                    crate::ExcelValue::Text(name) => Ok(name),
                    other => Err(ExcelCallError::MalformedResult {
                        function: XL_SHEET_NM.name,
                        expected: "sheet name text",
                        actual: other.kind_name(),
                    }),
                }
            }
        }
    };
}

sheet_context!(crate::WorksheetContext<'call>, CallPermission::Worksheet);
sheet_context!(crate::ThreadSafeContext<'call>, CallPermission::ThreadSafe);
sheet_context!(crate::MacroContext<'call>, CallPermission::Macro);

macro_rules! abort_context {
    ($type:ty, $permission:expr) => {
        impl<'call> $type {
            /// Poll whether Excel has received a user cancellation/break request.
            /// This zero-argument form preserves a pending break request.
            pub fn is_cancellation_requested(&self) -> Result<bool, ExcelCallError> {
                self.capability().cancellation_requested($permission, None)
            }

            /// Poll cancellation while explicitly preserving or clearing a pending break.
            pub fn is_cancellation_requested_with(
                &self,
                mode: AbortCheckMode,
            ) -> Result<bool, ExcelCallError> {
                self.capability()
                    .cancellation_requested($permission, Some(mode))
            }
        }
    };
}

abort_context!(crate::WorksheetContext<'call>, CallPermission::Worksheet);
abort_context!(crate::ThreadSafeContext<'call>, CallPermission::ThreadSafe);
abort_context!(crate::MacroContext<'call>, CallPermission::Macro);

macro_rules! caller_context {
    ($type:ty, $permission:expr) => {
        impl<'call> $type {
            /// Returns Excel's caller value with its Excel-owned result lifetime.
            pub fn caller(&self) -> Result<ExcelOwnedValue<'call>, ExcelCallError> {
                let mut arguments = [];
                self.capability()
                    .call($permission, XLF_CALLER, &mut arguments)?
                    .ok_or(ExcelCallError::MalformedResult {
                        function: XLF_CALLER.name,
                        expected: "caller value",
                        actual: "missing result",
                    })
            }
        }
    };
}

caller_context!(crate::WorksheetContext<'call>, CallPermission::Worksheet);
caller_context!(crate::MacroContext<'call>, CallPermission::Macro);

macro_rules! coerce_context {
    ($type:ty, $permission:expr) => {
        impl<'call> $type {
            /// Converts an Excel-owned callback result to a requested root type.
            ///
            /// The source remains owned by the caller and is not consumed or
            /// released by `xlCoerce`; the returned root has its own `xlFree`
            /// obligation encoded in [`ExcelOwnedValue`].
            pub fn coerce(
                &self,
                source: &ExcelOwnedValue<'call>,
                target: CoerceTarget,
            ) -> Result<ExcelOwnedValue<'call>, ExcelCallError> {
                let mut source_root = *source.raw_root();
                let mut requested = XLOPER12 {
                    val: XLOPER12Value { w: target.xltype() },
                    xltype: excel_api_sys::xltypeInt,
                };
                let mut arguments = [
                    &mut source_root as *mut XLOPER12,
                    &mut requested as *mut XLOPER12,
                ];
                self.capability()
                    .call($permission, XL_COERCE, &mut arguments)?
                    .ok_or(ExcelCallError::MalformedResult {
                        function: XL_COERCE.name,
                        expected: "coerced value",
                        actual: "missing result",
                    })
            }
        }
    };
}

coerce_context!(crate::WorksheetContext<'call>, CallPermission::Worksheet);
coerce_context!(crate::ThreadSafeContext<'call>, CallPermission::ThreadSafe);
coerce_context!(crate::MacroContext<'call>, CallPermission::Macro);

#[cfg(test)]
pub(crate) mod test_support {
    use super::*;
    pub(crate) struct UnavailableBackend;
    impl ExcelCallBackend for UnavailableBackend {
        fn link(&self) -> Result<(), ExcelCallError> {
            Err(ExcelCallError::BackendUnavailable)
        }
        fn unlink(&self) {}
        fn is_linked(&self) -> bool {
            false
        }
        unsafe fn excel12v_raw(
            &self,
            _: i32,
            _: *mut XLOPER12,
            _: c_int,
            _: *mut LPXLOPER12,
        ) -> i32 {
            excel_api_sys::xlretFailed
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::sync::Mutex;

    #[derive(Default)]
    struct AbortBackend {
        calls: Mutex<Vec<(i32, usize, Option<i32>)>>,
        result: Mutex<bool>,
        code: Mutex<i32>,
    }

    impl ExcelCallBackend for AbortBackend {
        fn link(&self) -> Result<(), ExcelCallError> {
            Ok(())
        }
        fn unlink(&self) {}
        fn is_linked(&self) -> bool {
            true
        }
        unsafe fn excel12v_raw(
            &self,
            function: i32,
            result: *mut XLOPER12,
            count: c_int,
            arguments: *mut LPXLOPER12,
        ) -> i32 {
            let retain = if count == 1 {
                // SAFETY: test calls supply one initialized Boolean root.
                Some(unsafe { (**arguments).val.xbool })
            } else {
                None
            };
            self.calls
                .lock()
                .unwrap()
                .push((function, count as usize, retain));
            // SAFETY: this test backend initializes the documented Boolean result root.
            unsafe {
                (*result).val = XLOPER12Value {
                    xbool: i32::from(*self.result.lock().unwrap()),
                };
                (*result).xltype = excel_api_sys::xltypeBool;
            }
            *self.code.lock().unwrap()
        }
    }

    #[test]
    fn abort_descriptor_and_zero_retain_clear_forms_are_exact() {
        assert_eq!(XL_ABORT.function, excel_api_sys::xlAbort);
        assert_eq!(XL_ABORT.minimum_arguments, 0);
        assert_eq!(XL_ABORT.maximum_arguments, 1);
        assert_eq!(XL_ABORT.release, ExcelReleasePolicy::NoReleaseRequired);
        assert_eq!(
            XL_ABORT.permissions,
            &[
                CallPermission::Worksheet,
                CallPermission::ThreadSafe,
                CallPermission::Macro,
            ]
        );
        let backend = AbortBackend::default();
        *backend.result.lock().unwrap() = true;
        let capability = CallCapability::new(&backend);
        let context = crate::WorksheetContext::new(&capability);
        assert!(context.is_cancellation_requested().unwrap());
        assert!(
            context
                .is_cancellation_requested_with(AbortCheckMode::PreservePendingBreak)
                .unwrap()
        );
        assert!(
            context
                .is_cancellation_requested_with(AbortCheckMode::ClearPendingBreak)
                .unwrap()
        );
        assert_eq!(
            *backend.calls.lock().unwrap(),
            vec![
                (excel_api_sys::xlAbort, 0, None),
                (excel_api_sys::xlAbort, 1, Some(1)),
                (excel_api_sys::xlAbort, 1, Some(0)),
            ]
        );
    }

    #[test]
    fn abort_is_available_to_verified_contexts_and_preserves_return_codes() {
        let backend = AbortBackend::default();
        let capability = CallCapability::new(&backend);
        assert!(
            !crate::ThreadSafeContext::new(&capability)
                .is_cancellation_requested()
                .unwrap()
        );
        assert!(
            !crate::MacroContext::new(&capability)
                .is_cancellation_requested()
                .unwrap()
        );
        *backend.code.lock().unwrap() = excel_api_sys::xlretAbort | excel_api_sys::xlretUncalced;
        assert_eq!(
            crate::WorksheetContext::new(&capability).is_cancellation_requested(),
            Err(ExcelCallError::ExcelFailure {
                function: "xlAbort",
                code: ExcelReturnCode(excel_api_sys::xlretAbort | excel_api_sys::xlretUncalced),
            })
        );
        assert_eq!(
            capability.cancellation_requested(CallPermission::Lifecycle, None),
            Err(ExcelCallError::IllegalContext {
                function: "xlAbort",
                context: CallPermission::Lifecycle,
            })
        );
        // Only xlAbort appears in the mock trace: its immediate Boolean root
        // creates no ExcelOwnedValue and therefore no xlFree call.
        assert_eq!(backend.calls.lock().unwrap().len(), 3);
    }

    #[test]
    fn selected_m11_descriptors_preserve_sdk_ids_and_contracts() {
        assert_eq!(XL_COERCE.function, excel_api_sys::xlCoerce);
        assert_eq!(XL_COERCE.minimum_arguments, 1);
        assert_eq!(XL_COERCE.maximum_arguments, 2);
        assert_eq!(XL_COERCE.release, ExcelReleasePolicy::XlFreeRequired);
        assert_eq!(XL_SHEET_ID.function, excel_api_sys::xlSheetId);
        assert_eq!(XL_SHEET_ID.minimum_arguments, 0);
        assert_eq!(XL_SHEET_ID.release, ExcelReleasePolicy::NoReleaseRequired);
        assert_eq!(XL_SHEET_NM.function, excel_api_sys::xlSheetNm);
        assert_eq!(XL_SHEET_NM.minimum_arguments, 1);
        assert_eq!(XL_SHEET_NM.release, ExcelReleasePolicy::XlFreeRequired);
        assert_eq!(XLF_CALLER.function, excel_api_sys::xlfCaller);
        assert_eq!(
            XLF_CALLER.permissions,
            &[CallPermission::Worksheet, CallPermission::Macro]
        );
        let caller_is_thread_safe = XLF_CALLER.thread_safe;
        assert!(!caller_is_thread_safe);
    }
}
