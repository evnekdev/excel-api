//! Narrow Excel12v callback capability and the M8 call catalogue.

use core::{
    ffi::{c_char, c_int, c_void},
    fmt, ptr,
};
use std::sync::atomic::{AtomicUsize, Ordering};

use excel_api_sys::{LPXLOPER12, XLOPER12, XLOPER12Value};

use crate::excel_owned::ExcelReleaseBackend;
use crate::{
    ExcelOwnedValue, ExcelReleaseError, ExcelReleasePolicy, ExcelString, FunctionRegistration,
    LifecycleContext, RegistrationError,
};

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum CallPermission {
    Lifecycle,
    Worksheet,
    ThreadSafe,
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
    pub permission: CallPermission,
    pub result: ResultRoot,
    pub release: ExcelReleasePolicy,
    pub thread_safe: bool,
    pub minimum_arguments: usize,
    pub maximum_arguments: usize,
}

pub const XL_GET_NAME: ExcelCallDescriptor = ExcelCallDescriptor {
    name: "xlGetName",
    function: excel_api_sys::xlGetName,
    permission: CallPermission::Lifecycle,
    result: ResultRoot::Required,
    release: ExcelReleasePolicy::XlFreeRequired,
    thread_safe: false,
    minimum_arguments: 0,
    maximum_arguments: 0,
};
pub const XLF_REGISTER: ExcelCallDescriptor = ExcelCallDescriptor {
    name: "xlfRegister",
    function: excel_api_sys::xlfRegister,
    permission: CallPermission::Lifecycle,
    result: ResultRoot::Required,
    release: ExcelReleasePolicy::XlFreeRequired,
    thread_safe: false,
    minimum_arguments: 10,
    maximum_arguments: 255,
};
pub const XLF_UNREGISTER: ExcelCallDescriptor = ExcelCallDescriptor {
    name: "xlfUnregister",
    function: excel_api_sys::xlfUnregister,
    permission: CallPermission::Lifecycle,
    result: ResultRoot::Required,
    release: ExcelReleasePolicy::XlFreeRequired,
    thread_safe: false,
    minimum_arguments: 1,
    maximum_arguments: 1,
};
pub const XLF_SET_NAME: ExcelCallDescriptor = ExcelCallDescriptor {
    name: "xlfSetName",
    function: excel_api_sys::xlfSetName,
    permission: CallPermission::Lifecycle,
    result: ResultRoot::Required,
    release: ExcelReleasePolicy::XlFreeRequired,
    thread_safe: false,
    minimum_arguments: 1,
    maximum_arguments: 2,
};
pub const XL_FREE: ExcelCallDescriptor = ExcelCallDescriptor {
    name: "xlFree",
    function: excel_api_sys::xlFree,
    permission: CallPermission::ThreadSafe,
    result: ResultRoot::None,
    release: ExcelReleasePolicy::NoReleaseRequired,
    thread_safe: true,
    minimum_arguments: 1,
    maximum_arguments: 255,
};

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
        descriptor: ExcelCallDescriptor,
        arguments: &mut [LPXLOPER12],
    ) -> Result<Option<ExcelOwnedValue<'call>>, ExcelCallError> {
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
            .call(XL_GET_NAME, &mut arguments)?
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
            .call(XLF_REGISTER, &mut arguments)?
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
            .call(XLF_UNREGISTER, &mut arguments)?
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
            .call(XLF_SET_NAME, &mut arguments)?
            .expect("descriptor requires a result");
        owner
            .release()
            .map_err(|error| ExcelCallError::ResultConversion(error.to_string()))
    }
}

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
