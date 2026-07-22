use super::bstr::Bstr;
use crate::{
    ExcelComError,
    internal::{ComPtr, Dispatch},
};
use std::slice;
use windows_sys::Win32::Foundation::SysStringLen;
use windows_sys::Win32::System::Com::CY;
use windows_sys::Win32::System::Variant::{
    VARIANT, VT_BOOL, VT_BSTR, VT_CY, VT_DISPATCH, VT_I4, VT_R8, VariantClear, VariantInit,
};

/// Private initialized VARIANT owner with exactly-once `VariantClear` cleanup.
pub(crate) struct OwnedVariant(pub(crate) VARIANT);
impl OwnedVariant {
    pub(crate) fn empty() -> Self {
        let mut value = VARIANT::default();
        // SAFETY: `value` is writable VARIANT storage initialized by the SDK routine.
        unsafe { VariantInit(&mut value) };
        Self(value)
    }
    pub(crate) fn bool(value: bool) -> Self {
        let mut result = Self::empty();
        result.0.Anonymous.Anonymous.vt = VT_BOOL;
        result.0.Anonymous.Anonymous.Anonymous.boolVal = if value { -1 } else { 0 };
        result
    }
    #[allow(dead_code)]
    pub(crate) fn i32(value: i32) -> Self {
        let mut result = Self::empty();
        result.0.Anonymous.Anonymous.vt = VT_I4;
        result.0.Anonymous.Anonymous.Anonymous.lVal = value;
        result
    }
    #[allow(dead_code)]
    pub(crate) fn f64(value: f64) -> Self {
        let mut result = Self::empty();
        result.0.Anonymous.Anonymous.vt = VT_R8;
        result.0.Anonymous.Anonymous.Anonymous.dblVal = value;
        result
    }
    #[allow(dead_code)]
    pub(crate) fn currency(value: i64) -> Self {
        let mut result = Self::empty();
        result.0.Anonymous.Anonymous.vt = VT_CY;
        result.0.Anonymous.Anonymous.Anonymous.cyVal = CY { int64: value };
        result
    }
    #[allow(dead_code)]
    pub(crate) fn bstr(value: &str) -> Result<Self, ExcelComError> {
        let mut result = Self::empty();
        result.0.Anonymous.Anonymous.vt = VT_BSTR;
        result.0.Anonymous.Anonymous.Anonymous.bstrVal = Bstr::new(value)?.into_raw();
        Ok(result)
    }
    pub(crate) fn vt(&self) -> u16 {
        // SAFETY: reading the VARIANT tag is valid for every initialized VARIANT.
        unsafe { self.0.Anonymous.Anonymous.vt }
    }
    pub(crate) fn as_bool(&self) -> Option<bool> {
        if self.vt() == VT_BOOL {
            // SAFETY: the VT_BOOL tag selects boolVal in the VARIANT union.
            Some(unsafe { self.0.Anonymous.Anonymous.Anonymous.boolVal != 0 })
        } else {
            None
        }
    }
    pub(crate) fn as_i32(&self) -> Option<i32> {
        if self.vt() == VT_I4 {
            // SAFETY: the VT_I4 tag selects lVal in the VARIANT union.
            Some(unsafe { self.0.Anonymous.Anonymous.Anonymous.lVal })
        } else {
            None
        }
    }
    pub(crate) fn as_string(&self) -> Result<String, ExcelComError> {
        if self.vt() != VT_BSTR {
            return Err(ExcelComError::Conversion {
                detail: "expected VT_BSTR result",
            });
        }
        // SAFETY: the VT_BSTR tag selects bstrVal in the VARIANT union.
        let pointer = unsafe { self.0.Anonymous.Anonymous.Anonymous.bstrVal };
        // SAFETY: `pointer` is the BSTR selected by the checked variant tag.
        let length = unsafe { SysStringLen(pointer) } as usize;
        let units = if pointer.is_null() {
            &[]
        } else {
            // SAFETY: SysStringLen bounds the BSTR's UTF-16 allocation.
            unsafe { slice::from_raw_parts(pointer, length) }
        };
        String::from_utf16(units).map_err(|_| ExcelComError::Conversion {
            detail: "BSTR contained invalid UTF-16",
        })
    }
    pub(crate) fn take_dispatch(&mut self) -> Result<ComPtr<Dispatch>, ExcelComError> {
        if self.vt() != VT_DISPATCH {
            return Err(ExcelComError::Conversion {
                detail: "expected VT_DISPATCH result",
            });
        }
        // SAFETY: the VT_DISPATCH tag selects pdispVal in the VARIANT union.
        let raw = unsafe { self.0.Anonymous.Anonymous.Anonymous.pdispVal };
        self.0.Anonymous.Anonymous.Anonymous.pdispVal = std::ptr::null_mut();
        self.0.Anonymous.Anonymous.vt = 0;
        // SAFETY: ownership moved out of the VARIANT by clearing its tag and pointer slot.
        unsafe { ComPtr::from_owned(raw) }
    }
}
impl Drop for OwnedVariant {
    fn drop(&mut self) {
        // SAFETY: VariantClear releases only resources owned by this initialized VARIANT.
        unsafe {
            let _ = VariantClear(&mut self.0);
        }
    }
}
