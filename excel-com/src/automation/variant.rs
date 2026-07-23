//! Private initialized VARIANT ownership.

use std::slice;

use windows_sys::Win32::Foundation::SysStringLen;
use windows_sys::Win32::System::Com::CY;
use windows_sys::Win32::System::Variant::{
    VARIANT, VT_ARRAY, VT_BOOL, VT_BSTR, VT_CY, VT_DATE, VT_DISPATCH, VT_EMPTY, VT_ERROR, VT_I4,
    VT_NULL, VT_R4, VT_R8, VT_UNKNOWN, VT_VARIANT, VariantClear, VariantInit,
};

use super::SafeArray;
use super::bstr::Bstr;
use crate::{
    ConversionError, ExcelComError,
    internal::{ComPtr, Dispatch, Unknown},
};

/// Private initialized `VARIANT` owner with exactly-once `VariantClear` cleanup.
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

    pub(crate) fn i32(value: i32) -> Self {
        let mut result = Self::empty();
        result.0.Anonymous.Anonymous.vt = VT_I4;
        result.0.Anonymous.Anonymous.Anonymous.lVal = value;
        result
    }

    pub(crate) fn f64(value: f64) -> Self {
        let mut result = Self::empty();
        result.0.Anonymous.Anonymous.vt = VT_R8;
        result.0.Anonymous.Anonymous.Anonymous.dblVal = value;
        result
    }

    /// Creates an exact `VT_R4` value for Office drawing members whose
    /// registered signature is single precision.
    pub(crate) fn f32(value: f32) -> Self {
        let mut result = Self::empty();
        result.0.Anonymous.Anonymous.vt = VT_R4;
        result.0.Anonymous.Anonymous.Anonymous.fltVal = value;
        result
    }

    pub(crate) fn null() -> Self {
        let mut result = Self::empty();
        result.0.Anonymous.Anonymous.vt = VT_NULL;
        result
    }

    pub(crate) fn error(value: i32) -> Self {
        let mut result = Self::empty();
        result.0.Anonymous.Anonymous.vt = VT_ERROR;
        result.0.Anonymous.Anonymous.Anonymous.scode = value;
        result
    }

    pub(crate) fn date(value: f64) -> Self {
        let mut result = Self::empty();
        result.0.Anonymous.Anonymous.vt = VT_DATE;
        result.0.Anonymous.Anonymous.Anonymous.date = value;
        result
    }

    pub(crate) fn currency(value: i64) -> Self {
        let mut result = Self::empty();
        result.0.Anonymous.Anonymous.vt = VT_CY;
        result.0.Anonymous.Anonymous.Anonymous.cyVal = CY { int64: value };
        result
    }

    pub(crate) fn bstr(value: &str) -> Result<Self, ExcelComError> {
        Self::bstr_wide(&value.encode_utf16().collect::<Vec<_>>())
    }

    pub(crate) fn bstr_wide(value: &[u16]) -> Result<Self, ExcelComError> {
        let mut result = Self::empty();
        result.0.Anonymous.Anonymous.vt = VT_BSTR;
        result.0.Anonymous.Anonymous.Anonymous.bstrVal = Bstr::from_wide(value)?.into_raw();
        Ok(result)
    }

    /// Encodes a borrowed `IDispatch` while giving the VARIANT its own COM reference.
    pub(crate) fn dispatch_borrowed(value: &ComPtr<Dispatch>) -> Self {
        let mut result = Self::empty();
        result.0.Anonymous.Anonymous.vt = VT_DISPATCH;
        result.0.Anonymous.Anonymous.Anonymous.pdispVal = value.clone().into_raw();
        result
    }

    pub(crate) fn array(value: SafeArray) -> Self {
        let mut result = Self::empty();
        result.0.Anonymous.Anonymous.vt = VT_ARRAY | VT_VARIANT;
        result.0.Anonymous.Anonymous.Anonymous.parray = value.into_raw();
        result
    }

    pub(crate) fn vt(&self) -> u16 {
        // SAFETY: every instance is initialized and the VARIANT tag is always readable.
        unsafe { self.0.Anonymous.Anonymous.vt }
    }

    pub(crate) fn as_bool(&self) -> Option<bool> {
        if self.vt() != VT_BOOL {
            return None;
        }
        // SAFETY: the checked VT_BOOL tag selects boolVal in the VARIANT union.
        Some(unsafe { self.0.Anonymous.Anonymous.Anonymous.boolVal != 0 })
    }

    pub(crate) fn as_i32(&self) -> Option<i32> {
        if self.vt() != VT_I4 {
            return None;
        }
        // SAFETY: the checked VT_I4 tag selects lVal in the VARIANT union.
        Some(unsafe { self.0.Anonymous.Anonymous.Anonymous.lVal })
    }

    pub(crate) fn as_f64(&self) -> Option<f64> {
        if self.vt() != VT_R8 {
            return None;
        }
        // SAFETY: the checked VT_R8 tag selects dblVal in the VARIANT union.
        Some(unsafe { self.0.Anonymous.Anonymous.Anonymous.dblVal })
    }

    pub(crate) fn as_scode(&self) -> Option<i32> {
        if self.vt() != VT_ERROR {
            return None;
        }
        // SAFETY: the checked VT_ERROR tag selects scode in the VARIANT union.
        Some(unsafe { self.0.Anonymous.Anonymous.Anonymous.scode })
    }

    pub(crate) fn as_string(&self) -> Result<String, ExcelComError> {
        if self.vt() != VT_BSTR {
            return Err(ExcelComError::Conversion(
                ConversionError::UnsupportedVariantType { vartype: self.vt() },
            ));
        }
        // SAFETY: the checked VT_BSTR tag selects bstrVal in the VARIANT union.
        let pointer = unsafe { self.0.Anonymous.Anonymous.Anonymous.bstrVal };
        // SAFETY: the BSTR pointer selected by VT_BSTR may be null and SysStringLen accepts null.
        let length = unsafe { SysStringLen(pointer) } as usize;
        let units = if pointer.is_null() {
            &[]
        } else {
            // SAFETY: SysStringLen bounds this BSTR's UTF-16 allocation.
            unsafe { slice::from_raw_parts(pointer, length) }
        };
        String::from_utf16(units)
            .map_err(|_| ExcelComError::Conversion(ConversionError::InvalidUtf16String))
    }

    pub(crate) fn take_dispatch(&mut self) -> Result<ComPtr<Dispatch>, ExcelComError> {
        if self.vt() != VT_DISPATCH {
            return Err(ExcelComError::Conversion(
                ConversionError::UnsupportedVariantType { vartype: self.vt() },
            ));
        }
        // SAFETY: the checked VT_DISPATCH tag selects pdispVal in the VARIANT union.
        let raw = unsafe { self.0.Anonymous.Anonymous.Anonymous.pdispVal };
        self.0.Anonymous.Anonymous.Anonymous.pdispVal = std::ptr::null_mut();
        self.0.Anonymous.Anonymous.vt = 0;
        // SAFETY: ownership moved out by clearing the tag and pointer slot above.
        unsafe { ComPtr::from_owned(raw) }
    }

    /// Moves out a dispatch result, treating an Automation null result as no object.
    ///
    /// Excel's `Range.Find` family reports a no-match result without a Range
    /// object. Both null/empty variants and a null `VT_DISPATCH` pointer are
    /// accepted here; every other VARTYPE remains a structured conversion
    /// failure.
    pub(crate) fn take_optional_dispatch(
        &mut self,
    ) -> Result<Option<ComPtr<Dispatch>>, ExcelComError> {
        if matches!(self.vt(), VT_EMPTY | VT_NULL) {
            return Ok(None);
        }
        if self.vt() != VT_DISPATCH {
            return Err(ExcelComError::Conversion(
                ConversionError::UnsupportedVariantType { vartype: self.vt() },
            ));
        }
        // SAFETY: the checked VT_DISPATCH tag selects pdispVal in the VARIANT union.
        let raw = unsafe { self.0.Anonymous.Anonymous.Anonymous.pdispVal };
        self.0.Anonymous.Anonymous.Anonymous.pdispVal = std::ptr::null_mut();
        self.0.Anonymous.Anonymous.vt = 0;
        if raw.is_null() {
            return Ok(None);
        }
        // SAFETY: ownership moved out by clearing the tag and pointer slot above.
        unsafe { ComPtr::from_owned(raw).map(Some) }
    }

    pub(crate) fn take_unknown(&mut self) -> Result<ComPtr<Unknown>, ExcelComError> {
        if self.vt() != VT_UNKNOWN {
            return Err(ExcelComError::Conversion(
                ConversionError::UnsupportedVariantType { vartype: self.vt() },
            ));
        }
        // SAFETY: the checked VT_UNKNOWN tag selects punkVal in the VARIANT union.
        let raw = unsafe { self.0.Anonymous.Anonymous.Anonymous.punkVal };
        self.0.Anonymous.Anonymous.Anonymous.punkVal = std::ptr::null_mut();
        self.0.Anonymous.Anonymous.vt = 0;
        // SAFETY: ownership moved out by clearing the tag and pointer slot above.
        unsafe { ComPtr::from_owned(raw) }
    }
}

impl Drop for OwnedVariant {
    fn drop(&mut self) {
        // SAFETY: VariantClear releases resources owned by this initialized VARIANT.
        unsafe {
            let _ = VariantClear(&mut self.0);
        }
    }
}
