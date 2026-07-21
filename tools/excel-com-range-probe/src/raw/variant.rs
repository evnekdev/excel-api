//! Owning `VARIANT` values with exactly-once `VariantClear` cleanup.

use windows_sys::Win32::System::Com::CY;
use windows_sys::Win32::System::Variant::{
    VariantClear, VariantInit, VARIANT, VT_BOOL, VT_BSTR, VT_CY, VT_DATE, VT_DISPATCH, VT_ERROR,
    VT_I2, VT_I4, VT_I8, VT_NULL, VT_R4, VT_R8,
};

use super::bstr::OwnedBstr;
use super::com_ptr::{ComPtr, Dispatch};
use super::safearray::OwnedSafeArray;

/// Owns one initialized Automation `VARIANT`.
pub(super) struct OwnedVariant(pub(super) VARIANT);

#[allow(dead_code)]
impl OwnedVariant {
    pub(super) fn empty() -> Self {
        let mut value = VARIANT::default();
        unsafe { VariantInit(&mut value) };
        Self(value)
    }

    pub(super) fn boolean(value: bool) -> Self {
        let mut result = Self::empty();
        result.0.Anonymous.Anonymous.vt = VT_BOOL;
        result.0.Anonymous.Anonymous.Anonymous.boolVal = if value { -1 } else { 0 };
        result
    }

    pub(super) fn null() -> Self {
        let mut result = Self::empty();
        result.0.Anonymous.Anonymous.vt = VT_NULL;
        result
    }

    pub(super) fn i2(value: i16) -> Self {
        let mut result = Self::empty();
        result.0.Anonymous.Anonymous.vt = VT_I2;
        result.0.Anonymous.Anonymous.Anonymous.iVal = value;
        result
    }

    pub(super) fn i4(value: i32) -> Self {
        let mut result = Self::empty();
        result.0.Anonymous.Anonymous.vt = VT_I4;
        result.0.Anonymous.Anonymous.Anonymous.lVal = value;
        result
    }

    pub(super) fn i8(value: i64) -> Self {
        let mut result = Self::empty();
        result.0.Anonymous.Anonymous.vt = VT_I8;
        result.0.Anonymous.Anonymous.Anonymous.llVal = value;
        result
    }

    pub(super) fn r4(value: f32) -> Self {
        let mut result = Self::empty();
        result.0.Anonymous.Anonymous.vt = VT_R4;
        result.0.Anonymous.Anonymous.Anonymous.fltVal = value;
        result
    }

    pub(super) fn r8(value: f64) -> Self {
        let mut result = Self::empty();
        result.0.Anonymous.Anonymous.vt = VT_R8;
        result.0.Anonymous.Anonymous.Anonymous.dblVal = value;
        result
    }

    pub(super) fn error(scode: i32) -> Self {
        let mut result = Self::empty();
        result.0.Anonymous.Anonymous.vt = VT_ERROR;
        result.0.Anonymous.Anonymous.Anonymous.scode = scode;
        result
    }

    pub(super) fn date(value: f64) -> Self {
        let mut result = Self::empty();
        result.0.Anonymous.Anonymous.vt = VT_DATE;
        result.0.Anonymous.Anonymous.Anonymous.date = value;
        result
    }

    pub(super) fn currency(scaled_value: i64) -> Self {
        let mut result = Self::empty();
        result.0.Anonymous.Anonymous.vt = VT_CY;
        result.0.Anonymous.Anonymous.Anonymous.cyVal = CY {
            int64: scaled_value,
        };
        result
    }

    pub(super) fn bstr(text: &str) -> Result<Self, String> {
        let bstr = OwnedBstr::from_text(text)?;
        let mut result = Self::empty();
        result.0.Anonymous.Anonymous.vt = VT_BSTR;
        result.0.Anonymous.Anonymous.Anonymous.bstrVal = bstr.0;
        std::mem::forget(bstr);
        Ok(result)
    }

    /// Transfers a `SAFEARRAY(VARIANT)` to the VARIANT owner, which destroys
    /// it through `VariantClear` exactly once.
    pub(super) fn array(array: OwnedSafeArray) -> Self {
        let mut result = Self::empty();
        result.0.Anonymous.Anonymous.vt = windows_sys::Win32::System::Variant::VT_ARRAY
            | windows_sys::Win32::System::Variant::VT_VARIANT;
        result.0.Anonymous.Anonymous.Anonymous.parray = array.into_raw();
        result
    }

    pub(super) fn vt(&self) -> u16 {
        unsafe { self.0.Anonymous.Anonymous.vt }
    }

    pub(super) fn i4_value(&self) -> Option<i32> {
        (self.vt() == VT_I4).then(|| unsafe { self.0.Anonymous.Anonymous.Anonymous.lVal })
    }

    pub(super) fn is_exact_42(&self) -> bool {
        match self.vt() {
            value if value == VT_I4 => self.i4_value() == Some(42),
            value if value == VT_R8 => unsafe { self.0.Anonymous.Anonymous.Anonymous.dblVal == 42.0 },
            _ => false,
        }
    }

    /// Clones the dispatch reference held by this `VARIANT` before the variant
    /// releases its own reference during `Drop`.
    pub(super) fn dispatch(&self) -> Option<ComPtr<Dispatch>> {
        if self.vt() != VT_DISPATCH {
            None
        } else {
            unsafe { ComPtr::from_borrowed(self.0.Anonymous.Anonymous.Anonymous.pdispVal) }
        }
    }
}

impl Drop for OwnedVariant {
    fn drop(&mut self) {
        unsafe {
            let _ = VariantClear(&mut self.0);
        }
    }
}
