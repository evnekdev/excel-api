//! Owning `VARIANT` values with exactly-once `VariantClear` cleanup.

use windows_sys::Win32::System::Com::CY;
use windows_sys::Win32::System::Variant::{
    VariantClear, VariantCopy, VariantInit, VARIANT, VT_BOOL, VT_BSTR, VT_CY, VT_DATE, VT_DISPATCH, VT_ERROR,
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

    /// Copies the Automation value with `VariantCopy`.  Prompt 05J uses this
    /// for formula-returned `VT_ERROR` values so the source union bits are not
    /// reconstructed or normalized before a direct write experiment.
    pub(super) fn copy(&self) -> Result<Self, String> {
        let mut result = Self::empty();
        let status = unsafe { VariantCopy(&mut result.0, &self.0) };
        (status == 0)
            .then_some(result)
            .ok_or_else(|| format!("VariantCopy failed with 0x{:08X}", status as u32))
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
        if self.vt() == VT_I4 {
            Some(unsafe { self.0.Anonymous.Anonymous.Anonymous.lVal })
        } else {
            None
        }
    }

    pub(super) fn error_scode(&self) -> Option<i32> {
        if self.vt() == VT_ERROR {
            Some(unsafe { self.0.Anonymous.Anonymous.Anonymous.scode })
        } else {
            None
        }
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn vt_error_constructor_preserves_exact_signed_scode_bits() {
        for scode in [2042_i32, 0x800A_07FA_u32 as i32, -1_i32] {
            let value = OwnedVariant::error(scode);
            assert_eq!(value.vt(), VT_ERROR);
            assert_eq!(value.error_scode(), Some(scode));
            assert_eq!(value.error_scode().expect("scode") as u32, scode as u32);
            unsafe {
                assert_eq!(value.0.Anonymous.Anonymous.wReserved1, 0);
                assert_eq!(value.0.Anonymous.Anonymous.wReserved2, 0);
                assert_eq!(value.0.Anonymous.Anonymous.wReserved3, 0);
            }
        }
    }

    #[test]
    fn variant_copy_preserves_vt_error_raw_bits() {
        let copied = OwnedVariant::error(0x800A_07FA_u32 as i32).copy().expect("VariantCopy");
        assert_eq!(copied.vt(), VT_ERROR);
        assert_eq!(copied.error_scode().map(|value| value as u32), Some(0x800A_07FA));
    }
}
