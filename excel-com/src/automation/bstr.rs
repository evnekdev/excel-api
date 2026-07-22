use crate::{ConversionError, ExcelComError};
use windows_sys::Win32::Foundation::{SysAllocStringLen, SysFreeString};

/// Private BSTR owner. BSTR length is preserved and no pointer escapes.
#[allow(dead_code)]
pub(crate) struct Bstr(*mut u16);
#[allow(dead_code)]
impl Bstr {
    pub(crate) fn new(text: &str) -> Result<Self, ExcelComError> {
        if text.contains('\0') {
            return Err(ExcelComError::Conversion(ConversionError::EmbeddedNul));
        }
        let units: Vec<u16> = text.encode_utf16().collect();
        // SAFETY: `units` is valid UTF-16 storage for the duration of this allocation call.
        let pointer = unsafe { SysAllocStringLen(units.as_ptr(), units.len() as u32) };
        (!pointer.is_null())
            .then_some(Self(pointer.cast_mut()))
            .ok_or(ExcelComError::Ownership {
                detail: "SysAllocStringLen returned null",
            })
    }
    pub(crate) fn into_raw(self) -> *mut u16 {
        let raw = self.0;
        std::mem::forget(self);
        raw
    }
}
impl Drop for Bstr {
    fn drop(&mut self) {
        // SAFETY: this owner either holds the BSTR allocated above or a null BSTR.
        unsafe { SysFreeString(self.0) }
    }
}
