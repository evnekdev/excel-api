//! BSTR allocation and release ownership.

use windows_sys::Win32::Foundation::{SysAllocStringLen, SysFreeString};

/// Owns one Automation BSTR allocated with `SysAllocString`.
pub(super) struct OwnedBstr(pub(super) *const u16);

impl OwnedBstr {
    pub(super) fn from_text(text: &str) -> Result<Self, String> {
        let wide: Vec<u16> = text.encode_utf16().collect();
        let length = u32::try_from(wide.len()).map_err(|_| "BSTR exceeds u32 UTF-16 units")?;
        let value = unsafe { SysAllocStringLen(wide.as_ptr(), length) };
        if value.is_null() {
            Err("SysAllocString returned null".to_owned())
        } else {
            Ok(Self(value))
        }
    }
}

impl Drop for OwnedBstr {
    fn drop(&mut self) {
        if !self.0.is_null() {
            unsafe { SysFreeString(self.0) }
        }
    }
}
