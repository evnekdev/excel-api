//! BSTR decoding retained separately from the numeric and SAFEARRAY codec.

use std::slice;

use windows_sys::Win32::Foundation::SysStringLen;

use super::{AutomationValue, ConversionError};

pub(super) fn decode_bstr(pointer: *const u16) -> Result<AutomationValue, ConversionError> {
    let length = usize::try_from(unsafe { SysStringLen(pointer) })
        .map_err(|_| ConversionError::StringTooLong)?;
    let units = if pointer.is_null() {
        &[]
    } else {
        unsafe { slice::from_raw_parts(pointer, length) }
    };
    String::from_utf16(units)
        .map(AutomationValue::Text)
        .map_err(|_| ConversionError::InvalidUtf16String)
}
