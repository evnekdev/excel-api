//! Observation-only Automation values for the Prompt 05H runtime matrix.

use serde::Serialize;
use windows_sys::Win32::Foundation::SysStringLen;
use windows_sys::Win32::System::Variant::{
    VT_ARRAY, VT_BOOL, VT_BSTR, VT_CY, VT_DATE, VT_EMPTY, VT_ERROR, VT_I2, VT_I4, VT_I8, VT_NULL,
    VT_R4, VT_R8,
};

use super::safearray::ObservedSafeArray;
use super::variant::OwnedVariant;

/// Preserves physical VARTYPE separately from the decoded research value.
#[derive(Clone, Debug, Eq, PartialEq, Serialize)]
#[serde(tag = "kind", rename_all = "snake_case")]
pub(super) enum ObservedVariant {
    Empty { vartype: u16 },
    Null { vartype: u16 },
    Bool { vartype: u16, value: bool },
    I16 { vartype: u16, value: i16 },
    I32 { vartype: u16, value: i32 },
    I64 { vartype: u16, value: i64 },
    F32 { vartype: u16, value_bits: u32 },
    F64 { vartype: u16, value_bits: u64 },
    Currency { vartype: u16, scaled_value: i64 },
    Date { vartype: u16, value_bits: u64 },
    String { vartype: u16, utf16: Vec<u16> },
    Error { vartype: u16, scode: i32 },
    Dispatch { vartype: u16 },
    Array { vartype: u16, layout: ObservedSafeArray },
    Other { vartype: u16 },
}

impl ObservedVariant {
    pub(super) fn from_variant(value: &OwnedVariant) -> Self {
        let vartype = value.vt();
        if vartype & VT_ARRAY != 0 {
            let array = unsafe { value.0.Anonymous.Anonymous.Anonymous.parray };
            return unsafe { ObservedSafeArray::inspect(array) }
                .map(|layout| Self::Array { vartype, layout })
                .unwrap_or(Self::Other { vartype });
        }
        unsafe {
            match vartype {
                VT_EMPTY => Self::Empty { vartype },
                VT_NULL => Self::Null { vartype },
                VT_BOOL => Self::Bool {
                    vartype,
                    value: value.0.Anonymous.Anonymous.Anonymous.boolVal != 0,
                },
                VT_I2 => Self::I16 {
                    vartype,
                    value: value.0.Anonymous.Anonymous.Anonymous.iVal,
                },
                VT_I4 => Self::I32 {
                    vartype,
                    value: value.0.Anonymous.Anonymous.Anonymous.lVal,
                },
                VT_I8 => Self::I64 {
                    vartype,
                    value: value.0.Anonymous.Anonymous.Anonymous.llVal,
                },
                VT_R4 => Self::F32 {
                    vartype,
                    value_bits: value.0.Anonymous.Anonymous.Anonymous.fltVal.to_bits(),
                },
                VT_R8 => Self::F64 {
                    vartype,
                    value_bits: value.0.Anonymous.Anonymous.Anonymous.dblVal.to_bits(),
                },
                VT_CY => Self::Currency {
                    vartype,
                    scaled_value: value.0.Anonymous.Anonymous.Anonymous.cyVal.int64,
                },
                VT_DATE => Self::Date {
                    vartype,
                    value_bits: value.0.Anonymous.Anonymous.Anonymous.date.to_bits(),
                },
                VT_BSTR => {
                    let pointer = value.0.Anonymous.Anonymous.Anonymous.bstrVal;
                    let len = usize::try_from(SysStringLen(pointer)).unwrap_or(0);
                    let utf16 = if pointer.is_null() {
                        Vec::new()
                    } else {
                        std::slice::from_raw_parts(pointer, len).to_vec()
                    };
                    Self::String { vartype, utf16 }
                }
                VT_ERROR => Self::Error {
                    vartype,
                    scode: value.0.Anonymous.Anonymous.Anonymous.scode,
                },
                9 => Self::Dispatch { vartype },
                _ => Self::Other { vartype },
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::raw::variant::OwnedVariant;

    #[test]
    fn preserves_i32_vartype_and_value() {
        let value = OwnedVariant::i4(-42);
        assert_eq!(
            ObservedVariant::from_variant(&value),
            ObservedVariant::I32 {
                vartype: VT_I4,
                value: -42,
            }
        );
    }

    #[test]
    fn preserves_embedded_nul_bstr_as_utf16() {
        let input = "left\0right";
        let value = OwnedVariant::bstr(input).expect("BSTR allocation");
        assert_eq!(
            ObservedVariant::from_variant(&value),
            ObservedVariant::String {
                vartype: VT_BSTR,
                utf16: input.encode_utf16().collect(),
            }
        );
    }
}
