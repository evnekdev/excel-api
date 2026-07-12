use core::ffi::c_void;

/// Excel Boolean storage.
pub type XlBool = i32;

/// One rectangular cell area.
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct XLREF12 {
    pub row_first: i32,
    pub row_last: i32,
    pub column_first: i32,
    pub column_last: i32,
}

/// Single-sheet reference value.
#[repr(C)]
#[derive(Clone, Copy, Debug)]
pub struct XlSRef12 {
    pub count: u16,
    pub reference: XLREF12,
}

/// Rectangular array value.
#[repr(C)]
#[derive(Clone, Copy, Debug)]
pub struct XlArray12 {
    pub values: *mut XLOPER12,
    pub rows: i32,
    pub columns: i32,
}

/// UTF-16 string pointer.
///
/// Excel strings are length-prefixed; the first UTF-16 code unit stores the
/// number of following code units.
#[repr(transparent)]
#[derive(Clone, Copy, Debug)]
pub struct XlString12(pub *mut u16);

/// Partial raw value union used by the initial implementation.
///
/// More fields will be added only after comparison with the official Excel SDK
/// headers and ABI verification tests.
#[repr(C)]
#[derive(Clone, Copy)]
pub union XlOperValue12 {
    pub number: f64,
    pub string: XlString12,
    pub boolean: XlBool,
    pub error: i32,
    pub array: XlArray12,
    pub sref: XlSRef12,
    pub pointer: *mut c_void,
}

/// Universal Excel value container.
#[repr(C)]
#[derive(Clone, Copy)]
pub struct XLOPER12 {
    pub value: XlOperValue12,
    pub xltype: u32,
}

impl XLOPER12 {
    /// Construct a raw numeric value.
    pub const fn number(value: f64) -> Self {
        Self {
            value: XlOperValue12 { number: value },
            xltype: crate::XLTYPE_NUM,
        }
    }

    /// Construct a raw Boolean value.
    pub const fn boolean(value: bool) -> Self {
        Self {
            value: XlOperValue12 {
                boolean: value as XlBool,
            },
            xltype: crate::XLTYPE_BOOL,
        }
    }

    /// Construct a raw Excel error value.
    pub const fn error(value: crate::XlError) -> Self {
        Self {
            value: XlOperValue12 {
                error: value as i32,
            },
            xltype: crate::XLTYPE_ERR,
        }
    }

    /// Construct a raw missing value.
    pub const fn missing() -> Self {
        Self {
            value: XlOperValue12 { number: 0.0 },
            xltype: crate::XLTYPE_MISSING,
        }
    }

    /// Construct a raw empty value.
    pub const fn nil() -> Self {
        Self {
            value: XlOperValue12 { number: 0.0 },
            xltype: crate::XLTYPE_NIL,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn xloper12_is_large_enough_for_union_members() {
        assert!(core::mem::size_of::<XLOPER12>() >= core::mem::size_of::<XlArray12>() + 4);
    }

    #[test]
    fn xlref12_has_expected_field_width() {
        assert_eq!(core::mem::size_of::<XLREF12>(), 16);
    }
}
