use core::ffi::c_void;

/// Excel 12 Boolean storage (`INT32` in `XLCALL.H`).
pub type BOOL = i32;
/// Excel 12 wide character (`WCHAR` in `XLCALL.H`).
pub type XCHAR = u16;
/// Excel 12 row index.
pub type RW = i32;
/// Excel 12 column index.
pub type COL = i32;
/// Excel 12 sheet identifier (`DWORD_PTR` in `XLCALL.H`).
pub type IDSHEET = usize;
/// Windows byte used by the Excel C API.
pub type BYTE = u8;
/// Windows word used by the Excel C API.
pub type WORD = u16;
/// Opaque Windows handle used by `xltypeBigData` results.
pub type HANDLE = *mut c_void;

/// Pointer aliases used by the SDK declarations.
pub type LPXLREF12 = *mut XLREF12;
pub type LPXLMREF12 = *mut XLMREF12;
pub type LPFP12 = *mut FP12;
pub type LPXLOPER12 = *mut XLOPER12;

/// One rectangular Excel 12 reference.
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct XLREF12 {
    pub rwFirst: RW,
    pub rwLast: RW,
    pub colFirst: COL,
    pub colLast: COL,
}

/// Variable-length table of Excel 12 references.
///
/// The SDK declares `reftbl[1]`, but an allocation contains `count` entries.
/// This raw definition intentionally preserves the one-element C tail.
#[repr(C)]
#[derive(Clone, Copy, Debug)]
pub struct XLMREF12 {
    pub count: WORD,
    pub reftbl: [XLREF12; 1],
}

/// Variable-length rectangular floating-point array.
///
/// The SDK declares `array[1]`, but an allocation contains `rows * columns`
/// entries. This raw definition intentionally preserves the one-element C
/// tail.
#[repr(C)]
#[derive(Clone, Copy, Debug)]
pub struct FP12 {
    pub rows: i32,
    pub columns: i32,
    pub array: [f64; 1],
}

/// Inline single-area reference member of `XLOPER12`.
#[repr(C)]
#[derive(Clone, Copy, Debug)]
pub struct XLOPER12SRef {
    pub count: WORD,
    /// Named `ref` in the C header; `reference` avoids the Rust keyword.
    pub reference: XLREF12,
}

/// External multi-area reference member of `XLOPER12`.
#[repr(C)]
#[derive(Clone, Copy, Debug)]
pub struct XLOPER12MRef {
    pub lpmref: LPXLMREF12,
    pub idSheet: IDSHEET,
}

/// Rectangular mixed-value array member of `XLOPER12`.
#[repr(C)]
#[derive(Clone, Copy, Debug)]
pub struct XLOPER12Array {
    pub lparray: LPXLOPER12,
    pub rows: RW,
    pub columns: COL,
}

/// Variant storage within the `xltypeFlow` member.
#[repr(C)]
#[derive(Clone, Copy)]
pub union XLOPER12FlowValue {
    pub level: i32,
    pub tbctrl: i32,
    pub idSheet: IDSHEET,
}

/// Flow-control member of `XLOPER12`.
#[repr(C)]
#[derive(Clone, Copy)]
pub struct XLOPER12Flow {
    pub valflow: XLOPER12FlowValue,
    pub rw: RW,
    pub col: COL,
    pub xlflow: BYTE,
}

/// Pointer-or-handle storage within the `xltypeBigData` member.
#[repr(C)]
#[derive(Clone, Copy)]
pub union XLOPER12BigDataHandle {
    /// Input buffer passed to Excel.
    pub lpbData: *mut BYTE,
    /// Handle returned by Excel.
    pub hdata: HANDLE,
}

/// Binary-data member of `XLOPER12`.
#[repr(C)]
#[derive(Clone, Copy)]
pub struct XLOPER12BigData {
    pub h: XLOPER12BigDataHandle,
    /// Windows `long` is 32-bit on the supported MSVC target.
    pub cbData: i32,
}

/// Complete value union from the SDK's `XLOPER12` definition.
#[repr(C)]
#[derive(Clone, Copy)]
pub union XLOPER12Value {
    pub num: f64,
    pub str: *mut XCHAR,
    pub xbool: BOOL,
    pub err: i32,
    pub w: i32,
    pub sref: XLOPER12SRef,
    pub mref: XLOPER12MRef,
    pub array: XLOPER12Array,
    pub flow: XLOPER12Flow,
    pub bigdata: XLOPER12BigData,
}

/// Excel 12's universal raw value container.
#[repr(C)]
#[derive(Clone, Copy)]
pub struct XLOPER12 {
    pub val: XLOPER12Value,
    pub xltype: u32,
}

#[cfg(test)]
mod tests {
    use core::mem::{align_of, offset_of, size_of};

    use super::*;

    #[test]
    fn fixed_width_primitives_match_the_sdk() {
        assert_eq!(size_of::<BOOL>(), 4);
        assert_eq!(size_of::<XCHAR>(), 2);
        assert_eq!(size_of::<RW>(), 4);
        assert_eq!(size_of::<COL>(), 4);
        assert_eq!(size_of::<IDSHEET>(), size_of::<usize>());
    }

    #[test]
    fn flexible_array_headers_preserve_one_c_element() {
        assert_eq!(offset_of!(XLMREF12, reftbl), 4);
        assert_eq!(size_of::<XLMREF12>(), 20);
        assert_eq!(offset_of!(FP12, array), 8);
        assert_eq!(size_of::<FP12>(), 16);
    }

    #[cfg(all(target_os = "windows", target_arch = "x86_64"))]
    #[test]
    fn xloper12_has_windows_x64_sdk_layout() {
        assert_eq!(size_of::<XLOPER12Value>(), 24);
        assert_eq!(align_of::<XLOPER12Value>(), 8);
        assert_eq!(offset_of!(XLOPER12, val), 0);
        assert_eq!(offset_of!(XLOPER12, xltype), 24);
        assert_eq!(size_of::<XLOPER12>(), 32);
        assert_eq!(align_of::<XLOPER12>(), 8);
    }
}
