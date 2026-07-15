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
/// Mutable pointer to a variable-length multi-reference header.
///
/// The pointer has no Rust ownership or bounds guarantee.
pub type LPXLMREF12 = *mut XLMREF12;
/// Mutable pointer to a variable-length `FP12` header.
///
/// The pointer has no Rust ownership or bounds guarantee.
pub type LPFP12 = *mut FP12;
/// Mutable pointer to an `XLOPER12` value.
///
/// The pointer has no Rust ownership or callback-lifetime guarantee.
pub type LPXLOPER12 = *mut XLOPER12;

/// One rectangular Excel 12 reference.
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct XLREF12 {
    /// First row, inclusive, in Excel's zero-based raw coordinate system.
    pub rwFirst: RW,
    /// Last row, inclusive, in Excel's zero-based raw coordinate system.
    pub rwLast: RW,
    /// First column, inclusive, in Excel's zero-based raw coordinate system.
    pub colFirst: COL,
    /// Last column, inclusive, in Excel's zero-based raw coordinate system.
    pub colLast: COL,
}

/// Variable-length table of Excel 12 references.
///
/// The SDK declares `reftbl[1]`, but an allocation contains `count` entries.
/// This raw definition intentionally preserves the one-element C tail.
#[repr(C)]
#[derive(Clone, Copy, Debug)]
pub struct XLMREF12 {
    /// Number of valid rectangular references in the trailing allocation.
    pub count: WORD,
    /// First element of the C flexible-array tail.
    ///
    /// # Safety
    ///
    /// Only index past this element after proving the allocation is large
    /// enough for `count` `XLREF12` values and remains valid for the access.
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
    /// Number of rows in the flexible-array payload.
    pub rows: i32,
    /// Number of columns in the flexible-array payload.
    pub columns: i32,
    /// First element of the C flexible-array tail.
    ///
    /// # Safety
    ///
    /// Only index past this element after validating the `rows * columns`
    /// allocation extent and pointer lifetime.
    pub array: [f64; 1],
}

/// Inline single-area reference member of `XLOPER12`.
#[repr(C)]
#[derive(Clone, Copy, Debug)]
pub struct XLOPER12SRef {
    /// Number of areas represented by this inline member; the supported single
    /// reference form has exactly one.
    pub count: WORD,
    /// Named `ref` in the C header; `reference` avoids the Rust keyword.
    pub reference: XLREF12,
}

/// External multi-area reference member of `XLOPER12`.
#[repr(C)]
#[derive(Clone, Copy, Debug)]
pub struct XLOPER12MRef {
    /// Pointer to the externally allocated multi-reference table.
    ///
    /// Its validity, allocation origin, and cleanup protocol are determined by
    /// the enclosing `XLOPER12` tag and Excel callback contract.
    pub lpmref: LPXLMREF12,
    /// Excel sheet identifier for the referenced areas.
    pub idSheet: IDSHEET,
}

/// Rectangular mixed-value array member of `XLOPER12`.
#[repr(C)]
#[derive(Clone, Copy, Debug)]
pub struct XLOPER12Array {
    /// Pointer to the first raw `XLOPER12` array element.
    ///
    /// The pointer must be interpreted only with a valid `xltypeMulti` tag and
    /// a proven `rows * columns` extent.
    pub lparray: LPXLOPER12,
    /// Number of array rows.
    pub rows: RW,
    /// Number of array columns.
    pub columns: COL,
}

/// Variant storage within the `xltypeFlow` member.
#[repr(C)]
#[derive(Clone, Copy)]
pub union XLOPER12FlowValue {
    /// Flow-control nesting level.
    pub level: i32,
    /// Toolbar-control identifier.
    pub tbctrl: i32,
    /// Sheet identifier for the applicable flow operation.
    pub idSheet: IDSHEET,
}

/// Flow-control member of `XLOPER12`.
#[repr(C)]
#[derive(Clone, Copy)]
pub struct XLOPER12Flow {
    /// Tagged flow-control payload.
    pub valflow: XLOPER12FlowValue,
    /// Row associated with the flow operation.
    pub rw: RW,
    /// Column associated with the flow operation.
    pub col: COL,
    /// SDK flow-control discriminator.
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
    /// Pointer-or-handle payload selected by the enclosing operation.
    pub h: XLOPER12BigDataHandle,
    /// Windows `long` is 32-bit on the supported MSVC target.
    pub cbData: i32,
}

/// Complete value union from the SDK's `XLOPER12` definition.
#[repr(C)]
#[derive(Clone, Copy)]
pub union XLOPER12Value {
    /// IEEE-754 numeric payload for `xltypeNum`.
    pub num: f64,
    /// Pointer to an Excel UTF-16 string representation.
    ///
    /// Its prefix/terminator convention and lifetime are selected by the
    /// enclosing `xltype` and callback contract.
    pub str: *mut XCHAR,
    /// Excel Boolean payload for `xltypeBool`.
    pub xbool: BOOL,
    /// Excel error code payload for `xltypeErr`.
    pub err: i32,
    /// Excel integer payload for `xltypeInt`.
    pub w: i32,
    /// Inline single-reference payload for `xltypeSRef`.
    pub sref: XLOPER12SRef,
    /// External multi-reference payload for `xltypeRef`.
    pub mref: XLOPER12MRef,
    /// Mixed-value array payload for `xltypeMulti`.
    pub array: XLOPER12Array,
    /// Flow-control payload for `xltypeFlow`.
    pub flow: XLOPER12Flow,
    /// Big-data payload for `xltypeBigData`.
    pub bigdata: XLOPER12BigData,
}

/// Excel 12's universal raw value container.
#[repr(C)]
#[derive(Clone, Copy)]
pub struct XLOPER12 {
    /// Untagged payload union; read only the member selected by [`Self::xltype`].
    pub val: XLOPER12Value,
    /// Excel type tag plus protocol ownership bits.
    ///
    /// This field does not itself transfer ownership. Its low type bits select
    /// the valid union member; ownership bits must be interpreted under the
    /// documented Excel callback or return protocol.
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
