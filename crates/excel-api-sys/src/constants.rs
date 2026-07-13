/// Base Excel value type mask.
pub const XLTYPE_MASK: u32 = 0x0FFF;

/// Numeric value.
pub const XLTYPE_NUM: u32 = 0x0001;
/// UTF-16 string value.
pub const XLTYPE_STR: u32 = 0x0002;
/// Boolean value.
pub const XLTYPE_BOOL: u32 = 0x0004;
/// Multi-area reference value.
pub const XLTYPE_REF: u32 = 0x0008;
/// Excel error value.
pub const XLTYPE_ERR: u32 = 0x0010;
/// Flow-control value.
pub const XLTYPE_FLOW: u32 = 0x0020;
/// Rectangular multi-cell value.
pub const XLTYPE_MULTI: u32 = 0x0040;
/// Missing function argument.
pub const XLTYPE_MISSING: u32 = 0x0080;
/// Empty value.
pub const XLTYPE_NIL: u32 = 0x0100;
/// Single-area reference value.
pub const XLTYPE_SREF: u32 = 0x0400;
/// Integer value.
pub const XLTYPE_INT: u32 = 0x0800;
/// Big-data value, represented by the combined string and integer type bits.
pub const XLTYPE_BIG_DATA: u32 = XLTYPE_STR | XLTYPE_INT;

/// Excel owns the value and must free it after consuming the result.
pub const XLBIT_XL_FREE: u32 = 0x1000;
/// The XLL owns the value and Excel will call `xlAutoFree12`.
pub const XLBIT_DLL_FREE: u32 = 0x4000;

/// Return code indicating a successful Excel C API call.
pub const XLRET_SUCCESS: i32 = 0;

/// Maximum number of rows supported by modern Excel.
pub const EXCEL12_MAX_ROWS: u32 = 1_048_576;
/// Maximum number of columns supported by modern Excel.
pub const EXCEL12_MAX_COLUMNS: u32 = 16_384;

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn type_constants_match_xlcall_header() {
        assert_eq!(XLTYPE_REF, 0x0008);
        assert_eq!(XLTYPE_INT, 0x0800);
        assert_eq!(XLTYPE_BIG_DATA, XLTYPE_STR | XLTYPE_INT);
    }
}
