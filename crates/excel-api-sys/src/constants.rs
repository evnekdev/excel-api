/// Base Excel value type mask.
pub const XLTYPE_MASK: u32 = 0x0FFF;
pub const XLTYPE_NUM: u32 = 0x0001;
pub const XLTYPE_STR: u32 = 0x0002;
pub const XLTYPE_BOOL: u32 = 0x0004;
pub const XLTYPE_REF: u32 = 0x0008;
pub const XLTYPE_ERR: u32 = 0x0010;
pub const XLTYPE_FLOW: u32 = 0x0020;
pub const XLTYPE_MULTI: u32 = 0x0040;
pub const XLTYPE_MISSING: u32 = 0x0080;
pub const XLTYPE_NIL: u32 = 0x0100;
pub const XLTYPE_SREF: u32 = 0x0400;
pub const XLTYPE_INT: u32 = 0x0800;
pub const XLTYPE_BIG_DATA: u32 = XLTYPE_STR | XLTYPE_INT;
pub const XLBIT_XL_FREE: u32 = 0x1000;
pub const XLBIT_DLL_FREE: u32 = 0x4000;
pub const XLRET_SUCCESS: i32 = 0;
pub const XLRET_ABORT: i32 = 1;
pub const XLRET_INV_XLFN: i32 = 2;
pub const XLRET_INV_COUNT: i32 = 4;
pub const XLRET_INV_XLOPER: i32 = 8;
pub const XLRET_STACK_OVFL: i32 = 16;
pub const XLRET_FAILED: i32 = 32;
pub const XLRET_UNCALCED: i32 = 64;
pub const XLRET_NOT_THREAD_SAFE: i32 = 128;
pub const XLRET_INV_ASYNC_CONTEXT: i32 = 256;
pub const XLRET_NOT_CLUSTER_SAFE: i32 = 512;
pub const EXCEL12_MAX_ROWS: u32 = 1_048_576;
pub const EXCEL12_MAX_COLUMNS: u32 = 16_384;
pub const EXCEL12_MAX_STRING_CODE_UNITS: usize = 32_767;

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn constants_match_xlcall_header() {
        assert_eq!(XLTYPE_REF, 0x0008);
        assert_eq!(XLTYPE_INT, 0x0800);
        assert_eq!(XLTYPE_BIG_DATA, XLTYPE_STR | XLTYPE_INT);
        assert_eq!(XLBIT_XL_FREE & XLTYPE_MASK, 0);
        assert_eq!(XLBIT_DLL_FREE & XLTYPE_MASK, 0);
    }
}
