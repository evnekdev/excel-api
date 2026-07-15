//! Raw constants from `XLCALL.H` and the Excel 12 registration ABI.

// XLOPER12 type tags.
pub const xltypeNum: u32 = 0x0001;
pub const xltypeStr: u32 = 0x0002;
pub const xltypeBool: u32 = 0x0004;
pub const xltypeRef: u32 = 0x0008;
pub const xltypeErr: u32 = 0x0010;
pub const xltypeFlow: u32 = 0x0020;
pub const xltypeMulti: u32 = 0x0040;
pub const xltypeMissing: u32 = 0x0080;
pub const xltypeNil: u32 = 0x0100;
pub const xltypeSRef: u32 = 0x0400;
pub const xltypeInt: u32 = 0x0800;
pub const xlbitXLFree: u32 = 0x1000;
pub const xlbitDLLFree: u32 = 0x4000;
pub const xltypeBigData: u32 = xltypeStr | xltypeInt;

/// Mask covering every base `XLOPER12` type bit, excluding ownership bits.
pub const XLTYPE_MASK: u32 = 0x0FFF;

// Excel error values.
pub const xlerrNull: i32 = 0;
pub const xlerrDiv0: i32 = 7;
pub const xlerrValue: i32 = 15;
pub const xlerrRef: i32 = 23;
pub const xlerrName: i32 = 29;
pub const xlerrNum: i32 = 36;
pub const xlerrNA: i32 = 42;
pub const xlerrGettingData: i32 = 43;

// Flow-control values.
pub const xlflowHalt: i32 = 1;
pub const xlflowGoto: i32 = 2;
pub const xlflowRestart: i32 = 8;
pub const xlflowPause: i32 = 16;
pub const xlflowResume: i32 = 64;

// Excel4/Excel12 return codes. Multiple failure bits may be combined.
pub const xlretSuccess: i32 = 0;
pub const xlretAbort: i32 = 1;
pub const xlretInvXlfn: i32 = 2;
pub const xlretInvCount: i32 = 4;
pub const xlretInvXloper: i32 = 8;
pub const xlretStackOvfl: i32 = 16;
pub const xlretFailed: i32 = 32;
pub const xlretUncalced: i32 = 64;
pub const xlretNotThreadSafe: i32 = 128;
pub const xlretInvAsynchronousContext: i32 = 256;
pub const xlretNotClusterSafe: i32 = 512;

// Cluster connector entry-point return codes.
pub const xlHpcRetSuccess: i32 = 0;
pub const xlHpcRetSessionIdInvalid: i32 = -1;
pub const xlHpcRetCallFailed: i32 = -2;

// Event identifiers.
pub const xleventCalculationEnded: i32 = 1;
pub const xleventCalculationCanceled: i32 = 2;

// Function-number category bits.
pub const xlCommand: i32 = 0x8000;
pub const xlSpecial: i32 = 0x4000;
pub const xlIntl: i32 = 0x2000;
pub const xlPrompt: i32 = 0x1000;

// C API-only function numbers used by the foundation milestones.
pub const xlFree: i32 = xlSpecial;
pub const xlStack: i32 = 1 | xlSpecial;
pub const xlCoerce: i32 = 2 | xlSpecial;
pub const xlSet: i32 = 3 | xlSpecial;
pub const xlSheetId: i32 = 4 | xlSpecial;
pub const xlSheetNm: i32 = 5 | xlSpecial;
pub const xlAbort: i32 = 6 | xlSpecial;
pub const xlGetInst: i32 = 7 | xlSpecial;
pub const xlGetHwnd: i32 = 8 | xlSpecial;
pub const xlGetName: i32 = 9 | xlSpecial;
pub const xlEnableXLMsgs: i32 = 10 | xlSpecial;
pub const xlDisableXLMsgs: i32 = 11 | xlSpecial;
pub const xlDefineBinaryName: i32 = 12 | xlSpecial;
pub const xlGetBinaryName: i32 = 13 | xlSpecial;
pub const xlAsyncReturn: i32 = 16 | xlSpecial;
pub const xlEventRegister: i32 = 17 | xlSpecial;
pub const xlRunningOnCluster: i32 = 18 | xlSpecial;
pub const xlGetInstPtr: i32 = 19 | xlSpecial;

// Worksheet/XLM function numbers required through manual registration.
/// Current date/time as an Excel serial value (`XLCALL.H`: `xlfNow = 74`).
pub const xlfNow: i32 = 74;
pub const xlfSetName: i32 = 88;
pub const xlfCaller: i32 = 89;
pub const xlfRegister: i32 = 149;
pub const xlfUnregister: i32 = 201;
pub const xlfRegisterId: i32 = 267;
pub const xlUDF: i32 = 255;

/// Historical XLM `ON.TIME` command (`XLCALL.H`: `148 | xlCommand`).
///
/// This raw constant does not by itself establish a modern supported contract.
pub const xlcOnTime: i32 = 148 | xlCommand;

// Excel 12 capacity limits. Rows and columns are counts; MAX_ROW/MAX_COLUMN
// are zero-based indices.
pub const EXCEL12_MAX_ROWS: i32 = 1_048_576;
pub const EXCEL12_MAX_COLUMNS: i32 = 16_384;
pub const EXCEL12_MAX_ROW: i32 = EXCEL12_MAX_ROWS - 1;
pub const EXCEL12_MAX_COLUMN: i32 = EXCEL12_MAX_COLUMNS - 1;
pub const EXCEL12_MAX_STRING_CODE_UNITS: usize = 32_767;
pub const EXCEL12_MAX_ARGUMENTS: i32 = 255;

// Registration data type codes used by the Excel 12 milestones.
pub const XLL_TYPE_BOOL: &str = "A";
pub const XLL_TYPE_DOUBLE: &str = "B";
pub const XLL_TYPE_XCHAR_NULL_TERMINATED: &str = "C%";
pub const XLL_TYPE_XCHAR_COUNTED: &str = "D%";
pub const XLL_TYPE_DOUBLE_POINTER: &str = "E";
pub const XLL_TYPE_XCHAR_NULL_TERMINATED_IN_PLACE: &str = "F%";
pub const XLL_TYPE_XCHAR_COUNTED_IN_PLACE: &str = "G%";
pub const XLL_TYPE_WORD: &str = "H";
pub const XLL_TYPE_I16: &str = "I";
pub const XLL_TYPE_I32: &str = "J";
pub const XLL_TYPE_FP12: &str = "K%";
pub const XLL_TYPE_BOOL_POINTER: &str = "L";
pub const XLL_TYPE_I16_POINTER: &str = "M";
pub const XLL_TYPE_I32_POINTER: &str = "N";
pub const XLL_TYPE_ARRAY12: &str = "O%";
/// Value-only `XLOPER12`; references are dereferenced before the call.
pub const XLL_TYPE_XLOPER12_VALUE: &str = "Q";
/// Reference-preserving `XLOPER12`.
pub const XLL_TYPE_XLOPER12_REFERENCE: &str = "U";
pub const XLL_TYPE_ASYNC_HANDLE: &str = "X";
/// Void return required for asynchronous UDF registration.
pub const XLL_TYPE_ASYNC_VOID: &str = ">";

// Legacy registration codes are exposed only to make the P/Q and R/U
// distinction explicit. Legacy XLOPER structures and Excel4 are unsupported.
pub const XLL_TYPE_LEGACY_XLOPER_VALUE: &str = "P";
pub const XLL_TYPE_LEGACY_XLOPER_REFERENCE: &str = "R";

pub const XLL_MODIFIER_VOLATILE: char = '!';
pub const XLL_MODIFIER_MACRO_SHEET: char = '#';
pub const XLL_MODIFIER_THREAD_SAFE: char = '$';
pub const XLL_MODIFIER_CLUSTER_SAFE: char = '&';
/// Historical void/modify-first-argument prefix; modern synchronous
/// modify-in-place registration uses a return digit from `1` through `9`.
pub const XLL_MODIFIER_VOID: char = '>';
pub const XLL_MODIFY_IN_PLACE_FIRST: char = '1';
pub const XLL_MODIFY_IN_PLACE_LAST: char = '9';

// Compatibility spellings retained from the initial scaffold.
pub const XLTYPE_NUM: u32 = xltypeNum;
pub const XLTYPE_STR: u32 = xltypeStr;
pub const XLTYPE_BOOL: u32 = xltypeBool;
pub const XLTYPE_REF: u32 = xltypeRef;
pub const XLTYPE_ERR: u32 = xltypeErr;
pub const XLTYPE_FLOW: u32 = xltypeFlow;
pub const XLTYPE_MULTI: u32 = xltypeMulti;
pub const XLTYPE_MISSING: u32 = xltypeMissing;
pub const XLTYPE_NIL: u32 = xltypeNil;
pub const XLTYPE_SREF: u32 = xltypeSRef;
pub const XLTYPE_INT: u32 = xltypeInt;
pub const XLTYPE_BIG_DATA: u32 = xltypeBigData;
pub const XLBIT_XL_FREE: u32 = xlbitXLFree;
pub const XLBIT_DLL_FREE: u32 = xlbitDLLFree;
pub const XLRET_SUCCESS: i32 = xlretSuccess;
pub const XLRET_ABORT: i32 = xlretAbort;
pub const XLRET_INV_XLFN: i32 = xlretInvXlfn;
pub const XLRET_INV_COUNT: i32 = xlretInvCount;
pub const XLRET_INV_XLOPER: i32 = xlretInvXloper;
pub const XLRET_STACK_OVFL: i32 = xlretStackOvfl;
pub const XLRET_FAILED: i32 = xlretFailed;
pub const XLRET_UNCALCED: i32 = xlretUncalced;
pub const XLRET_NOT_THREAD_SAFE: i32 = xlretNotThreadSafe;
pub const XLRET_INV_ASYNC_CONTEXT: i32 = xlretInvAsynchronousContext;
pub const XLRET_NOT_CLUSTER_SAFE: i32 = xlretNotClusterSafe;

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn type_and_ownership_bits_do_not_overlap() {
        assert_eq!(xltypeRef, 0x0008);
        assert_eq!(xltypeInt, 0x0800);
        assert_eq!(xltypeBigData, xltypeStr | xltypeInt);
        assert_eq!(xlbitXLFree & XLTYPE_MASK, 0);
        assert_eq!(xlbitDLLFree & XLTYPE_MASK, 0);
    }

    #[test]
    fn modern_general_registration_modes_remain_distinct() {
        assert_eq!(XLL_TYPE_XLOPER12_VALUE, "Q");
        assert_eq!(XLL_TYPE_XLOPER12_REFERENCE, "U");
        assert_ne!(XLL_TYPE_XLOPER12_VALUE, XLL_TYPE_XLOPER12_REFERENCE);
    }

    #[test]
    fn on_time_and_now_match_checked_in_xlcall_header() {
        assert_eq!(xlfNow, 74);
        assert_eq!(xlcOnTime, 32_916);
        assert_eq!(xlcOnTime, 0x8094);
    }

    #[test]
    fn excel12_limits_are_internally_consistent() {
        assert_eq!(EXCEL12_MAX_ROW + 1, EXCEL12_MAX_ROWS);
        assert_eq!(EXCEL12_MAX_COLUMN + 1, EXCEL12_MAX_COLUMNS);
        assert_eq!(EXCEL12_MAX_STRING_CODE_UNITS, 32_767);
        assert_eq!(EXCEL12_MAX_ARGUMENTS, 255);
    }
}
