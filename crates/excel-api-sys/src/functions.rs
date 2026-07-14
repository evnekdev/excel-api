use core::ffi::c_int;

use crate::{LPXLOPER12, XLOPER12};

/// Raw signature of the variadic Excel 12 C API entry point (`_cdecl`).
pub type Excel12Fn =
    unsafe extern "C" fn(xlfn: c_int, operRes: LPXLOPER12, count: c_int, ...) -> c_int;

/// Raw signature of the vector Excel 12 C API entry point (`pascal`/WinAPI).
pub type Excel12vFn = unsafe extern "system" fn(
    xlfn: c_int,
    operRes: LPXLOPER12,
    count: c_int,
    opers: *mut LPXLOPER12,
) -> c_int;

/// Signature of Excel's internal `MdCallBack12` entry point used by
/// `XLCALL.CPP`.
pub type Excel12EntryPtFn = unsafe extern "system" fn(
    xlfn: c_int,
    coper: c_int,
    rgpxloper12: *mut LPXLOPER12,
    xloper12Res: LPXLOPER12,
) -> c_int;

/// Signature of the SDK bridge export used by cluster hosts.
pub type SetExcel12EntryPtFn = unsafe extern "system" fn(callback: Excel12EntryPtFn);

/// Raw signature of `XLCallVer`.
pub type XLCallVerFn = unsafe extern "system" fn() -> c_int;

/// Cluster connector asynchronous callback signature from `XLCALL.H`.
pub type PXL_HPC_ASYNC_CALLBACK =
    unsafe extern "system" fn(operAsyncHandle: LPXLOPER12, operReturn: LPXLOPER12) -> c_int;

/// Standard XLL initialization callback.
pub type XlAutoOpenFn = unsafe extern "system" fn() -> c_int;
/// Standard XLL shutdown callback.
pub type XlAutoCloseFn = unsafe extern "system" fn() -> c_int;
/// Add-in Manager installation callback.
pub type XlAutoAddFn = unsafe extern "system" fn() -> c_int;
/// Add-in Manager removal callback.
pub type XlAutoRemoveFn = unsafe extern "system" fn() -> c_int;
/// Add-in Manager information callback for the Excel 12 ABI.
pub type XlAddInManagerInfo12Fn = unsafe extern "system" fn(action: LPXLOPER12) -> LPXLOPER12;
/// On-demand registration callback for the Excel 12 ABI.
pub type XlAutoRegister12Fn = unsafe extern "system" fn(name: LPXLOPER12) -> LPXLOPER12;
/// XLL-owned return-memory cleanup callback.
pub type XlAutoFree12Fn = unsafe extern "system" fn(value: LPXLOPER12);

// Ensure the aliases remain pointer-compatible with the concrete raw type.
const _: Option<LPXLOPER12> = None::<*mut XLOPER12>;
