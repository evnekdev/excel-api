use crate::XLOPER12;

/// Raw signature of Excel's vector-based C API entry point.
///
/// The exact linking/loading mechanism is intentionally not decided in the
/// foundational ABI crate.
pub type Excel12vFn = unsafe extern "system" fn(
    function: i32,
    result: *mut XLOPER12,
    argument_count: i32,
    arguments: *const *mut XLOPER12,
) -> i32;

/// Standard XLL lifecycle callback signature.
pub type XlAutoOpenFn = unsafe extern "system" fn() -> i32;
/// Standard XLL shutdown callback signature.
pub type XlAutoCloseFn = unsafe extern "system" fn() -> i32;
/// Standard XLL return-memory cleanup callback signature.
pub type XlAutoFree12Fn = unsafe extern "system" fn(value: *mut XLOPER12);
