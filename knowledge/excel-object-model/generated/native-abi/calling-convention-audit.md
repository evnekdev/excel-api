# Calling-convention audit

The SDK and `windows 0.62.2` generated `IDispatch_Vtbl::Invoke` use `unsafe extern "system" fn(*mut c_void, i32, *const GUID, u32, DISPATCH_FLAGS/WORD, *const DISPPARAMS, *mut VARIANT, *mut EXCEPINFO, *mut u32) -> HRESULT`. `HRESULT` and `DISPID` are signed 32-bit; `LCID` and `UINT` counts are unsigned 32-bit; flags are 16-bit. `IUnknown` precedes `GetTypeInfoCount`, `GetTypeInfo`, `GetIDsOfNames`, and `Invoke`. `CoCreateInstance` and `CoCreateInstanceEx` are generated `system` imports in both windows and windows-sys.
