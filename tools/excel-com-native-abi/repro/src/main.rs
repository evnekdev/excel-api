//! Smallest high-level windows crate reproduction retained by Prompt 05E.
//! It activates Excel, gets Workbooks, invokes Add with a canonical empty
//! DISPPARAMS, then requests Quit. It contains no Excel-specific vtable.

use std::mem::ManuallyDrop;
use windows::Win32::System::Com::{
    CLSCTX_LOCAL_SERVER, COINIT_APARTMENTTHREADED, CoCreateInstance, CoInitializeEx,
    CoUninitialize, DISPATCH_METHOD, DISPATCH_PROPERTYGET, DISPPARAMS, EXCEPINFO, IDispatch,
};
use windows::Win32::System::Variant::{VARIANT, VT_DISPATCH, VariantClear, VariantInit};
use windows::core::{GUID, HSTRING, PCWSTR};

fn status(result: windows::core::Result<()>) -> i32 {
    result.map(|_| 0).unwrap_or_else(|error| error.code().0)
}

fn main() -> Result<(), windows::core::Error> {
    unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok()? };
    let prog_id = HSTRING::from("Excel.Application");
    let clsid = unsafe { windows::Win32::System::Com::CLSIDFromProgID(&prog_id)? };
    let app: IDispatch = unsafe { CoCreateInstance(&clsid, None, CLSCTX_LOCAL_SERVER)? };
    let empty = DISPPARAMS::default();
    let workbooks = get_dispatch(&app, "Workbooks", DISPATCH_PROPERTYGET, &empty);
    let add = workbooks
        .as_ref()
        .map_or(-1, |value| invoke(value, "Add", DISPATCH_METHOD, &empty));
    let quit = invoke(&app, "Quit", DISPATCH_METHOD, &empty);
    println!(
        "{{\"windows\":\"0.62.2\",\"windows-core\":\"0.62.2\",\"windows-result\":\"0.4.1\",\"windows-strings\":\"0.5.1\",\"windows-sys\":\"0.61.2\",\"workbooks_hresult\":{},\"add_hresult\":{},\"quit_hresult\":{}}}",
        workbooks.as_ref().map_or(-1, |_| 0),
        add,
        quit
    );
    drop(app);
    unsafe { CoUninitialize() };
    Ok(())
}

fn get_dispatch(
    dispatch: &IDispatch,
    name: &str,
    flags: windows::Win32::System::Com::DISPATCH_FLAGS,
    params: &DISPPARAMS,
) -> Result<IDispatch, i32> {
    let mut result = empty_variant();
    let mut exception = EXCEPINFO::default();
    let mut arg_error = u32::MAX;
    let code = invoke_into(
        dispatch,
        name,
        flags,
        params,
        &mut result,
        &mut exception,
        &mut arg_error,
    );
    if code != 0 {
        clear_exception(&mut exception);
        clear_variant(&mut result);
        return Err(code);
    }
    let value = if vartype(&result) == VT_DISPATCH.0 {
        unsafe {
            let pointer = &result.Anonymous.Anonymous.Anonymous.pdispVal
                as *const ManuallyDrop<Option<IDispatch>>
                as *const Option<IDispatch>;
            (*pointer).clone().ok_or(0x8000_4005_u32 as i32)
        }
    } else {
        Err(0x8000_4005_u32 as i32)
    };
    clear_exception(&mut exception);
    clear_variant(&mut result);
    value
}

fn invoke(
    dispatch: &IDispatch,
    name: &str,
    flags: windows::Win32::System::Com::DISPATCH_FLAGS,
    params: &DISPPARAMS,
) -> i32 {
    let mut result = empty_variant();
    let mut exception = EXCEPINFO::default();
    let mut arg_error = u32::MAX;
    let code = invoke_into(
        dispatch,
        name,
        flags,
        params,
        &mut result,
        &mut exception,
        &mut arg_error,
    );
    clear_exception(&mut exception);
    clear_variant(&mut result);
    code
}

fn invoke_into(
    dispatch: &IDispatch,
    name: &str,
    flags: windows::Win32::System::Com::DISPATCH_FLAGS,
    params: &DISPPARAMS,
    result: &mut VARIANT,
    exception: &mut EXCEPINFO,
    arg_error: &mut u32,
) -> i32 {
    let name = HSTRING::from(name);
    let names = [PCWSTR(name.as_ptr())];
    let mut dispid = 0;
    let named = unsafe {
        dispatch.GetIDsOfNames(&GUID::from_u128(0), names.as_ptr(), 1, 0x0400, &mut dispid)
    };
    let Ok(()) = named else {
        return named.unwrap_err().code().0;
    };
    status(unsafe {
        dispatch.Invoke(
            dispid,
            &GUID::from_u128(0),
            0x0400,
            flags,
            params,
            Some(result),
            Some(exception),
            Some(arg_error),
        )
    })
}

fn empty_variant() -> VARIANT {
    unsafe { VariantInit() }
}
fn vartype(value: &VARIANT) -> u16 {
    unsafe { value.Anonymous.Anonymous.vt.0 }
}
fn clear_variant(value: &mut VARIANT) {
    let _ = unsafe { VariantClear(value) };
}
fn clear_exception(value: &mut EXCEPINFO) {
    unsafe {
        if let Some(fill) = value.pfnDeferredFillIn {
            let _ = fill(value);
        }
        let _ = ManuallyDrop::take(&mut value.bstrSource);
        let _ = ManuallyDrop::take(&mut value.bstrDescription);
        let _ = ManuallyDrop::take(&mut value.bstrHelpFile);
    }
}
