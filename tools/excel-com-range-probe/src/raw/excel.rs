//! Excel member names and workflow built only on the generic Automation layer.

use std::path::Path;

use serde_json::{json, Value};
use windows_sys::core::GUID;
use windows_sys::Win32::System::Com::{
    CLSIDFromProgID, CoCreateInstance, CoCreateInstanceEx, CLSCTX_LOCAL_SERVER, CLSCTX_SERVER,
    MULTI_QI,
};

use super::apartment::ComApartment;
use super::com_ptr::{ComPtr, Dispatch};
use super::dispatch::{call, Call, Frame, DISPATCH_METHOD, DISPATCH_PROPERTYGET, DISPATCH_PROPERTYPUT};
use super::process::OwnedProcess;
use super::variant::OwnedVariant;
use super::{Mode, Value as RootValue};

const IID_IDISPATCH: GUID = GUID::from_u128(0x00020400_0000_0000_c000_000000000046);

pub(super) fn activate(mode: Mode) -> Result<ComPtr<Dispatch>, String> {
    let program: Vec<u16> = "Excel.Application".encode_utf16().chain(Some(0)).collect();
    let mut class_id = GUID::default();
    let status = unsafe { CLSIDFromProgID(program.as_ptr(), &mut class_id) };
    if status != 0 {
        return Err(format!("CLSIDFromProgID failed: {}", super::hex(status)));
    }
    let mut raw = std::ptr::null_mut();
    let status = match mode {
        Mode::L | Mode::S => unsafe {
            CoCreateInstance(
                &class_id,
                std::ptr::null_mut(),
                if matches!(mode, Mode::L) { CLSCTX_LOCAL_SERVER } else { CLSCTX_SERVER },
                &IID_IDISPATCH,
                &mut raw,
            )
        },
        Mode::X => {
            let mut request = MULTI_QI {
                pIID: &IID_IDISPATCH,
                pItf: std::ptr::null_mut(),
                hr: 0,
            };
            let outer = unsafe {
                CoCreateInstanceEx(
                    &class_id,
                    std::ptr::null_mut(),
                    CLSCTX_SERVER,
                    std::ptr::null(),
                    1,
                    &mut request,
                )
            };
            raw = request.pItf;
            if outer == 0 { request.hr } else { outer }
        }
    };
    if status != 0 {
        return Err(format!("activation failed: {}", super::hex(status)));
    }
    unsafe { ComPtr::from_owned(raw) }
        .ok_or_else(|| "activation returned a null IDispatch pointer".to_owned())
}

pub(super) fn get(
    target: &ComPtr<Dispatch>,
    name: &str,
    lcid: u32,
    args: Vec<OwnedVariant>,
) -> Result<ComPtr<Dispatch>, String> {
    let result = call(
        target,
        name,
        DISPATCH_PROPERTYGET,
        lcid,
        Frame::positional(args),
    );
    if result.hr != 0 {
        return Err(format!("{name} failed: {}", super::hex(result.hr)));
    }
    result
        .result
        .dispatch()
        .ok_or_else(|| format!("{name} did not return VT_DISPATCH"))
}

pub(super) fn brief(call: &Call) -> RootValue {
    let argument_error_applies =
        matches!(call.hr, value if value == 0x8002_0004_u32 as i32 || value == 0x8002_0005_u32 as i32);
    json!({
        "hresult": call.hr,
        "hresult_hex": super::hex(call.hr),
        "result_vartype": call.result.vt(),
        "excepinfo": call.exception,
        "pu_arg_err": if argument_error_applies && call.arg_error != u32::MAX {
            RootValue::from(call.arg_error)
        } else {
            RootValue::String("not-applicable".to_owned())
        },
    })
}

/// Runs the existing research smoke without defining an Excel-specific vtable.
pub(super) fn run(mode: Mode, fixture: &Path) -> Result<Value, String> {
    let _apartment = ComApartment::initialize()?;
    let app = activate(mode)?;
    let lcid = if matches!(mode, Mode::L) { 0x0400 } else { 0 };
    let owned = OwnedProcess::from_app(&app, lcid)?;
    let version = call(&app, "Version", DISPATCH_PROPERTYGET, lcid, Frame::empty());
    let workbooks = get(&app, "Workbooks", lcid, vec![])?;
    let add = call(&workbooks, "Add", DISPATCH_METHOD, lcid, Frame::empty());
    let add_report = brief(&add);
    if add.hr == 0 && let Some(book) = add.result.dispatch() {
        let _ = call(
            &book,
            "Close",
            DISPATCH_METHOD,
            lcid,
            Frame::positional(vec![OwnedVariant::boolean(false)]),
        );
    }
    let open = call(
        &workbooks,
        "Open",
        DISPATCH_METHOD,
        lcid,
        Frame::positional(vec![OwnedVariant::bstr(&fixture.to_string_lossy())?]),
    );
    let open_report = brief(&open);
    let mut smoke = json!({"entered": false});
    let mut failure = None;
    let opened = if open.hr == 0 { open.result.dispatch() } else { None };
    if let Some(book) = opened.as_ref() {
        match get(book, "Worksheets", lcid, vec![])
            .and_then(|sheets| get(&sheets, "Item", lcid, vec![OwnedVariant::i4(1)]))
            .and_then(|sheet| get(&sheet, "Range", lcid, vec![OwnedVariant::bstr("A1")?]))
        {
            Ok(range) => {
                let write = call(
                    &range,
                    "Value2",
                    DISPATCH_PROPERTYPUT,
                    lcid,
                    Frame::put(OwnedVariant::i4(42)),
                );
                let read = call(&range, "Value2", DISPATCH_PROPERTYGET, lcid, Frame::empty());
                let clear = call(&range, "ClearContents", DISPATCH_METHOD, lcid, Frame::empty());
                let exact = read.hr == 0 && read.result.is_exact_42();
                smoke = json!({
                    "entered": true,
                    "write": brief(&write),
                    "read": brief(&read),
                    "read_value_exactly_42": exact,
                    "clear": brief(&clear),
                });
                if !exact || write.hr != 0 || clear.hr != 0 {
                    failure = Some(json!({
                        "inner_scode_hex": read.exception.get("scode").and_then(Value::as_str).unwrap_or("--"),
                        "workbook_returned": true,
                    }));
                }
            }
            Err(error) => {
                failure = Some(json!({
                    "detail": error,
                    "inner_scode_hex": "--",
                    "workbook_returned": true,
                }));
            }
        }
    } else {
        failure = Some(json!({
            "inner_scode_hex": open.exception.get("scode").and_then(Value::as_str).unwrap_or("--"),
            "workbook_returned": false,
        }));
    }
    if let Some(book) = opened {
        let _ = call(
            &book,
            "Close",
            DISPATCH_METHOD,
            lcid,
            Frame::positional(vec![OwnedVariant::boolean(false)]),
        );
    }
    let quit = call(&app, "Quit", DISPATCH_METHOD, lcid, Frame::empty());
    drop(workbooks);
    drop(app);
    let exited = owned.wait();
    let success = add_report.get("hresult") == Some(&Value::from(0))
        && open_report.get("hresult") == Some(&Value::from(0))
        && smoke.get("read_value_exactly_42") == Some(&Value::Bool(true))
        && exited;
    Ok(json!({
        "backend": "raw-windows-sys",
        "activation_mode": mode.id(),
        "activation": mode.activation(),
        "version": brief(&version),
        "workbooks_add": add_report,
        "workbooks_open": open_report,
        "range_smoke": smoke,
        "cleanup": {
            "excel_quit": brief(&quit),
            "owned_process_exit_verified": exited,
            "forced_termination": false,
        },
        "failure": failure,
        "success": success,
        "raw_pointer_values_recorded": false,
    }))
}
