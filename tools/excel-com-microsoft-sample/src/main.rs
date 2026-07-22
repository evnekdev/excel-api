//! Independent `windows-sys` port of Microsoft's raw Excel C++ Automation sample.
//!
//! This executable deliberately does not import the repository's existing raw
//! kernel, semantic codec, dispatch frames, or retry policy.  Its baseline
//! instead follows the source recorded in `microsoft-cpp-port` field for field.

use std::collections::BTreeMap;
use std::env;
use std::ffi::c_void;
use std::fs;
use std::path::{Path, PathBuf};
use std::process::Command;
use std::thread;
use std::time::Duration;

use serde::Serialize;
use serde_json::{Value, json};
use windows_sys::Win32::Foundation::{
    CloseHandle, INVALID_HANDLE_VALUE, SysAllocString, SysFreeString,
};
use windows_sys::Win32::Globalization::{LOCALE_SYSTEM_DEFAULT, LOCALE_USER_DEFAULT};
use windows_sys::Win32::System::Com::{
    CLSCTX_LOCAL_SERVER, CLSIDFromProgID, COINIT_APARTMENTTHREADED, CoCreateInstance, CoInitialize,
    CoInitializeEx, CoUninitialize, DISPATCH_METHOD, DISPATCH_PROPERTYGET, DISPATCH_PROPERTYPUT,
    DISPPARAMS, EXCEPINFO, SAFEARRAYBOUND,
};
use windows_sys::Win32::System::Diagnostics::ToolHelp::{
    CreateToolhelp32Snapshot, PROCESSENTRY32W, Process32FirstW, Process32NextW, TH32CS_SNAPPROCESS,
};
use windows_sys::Win32::System::Ole::{SafeArrayCreate, SafeArrayPutElement};
use windows_sys::Win32::System::Variant::{
    VARIANT, VT_ARRAY, VT_BSTR, VT_DISPATCH, VT_EMPTY, VT_I4, VT_R8, VT_VARIANT, VariantClear,
    VariantInit,
};
use windows_sys::core::{GUID, HRESULT, IUnknown_Vtbl};

const IID_IDISPATCH: GUID = GUID::from_u128(0x0002_0400_0000_0000_c000_0000_0000_0046);
const IID_NULL: GUID = GUID::from_u128(0);
const DISPID_PROPERTYPUT: i32 = -3;
const DISP_E_TYPEMISMATCH: HRESULT = 0x8002_0005_u32 as i32;
const E_POINTER: HRESULT = 0x8000_4003_u32 as i32;
const E_OUTOFMEMORY: HRESULT = 0x8007_000e_u32 as i32;

#[repr(C)]
struct IDispatchVtbl {
    base: IUnknown_Vtbl,
    get_type_info_count: unsafe extern "system" fn(*mut c_void, *mut u32) -> HRESULT,
    get_type_info: unsafe extern "system" fn(*mut c_void, u32, u32, *mut *mut c_void) -> HRESULT,
    get_ids_of_names: unsafe extern "system" fn(
        *mut c_void,
        *const GUID,
        *const *const u16,
        u32,
        u32,
        *mut i32,
    ) -> HRESULT,
    invoke: unsafe extern "system" fn(
        *mut c_void,
        i32,
        *const GUID,
        u32,
        u16,
        *const DISPPARAMS,
        *mut VARIANT,
        *mut EXCEPINFO,
        *mut u32,
    ) -> HRESULT,
}

struct Dispatch {
    raw: *mut c_void,
}

impl Dispatch {
    unsafe fn from_owned(raw: *mut c_void) -> Result<Self, Failure> {
        (!raw.is_null())
            .then_some(Self { raw })
            .ok_or(Failure::new("interface extraction", E_POINTER))
    }

    unsafe fn vtbl(&self) -> &IDispatchVtbl {
        unsafe { &**(self.raw as *const *const IDispatchVtbl) }
    }
}

impl Drop for Dispatch {
    fn drop(&mut self) {
        if !self.raw.is_null() {
            unsafe {
                let _ = (self.vtbl().base.Release)(self.raw);
            }
            self.raw = std::ptr::null_mut();
        }
    }
}

struct VariantOwner {
    value: VARIANT,
}

impl VariantOwner {
    fn empty() -> Self {
        let mut value = VARIANT::default();
        unsafe { VariantInit(&mut value) };
        Self { value }
    }

    fn i4(value: i32) -> Self {
        let mut result = Self::empty();
        result.value.Anonymous.Anonymous.vt = VT_I4;
        result.value.Anonymous.Anonymous.Anonymous.lVal = value;
        result
    }

    fn bstr(text: &str) -> Result<Self, Failure> {
        let text: Vec<u16> = text.encode_utf16().chain(Some(0)).collect();
        let bstr = unsafe { SysAllocString(text.as_ptr()) };
        if bstr.is_null() {
            return Err(Failure::new("SysAllocString", E_OUTOFMEMORY));
        }
        let mut result = Self::empty();
        result.value.Anonymous.Anonymous.vt = VT_BSTR;
        result.value.Anonymous.Anonymous.Anonymous.bstrVal = bstr;
        Ok(result)
    }

    fn safearray_variant() -> Result<Self, Failure> {
        let bounds = [
            SAFEARRAYBOUND {
                lLbound: 1,
                cElements: 15,
            },
            SAFEARRAYBOUND {
                lLbound: 1,
                cElements: 15,
            },
        ];
        let array = unsafe { SafeArrayCreate(VT_VARIANT, 2, bounds.as_ptr()) };
        if array.is_null() {
            return Err(Failure::new("SAFEARRAY", E_OUTOFMEMORY));
        }
        let mut result = Self::empty();
        result.value.Anonymous.Anonymous.vt = VT_ARRAY | VT_VARIANT;
        result.value.Anonymous.Anonymous.Anonymous.parray = array;
        for row in 1_i32..=15 {
            for column in 1_i32..=15 {
                let element = Self::i4(row * column);
                let indices = [row, column];
                let status = unsafe {
                    SafeArrayPutElement(
                        array,
                        indices.as_ptr(),
                        &element.value as *const VARIANT as *const c_void,
                    )
                };
                if failed(status) {
                    return Err(Failure::new("SAFEARRAY", status));
                }
            }
        }
        Ok(result)
    }

    fn vartype(&self) -> u16 {
        unsafe { self.value.Anonymous.Anonymous.vt }
    }

    fn exact_number(&self, expected: i32) -> bool {
        unsafe {
            match self.vartype() {
                VT_I4 => self.value.Anonymous.Anonymous.Anonymous.lVal == expected,
                VT_R8 => self.value.Anonymous.Anonymous.Anonymous.dblVal == f64::from(expected),
                _ => false,
            }
        }
    }

    fn take_dispatch(&mut self, stage: &'static str) -> Result<Dispatch, Failure> {
        if self.vartype() != VT_DISPATCH {
            return Err(Failure::new(stage, DISP_E_TYPEMISMATCH));
        }
        let raw: *mut c_void = unsafe { self.value.Anonymous.Anonymous.Anonymous.pdispVal.cast() };
        if raw.is_null() {
            return Err(Failure::new(stage, E_POINTER));
        }
        // The native sample assigns result.pdispVal to an IDispatch* and later
        // releases that interface once.  The Rust port transfers exactly that
        // ownership and then leaves this VARIANT empty so Drop cannot release it.
        unsafe {
            self.value.Anonymous.Anonymous.vt = VT_EMPTY;
            self.value.Anonymous.Anonymous.Anonymous.pdispVal = std::ptr::null_mut();
            Dispatch::from_owned(raw)
        }
    }
}

impl Drop for VariantOwner {
    fn drop(&mut self) {
        unsafe {
            let _ = VariantClear(&mut self.value);
        }
    }
}

#[derive(Debug, Clone, Copy)]
struct Failure {
    stage: &'static str,
    hresult: HRESULT,
}

impl Failure {
    const fn new(stage: &'static str, hresult: HRESULT) -> Self {
        Self { stage, hresult }
    }
}

#[derive(Clone, Copy)]
enum Initialization {
    CoInitialize,
    CoInitializeExApartment,
}

#[derive(Clone, Copy)]
struct Profile {
    id: &'static str,
    add_flags: u16,
    lookup_lcid: u32,
    invoke_lcid: u32,
    initialization: Initialization,
    excepinfo_output: bool,
    pu_arg_err_output: bool,
    set_visible: bool,
    pre_add_hwnd: bool,
    pre_add_version: bool,
}

impl Profile {
    fn parse(id: &str) -> Result<Self, String> {
        let baseline = Self {
            id: "baseline",
            add_flags: DISPATCH_PROPERTYGET,
            lookup_lcid: LOCALE_USER_DEFAULT,
            invoke_lcid: LOCALE_SYSTEM_DEFAULT,
            initialization: Initialization::CoInitialize,
            excepinfo_output: false,
            pu_arg_err_output: false,
            set_visible: true,
            pre_add_hwnd: false,
            pre_add_version: false,
        };
        let profile = match id {
            "baseline" => baseline,
            "observational" => Self {
                id: "observational",
                excepinfo_output: true,
                pu_arg_err_output: true,
                ..baseline
            },
            "excepinfo-output" => Self {
                id: "excepinfo-output",
                excepinfo_output: true,
                ..baseline
            },
            "puargerr-output" => Self {
                id: "puargerr-output",
                pu_arg_err_output: true,
                ..baseline
            },
            "no-visible" => Self {
                id: "no-visible",
                set_visible: false,
                ..baseline
            },
            "preadd-hwnd" => Self {
                id: "preadd-hwnd",
                pre_add_hwnd: true,
                ..baseline
            },
            "preadd-version" => Self {
                id: "preadd-version",
                pre_add_version: true,
                ..baseline
            },
            "current-preadd-sequence" => Self {
                id: "current-preadd-sequence",
                add_flags: DISPATCH_METHOD,
                invoke_lcid: LOCALE_USER_DEFAULT,
                initialization: Initialization::CoInitializeExApartment,
                excepinfo_output: true,
                pu_arg_err_output: true,
                set_visible: false,
                pre_add_hwnd: true,
                pre_add_version: true,
                ..baseline
            },
            "add-method" => Self {
                id: "add-method",
                add_flags: DISPATCH_METHOD,
                ..baseline
            },
            "add-method-propertyget" => Self {
                id: "add-method-propertyget",
                add_flags: DISPATCH_METHOD | DISPATCH_PROPERTYGET,
                ..baseline
            },
            "invoke-project-lcid" => Self {
                id: "invoke-project-lcid",
                invoke_lcid: LOCALE_USER_DEFAULT,
                ..baseline
            },
            "lookup-project-lcid" => Self {
                id: "lookup-project-lcid",
                lookup_lcid: LOCALE_USER_DEFAULT,
                ..baseline
            },
            "coinitex-apartment" => Self {
                id: "coinitex-apartment",
                initialization: Initialization::CoInitializeExApartment,
                ..baseline
            },
            _ => {
                return Err("profile must be baseline, observational, excepinfo-output, puargerr-output, no-visible, preadd-hwnd, preadd-version, current-preadd-sequence, add-method, add-method-propertyget, invoke-project-lcid, lookup-project-lcid, or coinitex-apartment".to_owned());
            }
        };
        Ok(profile)
    }
}

#[derive(Serialize)]
struct InvocationTrace {
    sequence: u32,
    member: &'static str,
    object_role: &'static str,
    get_ids_of_names_lcid: u32,
    get_ids_of_names_hresult: HRESULT,
    dispid: Option<i32>,
    invoke_lcid: u32,
    invoke_flags: u16,
    argument_count: u32,
    named_argument_count: u32,
    argument_vartypes: Vec<u16>,
    result_hresult: HRESULT,
    result_requested: bool,
    result_vartype: Option<u16>,
    excepinfo_mode: &'static str,
    excepinfo_scode: Option<i32>,
    pu_arg_err: Option<u32>,
}

struct Invocation {
    status: HRESULT,
    result: VariantOwner,
    trace: InvocationTrace,
}

fn failed(status: HRESULT) -> bool {
    status < 0
}

// The independent port deliberately exposes each AutoWrap input as a distinct
// parameter so the Microsoft call fields remain directly auditable.
#[allow(clippy::too_many_arguments)]
fn auto_wrap(
    profile: Profile,
    result_requested: bool,
    dispatch: &Dispatch,
    member: &'static str,
    object_role: &'static str,
    flags: u16,
    args_in_com_order: &mut [VariantOwner],
    traces: &[InvocationTrace],
) -> Invocation {
    let name: Vec<u16> = member.encode_utf16().chain(Some(0)).collect();
    let names = [name.as_ptr()];
    let mut dispid = 0_i32;
    let name_status = unsafe {
        (dispatch.vtbl().get_ids_of_names)(
            dispatch.raw,
            &IID_NULL,
            names.as_ptr(),
            1,
            profile.lookup_lcid,
            &mut dispid,
        )
    };
    let mut result = VariantOwner::empty();
    let argument_vartypes = args_in_com_order
        .iter()
        .map(VariantOwner::vartype)
        .collect();
    let mut trace = InvocationTrace {
        sequence: (traces.len() + 1) as u32,
        member,
        object_role,
        get_ids_of_names_lcid: profile.lookup_lcid,
        get_ids_of_names_hresult: name_status,
        dispid: (!failed(name_status)).then_some(dispid),
        invoke_lcid: profile.invoke_lcid,
        invoke_flags: flags,
        argument_count: args_in_com_order.len() as u32,
        named_argument_count: u32::from(flags & DISPATCH_PROPERTYPUT != 0),
        argument_vartypes,
        result_hresult: name_status,
        result_requested,
        result_vartype: None,
        excepinfo_mode: match (profile.excepinfo_output, profile.pu_arg_err_output) {
            (false, false) => "official-null",
            (true, false) => "observational-excepinfo",
            (false, true) => "observational-puargerr",
            (true, true) => "observational-both",
        },
        excepinfo_scode: None,
        pu_arg_err: None,
    };
    if failed(name_status) {
        return Invocation {
            status: name_status,
            result,
            trace,
        };
    }

    let mut named = DISPID_PROPERTYPUT;
    let params = DISPPARAMS {
        rgvarg: if args_in_com_order.is_empty() {
            std::ptr::null_mut()
        } else {
            args_in_com_order.as_mut_ptr().cast::<VARIANT>()
        },
        rgdispidNamedArgs: if flags & DISPATCH_PROPERTYPUT != 0 {
            &mut named
        } else {
            std::ptr::null_mut()
        },
        cArgs: args_in_com_order.len() as u32,
        cNamedArgs: u32::from(flags & DISPATCH_PROPERTYPUT != 0),
    };
    let mut exception = EXCEPINFO::default();
    let mut arg_error = u32::MAX;
    let status = unsafe {
        (dispatch.vtbl().invoke)(
            dispatch.raw,
            dispid,
            &IID_NULL,
            profile.invoke_lcid,
            flags,
            &params,
            if result_requested {
                &mut result.value
            } else {
                std::ptr::null_mut()
            },
            if profile.excepinfo_output {
                &mut exception
            } else {
                std::ptr::null_mut()
            },
            if profile.pu_arg_err_output {
                &mut arg_error
            } else {
                std::ptr::null_mut()
            },
        )
    };
    trace.result_hresult = status;
    trace.result_vartype = result_requested.then(|| result.vartype());
    if profile.excepinfo_output {
        trace.excepinfo_scode = Some(exception.scode);
        trace.pu_arg_err = (arg_error != u32::MAX).then_some(arg_error);
        clear_excepinfo(&mut exception);
    }
    Invocation {
        status,
        result,
        trace,
    }
}

fn clear_excepinfo(info: &mut EXCEPINFO) {
    unsafe {
        if let Some(fill) = info.pfnDeferredFillIn {
            let _ = fill(info);
        }
        for value in [
            &mut info.bstrSource,
            &mut info.bstrDescription,
            &mut info.bstrHelpFile,
        ] {
            if !(*value).is_null() {
                SysFreeString(*value);
                *value = std::ptr::null();
            }
        }
    }
}

// Kept parallel to AutoWrap rather than hidden in a project-level dispatch
// abstraction; this is an isolated source-faithful control.
#[allow(clippy::too_many_arguments)]
fn call(
    profile: Profile,
    dispatch: &Dispatch,
    member: &'static str,
    object_role: &'static str,
    flags: u16,
    args_in_com_order: &mut [VariantOwner],
    result_requested: bool,
    traces: &mut Vec<InvocationTrace>,
) -> Result<VariantOwner, Failure> {
    let Invocation {
        status,
        result,
        trace,
    } = auto_wrap(
        profile,
        result_requested,
        dispatch,
        member,
        object_role,
        flags,
        args_in_com_order,
        traces,
    );
    traces.push(trace);
    if failed(status) {
        return Err(Failure::new(member, status));
    }
    Ok(result)
}

fn get_dispatch(
    profile: Profile,
    dispatch: &Dispatch,
    member: &'static str,
    object_role: &'static str,
    flags: u16,
    args_in_com_order: &mut [VariantOwner],
    traces: &mut Vec<InvocationTrace>,
) -> Result<Dispatch, Failure> {
    let mut result = call(
        profile,
        dispatch,
        member,
        object_role,
        flags,
        args_in_com_order,
        true,
        traces,
    )?;
    result.take_dispatch(member)
}

fn activate() -> Result<Dispatch, Failure> {
    let progid: Vec<u16> = "Excel.Application".encode_utf16().chain(Some(0)).collect();
    let mut clsid = GUID::default();
    let status = unsafe { CLSIDFromProgID(progid.as_ptr(), &mut clsid) };
    if failed(status) {
        return Err(Failure::new("CLSIDFromProgID", status));
    }
    let mut raw = std::ptr::null_mut();
    let status = unsafe {
        CoCreateInstance(
            &clsid,
            std::ptr::null_mut(),
            CLSCTX_LOCAL_SERVER,
            &IID_IDISPATCH,
            &mut raw,
        )
    };
    if failed(status) {
        return Err(Failure::new("CoCreateInstance", status));
    }
    unsafe { Dispatch::from_owned(raw) }
}

fn read_cell(
    profile: Profile,
    sheet: &Dispatch,
    address: &str,
    expected: i32,
    traces: &mut Vec<InvocationTrace>,
) -> Result<bool, Failure> {
    let mut address = [VariantOwner::bstr(address)?];
    let range = get_dispatch(
        profile,
        sheet,
        "Range",
        "worksheet",
        DISPATCH_PROPERTYGET,
        &mut address,
        traces,
    )?;
    let mut no_args = [];
    let value = call(
        profile,
        &range,
        "Value",
        "range",
        DISPATCH_PROPERTYGET,
        &mut no_args,
        true,
        traces,
    )?;
    Ok(value.exact_number(expected))
}

fn excel_process_count() -> u32 {
    let snapshot = unsafe { CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0) };
    if snapshot == INVALID_HANDLE_VALUE {
        return 0;
    }
    let mut entry = PROCESSENTRY32W {
        dwSize: std::mem::size_of::<PROCESSENTRY32W>() as u32,
        ..Default::default()
    };
    let mut count = 0;
    let mut more = unsafe { Process32FirstW(snapshot, &mut entry) } != 0;
    while more {
        let end = entry
            .szExeFile
            .iter()
            .position(|character| *character == 0)
            .unwrap_or(entry.szExeFile.len());
        if String::from_utf16_lossy(&entry.szExeFile[..end]).eq_ignore_ascii_case("EXCEL.EXE") {
            count += 1;
        }
        more = unsafe { Process32NextW(snapshot, &mut entry) } != 0;
    }
    unsafe {
        let _ = CloseHandle(snapshot);
    }
    count
}

fn wait_for_zero_excel_processes() -> bool {
    for _ in 0..200 {
        if excel_process_count() == 0 {
            return true;
        }
        thread::sleep(Duration::from_millis(50));
    }
    excel_process_count() == 0
}

fn run_profile(profile: Profile) -> Value {
    let preexisting = excel_process_count();
    if preexisting != 0 {
        return json!({
            "control": "faithful-rust",
            "profile": profile.id,
            "classification": "activation failure",
            "failure_stage": "preexisting-excel",
            "failure_hresult": 0x8007_00AA_u32 as i32,
            "preexisting_excel_process_count": preexisting,
            "trace": [],
            "cleanup": {"owned_process_exit_verified": false, "forced_termination": false},
            "raw_pointer_values_recorded": false,
        });
    }

    let initialization_status = unsafe {
        match profile.initialization {
            Initialization::CoInitialize => CoInitialize(std::ptr::null()),
            Initialization::CoInitializeExApartment => {
                CoInitializeEx(std::ptr::null(), COINIT_APARTMENTTHREADED as u32)
            }
        }
    };
    if failed(initialization_status) {
        return json!({
            "control": "faithful-rust",
            "profile": profile.id,
            "classification": "activation failure",
            "failure_stage": "CoInitialize",
            "failure_hresult": initialization_status,
            "preexisting_excel_process_count": preexisting,
            "trace": [],
            "cleanup": {"owned_process_exit_verified": false, "forced_termination": false},
            "raw_pointer_values_recorded": false,
        });
    }

    let mut traces = Vec::new();
    let mut app = None;
    let mut books = None;
    let mut book = None;
    let mut sheet = None;
    let mut range = None;
    let mut array = None;
    let mut readback = [false; 3];
    let mut quit_requested = false;
    let result: Result<(), Failure> = (|| {
        app = Some(activate()?);
        let application = app.as_ref().expect("application was set");
        if profile.set_visible {
            let mut visible = [VariantOwner::i4(1)];
            let _ = call(
                profile,
                application,
                "Visible",
                "application",
                DISPATCH_PROPERTYPUT,
                &mut visible,
                false,
                &mut traces,
            )?;
        }
        let mut no_args = [];
        if profile.pre_add_hwnd {
            let _ = call(
                profile,
                application,
                "Hwnd",
                "application",
                DISPATCH_PROPERTYGET,
                &mut no_args,
                true,
                &mut traces,
            )?;
        }
        if profile.pre_add_version {
            let _ = call(
                profile,
                application,
                "Version",
                "application",
                DISPATCH_PROPERTYGET,
                &mut no_args,
                true,
                &mut traces,
            )?;
        }
        books = Some(get_dispatch(
            profile,
            application,
            "Workbooks",
            "application",
            DISPATCH_PROPERTYGET,
            &mut no_args,
            &mut traces,
        )?);
        book = Some(get_dispatch(
            profile,
            books.as_ref().expect("workbooks was set"),
            "Add",
            "workbooks",
            profile.add_flags,
            &mut no_args,
            &mut traces,
        )?);
        array = Some(VariantOwner::safearray_variant()?);
        sheet = Some(get_dispatch(
            profile,
            application,
            "ActiveSheet",
            "application",
            DISPATCH_PROPERTYGET,
            &mut no_args,
            &mut traces,
        )?);
        let mut address = [VariantOwner::bstr("A1:O15")?];
        range = Some(get_dispatch(
            profile,
            sheet.as_ref().expect("sheet was set"),
            "Range",
            "worksheet",
            DISPATCH_PROPERTYGET,
            &mut address,
            &mut traces,
        )?);
        let array_value = array.take().expect("array was initialized");
        let mut value_argument = [array_value];
        let _ = call(
            profile,
            range.as_ref().expect("range was set"),
            "Value",
            "range",
            DISPATCH_PROPERTYPUT,
            &mut value_argument,
            false,
            &mut traces,
        )?;
        // Restore the owning value after the Invoke so it remains alive until
        // the original sample's final VariantClear-equivalent cleanup point.
        array = Some(
            value_argument
                .into_iter()
                .next()
                .expect("one array argument"),
        );
        let active_sheet = sheet.as_ref().expect("sheet was set");
        readback[0] = read_cell(profile, active_sheet, "A1", 1, &mut traces)?;
        readback[1] = read_cell(profile, active_sheet, "B3", 6, &mut traces)?;
        readback[2] = read_cell(profile, active_sheet, "O15", 225, &mut traces)?;
        if !readback.iter().all(|value| *value) {
            return Err(Failure::new("read-back", DISP_E_TYPEMISMATCH));
        }
        let mut saved = [VariantOwner::i4(1)];
        let _ = call(
            profile,
            book.as_ref().expect("book was set"),
            "Saved",
            "workbook",
            DISPATCH_PROPERTYPUT,
            &mut saved,
            false,
            &mut traces,
        )?;
        let _ = call(
            profile,
            application,
            "Quit",
            "application",
            DISPATCH_METHOD,
            &mut no_args,
            false,
            &mut traces,
        )?;
        quit_requested = true;
        Ok(())
    })();

    let mut final_failure = result.err();
    if !quit_requested && let Some(application) = app.as_ref() {
        let mut no_args = [];
        let quit = call(
            profile,
            application,
            "Quit",
            "application",
            DISPATCH_METHOD,
            &mut no_args,
            false,
            &mut traces,
        );
        if let Err(failure) = quit
            && final_failure.is_none()
        {
            final_failure = Some(failure);
        }
        quit_requested = final_failure.is_none();
    }
    // Preserve the exact successful-sample release order.
    drop(range.take());
    drop(sheet.take());
    drop(book.take());
    drop(books.take());
    drop(app.take());
    drop(array.take());
    unsafe { CoUninitialize() };
    let process_exited = wait_for_zero_excel_processes();
    if !process_exited && final_failure.is_none() {
        final_failure = Some(Failure::new("cleanup", E_POINTER));
    }
    let classification = if final_failure.is_none() && readback.iter().all(|value| *value) {
        "complete"
    } else {
        "failure"
    };
    json!({
        "control": "faithful-rust",
        "profile": profile.id,
        "classification": classification,
        "failure_stage": final_failure.map(|failure| failure.stage).unwrap_or(""),
        "failure_hresult": final_failure.map(|failure| failure.hresult).unwrap_or(0),
        "preexisting_excel_process_count": preexisting,
        "initialization": match profile.initialization { Initialization::CoInitialize => "CoInitialize(NULL)", Initialization::CoInitializeExApartment => "CoInitializeEx(COINIT_APARTMENTTHREADED)" },
        "clsctx": "CLSCTX_LOCAL_SERVER",
        "requested_iid": "IID_IDispatch",
        "visible_set": profile.set_visible,
        "readback": {"A1": readback[0], "B3": readback[1], "O15": readback[2]},
        "trace": traces,
        "cleanup": {"quit_requested": quit_requested, "owned_process_exit_verified": process_exited, "forced_termination": false},
        "raw_pointer_values_recorded": false,
    })
}

fn options(arguments: &[String]) -> Result<BTreeMap<String, String>, String> {
    let mut values = BTreeMap::new();
    let mut index = 0;
    while index < arguments.len() {
        let key = arguments[index]
            .strip_prefix("--")
            .ok_or_else(usage)?
            .to_owned();
        let value = arguments.get(index + 1).ok_or_else(usage)?.to_owned();
        if values.insert(key.clone(), value).is_some() {
            return Err(format!("duplicate option --{key}"));
        }
        index += 2;
    }
    Ok(values)
}

fn root(options: &BTreeMap<String, String>) -> Result<PathBuf, String> {
    options.get("root").map(PathBuf::from).ok_or_else(usage)
}

fn write_text(path: &Path, text: &str) -> Result<(), String> {
    if let Some(parent) = path.parent() {
        fs::create_dir_all(parent).map_err(|error| error.to_string())?;
    }
    fs::write(path, text).map_err(|error| error.to_string())
}

fn read_jsonl(path: &Path) -> Result<Vec<Value>, String> {
    if !path.exists() {
        return Ok(Vec::new());
    }
    let text = fs::read_to_string(path).map_err(|error| error.to_string())?;
    text.lines()
        .map(|line| serde_json::from_str(line).map_err(|error| error.to_string()))
        .collect()
}

fn merge_rows(path: &Path, fresh: Vec<Value>) -> Result<(), String> {
    let mut rows = BTreeMap::new();
    for value in read_jsonl(path)?.into_iter().chain(fresh) {
        let id = value
            .get("id")
            .and_then(Value::as_str)
            .ok_or_else(|| format!("{} has a row without id", path.display()))?
            .to_owned();
        rows.insert(id, value);
    }
    let text = rows
        .into_values()
        .map(|value| serde_json::to_string(&value).map_err(|error| error.to_string()))
        .collect::<Result<Vec<_>, _>>()?
        .join("\n")
        + "\n";
    write_text(path, &text)
}

fn initialize(root: &Path) -> Result<(), String> {
    let source = root.join("microsoft-cpp-port");
    fs::create_dir_all(&source).map_err(|error| error.to_string())?;
    write_text(
        &source.join("SOURCE_MANIFEST.toml"),
        "schema_version = 1\nname = \"microsoft-cpp-port\"\nclassification = \"research-only\"\nsource = \"Microsoft Learn previous-versions article; independently written C++ and Rust controls\"\n",
    )?;
    merge_rows(
        &source.join("documentation-sources.jsonl"),
        vec![json!({
            "schema_version": 1,
            "id": "microsoft.learn.automate-excel-from-c",
            "title": "How to automate Excel from C++ without using MFC or #import",
            "url": "https://learn.microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/automate-excel-from-c",
            "access_date": "2026-07-22",
            "publisher": "Microsoft",
            "sample_language": "C++",
            "sample_age_status": "Microsoft Learn previous-versions article; original copyright notice 1999; page last updated 2021-10-22",
            "license_or_reuse_statement": "(c) Microsoft Corporation 1999, All Rights Reserved. Contributions by Joe Crump, Microsoft Corporation.",
            "local_download_sha256": "71e1575d2a43930eb941138cad2c3ebb0b4f52b54b9316293636758a5dac51c0",
            "downloaded_artifact_committed": false,
            "raw_pointer_values_recorded": false,
        })],
    )?;
    merge_rows(
        &source.join("unresolved.jsonl"),
        vec![json!({
            "schema_version": 1,
            "id": "unresolved.microsoft-cpp-port.live-baseline",
            "classification": "Not tested",
            "detail": "C++ and faithful Rust 20-process baselines have not been captured yet.",
            "raw_pointer_values_recorded": false,
        })],
    )?;
    refresh_reports(root)
}

fn summary(rows: &[Value]) -> (usize, usize, Vec<String>) {
    let complete = rows
        .iter()
        .filter(|row| row.get("classification") == Some(&Value::String("complete".to_owned())))
        .count();
    let failures = rows
        .iter()
        .filter_map(|row| row.get("failure_stage").and_then(Value::as_str))
        .filter(|stage| !stage.is_empty())
        .map(ToOwned::to_owned)
        .collect();
    (rows.len(), complete, failures)
}

fn records_raw_pointer_values(value: &Value) -> bool {
    value
        .get("raw_pointer_values_recorded")
        .and_then(Value::as_bool)
        == Some(true)
}

fn refresh_reports(root: &Path) -> Result<(), String> {
    let source = root.join("microsoft-cpp-port");
    let generated = root.join("generated/microsoft-cpp-port");
    fs::create_dir_all(&generated).map_err(|error| error.to_string())?;
    let cpp = read_jsonl(&source.join("official-cpp-runs.jsonl"))?;
    let rust = read_jsonl(&source.join("faithful-rust-runs.jsonl"))?;
    let differentials = read_jsonl(&source.join("controlled-differentials.jsonl"))?;
    let (cpp_runs, cpp_complete, cpp_failures) = summary(&cpp);
    let (rust_runs, rust_complete, rust_failures) = summary(&rust);
    let reports = [
        ("official-sample-summary.md", "# Official Microsoft sample summary\n\nThe source is Microsoft Learn's previous-versions C++ sample. Its required baseline uses `CoInitialize(NULL)`, `CLSCTX_LOCAL_SERVER`, `LOCALE_USER_DEFAULT` for `GetIDsOfNames`, `LOCALE_SYSTEM_DEFAULT` for `Invoke`, and `DISPATCH_PROPERTYGET` for `Workbooks.Add`.\n".to_owned()),
        ("cpp-baseline.md", format!("# C++ baseline\n\nCaptured runs: {cpp_runs}. Complete runs: {cpp_complete}. Failure stages: {}.\n", if cpp_failures.is_empty() { "none".to_owned() } else { cpp_failures.join(", ") })),
        ("faithful-rust-baseline.md", format!("# Faithful Rust baseline\n\nCaptured runs: {rust_runs}. Complete runs: {rust_complete}. Failure stages: {}.\n", if rust_failures.is_empty() { "none".to_owned() } else { rust_failures.join(", ") })),
        ("call-field-comparison.md", "# Call-field comparison\n\nThe C++ and Rust controls record initialization, activation, lookup/invocation LCIDs, flags, DISPPARAMS counts, result VARTYPEs, exception-output mode, and cleanup without recording pointer values.\n".to_owned()),
        ("add-invocation-classification.md", format!("# Workbooks.Add invocation classification\n\nThe faithful baseline uses `DISPATCH_PROPERTYGET`. Controlled differential rows captured: {}.\n", differentials.len())),
        ("locale-comparison.md", "# Locale comparison\n\nThe faithful baseline keeps the Microsoft asymmetry: `LOCALE_USER_DEFAULT` (1024) for name lookup and `LOCALE_SYSTEM_DEFAULT` (2048) for invocation.\n".to_owned()),
        ("initialization-comparison.md", "# Initialization comparison\n\nThe faithful baseline calls `CoInitialize(NULL)`; `CoInitializeEx(COINIT_APARTMENTTHREADED)` is only a post-baseline one-variable differential.\n".to_owned()),
        ("adopted-corrections.md", "# Adopted corrections\n\nNo current-kernel correction is recorded until an independent C++ and faithful Rust baseline establishes it.\n".to_owned()),
        ("semantic-live-results.md", "# Semantic live results\n\nThe Prompt 06 semantic live gate is rerun only after a confirmed transport correction.\n".to_owned()),
        ("remaining-blockers.md", "# Remaining blockers\n\nDo not treat raw external Excel Automation as production-ready until the independent controls and corrected semantic live gate complete reliably.\n".to_owned()),
    ];
    for (name, content) in reports {
        write_text(&generated.join(name), &content)?;
    }
    Ok(())
}

fn check(root: &Path) -> Result<(), String> {
    let source = root.join("microsoft-cpp-port");
    let required = [
        "documentation-sources.jsonl",
        "environment.jsonl",
        "official-cpp-runs.jsonl",
        "faithful-rust-runs.jsonl",
        "call-field-comparison.jsonl",
        "current-kernel-differences.jsonl",
        "controlled-differentials.jsonl",
        "adopted-corrections.jsonl",
        "semantic-live-observations.jsonl",
        "unresolved.jsonl",
    ];
    for name in required {
        let path = source.join(name);
        let text = fs::read_to_string(&path).map_err(|error| error.to_string())?;
        if text.contains('\r') || !text.ends_with('\n') {
            return Err(format!(
                "{} must use LF and a final newline",
                path.display()
            ));
        }
        for line in text.lines() {
            let value: Value = serde_json::from_str(line).map_err(|error| error.to_string())?;
            if records_raw_pointer_values(&value) {
                return Err(format!("{} records raw pointer values", path.display()));
            }
        }
    }
    for name in [
        "official-sample-summary.md",
        "cpp-baseline.md",
        "faithful-rust-baseline.md",
        "call-field-comparison.md",
        "add-invocation-classification.md",
        "locale-comparison.md",
        "initialization-comparison.md",
        "adopted-corrections.md",
        "semantic-live-results.md",
        "remaining-blockers.md",
    ] {
        let path = root.join("generated/microsoft-cpp-port").join(name);
        let text = fs::read_to_string(&path).map_err(|error| error.to_string())?;
        if text.contains('\r') || !text.ends_with('\n') || !text.starts_with("# ") {
            return Err(format!(
                "{} is not a valid generated report",
                path.display()
            ));
        }
    }
    Ok(())
}

fn capture_repeat(
    root: &Path,
    control: &str,
    profile: Option<&str>,
    executable: Option<&Path>,
    count: usize,
) -> Result<(), String> {
    let source = root.join("microsoft-cpp-port");
    fs::create_dir_all(&source).map_err(|error| error.to_string())?;
    let current = env::current_exe().map_err(|error| error.to_string())?;
    let mut rows = Vec::new();
    for run in 1..=count {
        let preexisting = excel_process_count();
        let mut record = if preexisting == 0 {
            let mut command = if control == "faithful-rust" {
                let mut command = Command::new(&current);
                command
                    .arg("rust-child")
                    .arg("--profile")
                    .arg(profile.unwrap_or("baseline"));
                command
            } else {
                Command::new(executable.ok_or("cpp-repeat requires --exe")?)
            };
            let output = command.output().map_err(|error| error.to_string())?;
            let stdout = String::from_utf8(output.stdout).map_err(|error| error.to_string())?;
            serde_json::from_str::<Value>(stdout.trim()).unwrap_or_else(|_| {
                json!({
                    "control": control,
                    "classification": "cleanup failure",
                    "failure_stage": "control-output",
                    "failure_hresult": 0x8000_4005_u32 as i32,
                    "trace": [],
                    "cleanup": {"owned_process_exit_verified": false, "forced_termination": false},
                    "raw_pointer_values_recorded": false,
                })
            })
        } else {
            json!({
                "control": control,
                "classification": "activation failure",
                "failure_stage": "preexisting-excel",
                "failure_hresult": 0x8007_00AA_u32 as i32,
                "trace": [],
                "cleanup": {"owned_process_exit_verified": false, "forced_termination": false},
                "raw_pointer_values_recorded": false,
            })
        };
        let process_exited = excel_process_count() == 0;
        record["id"] = Value::String(format!(
            "{}.{}.run-{:02}",
            if control == "official-cpp" {
                "official-cpp"
            } else {
                "faithful-rust"
            },
            profile.unwrap_or("baseline"),
            run
        ));
        record["schema_version"] = Value::from(1);
        record["fresh_process"] = Value::Bool(true);
        record["parent_post_run_excel_process_count"] =
            Value::from(if process_exited { 0 } else { 1 });
        record["raw_pointer_values_recorded"] = Value::Bool(false);
        rows.push(record);
        if !process_exited {
            // No process is terminated. Remaining runs are still recorded as
            // blocked by the required zero-pre-existing-process gate.
        }
    }
    let filename = if control == "official-cpp" {
        "official-cpp-runs.jsonl"
    } else {
        "faithful-rust-runs.jsonl"
    };
    merge_rows(&source.join(filename), rows)?;
    refresh_reports(root)
}

fn capture_differential(root: &Path, profile: &str, count: usize) -> Result<(), String> {
    let parsed = Profile::parse(profile)?;
    let current = env::current_exe().map_err(|error| error.to_string())?;
    let mut rows = Vec::new();
    for run in 1..=count {
        if excel_process_count() != 0 {
            rows.push(json!({
                "schema_version": 1,
                "id": format!("differential.{}.run-{:02}", profile, run),
                "profile": profile,
                "classification": "activation failure",
                "failure_stage": "preexisting-excel",
                "raw_pointer_values_recorded": false,
            }));
            continue;
        }
        let output = Command::new(&current)
            .arg("rust-child")
            .arg("--profile")
            .arg(profile)
            .output()
            .map_err(|error| error.to_string())?;
        let mut row: Value =
            serde_json::from_slice(&output.stdout).map_err(|error| error.to_string())?;
        row["id"] = Value::String(format!("differential.{}.run-{:02}", profile, run));
        row["schema_version"] = Value::from(1);
        row["baseline"] = Value::String("baseline".to_owned());
        row["one_variable_change"] = Value::String(
            match parsed.id {
                "add-method" | "add-method-propertyget" => "Workbooks.Add invocation flags",
                "invoke-project-lcid" => "Invoke LCID",
                "lookup-project-lcid" => "GetIDsOfNames LCID",
                "coinitex-apartment" => "COM initialization function",
                "observational" => "EXCEPINFO and puArgErr output pointers",
                "excepinfo-output" => "EXCEPINFO output pointer",
                "puargerr-output" => "puArgErr output pointer",
                "no-visible" => "Visible property assignment",
                "preadd-hwnd" => "pre-Add Hwnd property get",
                "preadd-version" => "pre-Add Version property get",
                "current-preadd-sequence" => {
                    "the existing kernel's combined pre-Add sequence (comparison only)"
                }
                _ => "none",
            }
            .to_owned(),
        );
        row["raw_pointer_values_recorded"] = Value::Bool(false);
        rows.push(row);
    }
    merge_rows(
        &root.join("microsoft-cpp-port/controlled-differentials.jsonl"),
        rows,
    )?;
    refresh_reports(root)
}

fn capture_semantic_live(root: &Path, probe: &Path) -> Result<(), String> {
    if excel_process_count() != 0 {
        return Err(
            "semantic live capture requires zero pre-existing EXCEL.EXE processes".to_owned(),
        );
    }
    let output = Command::new(probe)
        .arg("automation-value-live-observe")
        .arg("--root")
        .arg(root)
        .output()
        .map_err(|error| error.to_string())?;
    if !output.status.success() {
        return Err(format!(
            "semantic live observer failed: {}",
            String::from_utf8_lossy(&output.stderr).trim()
        ));
    }
    let report: Value =
        serde_json::from_slice(&output.stdout).map_err(|error| error.to_string())?;
    if report.get("prompt_06_evidence_modified") != Some(&Value::Bool(false)) {
        return Err(
            "semantic live observer did not confirm Prompt 06 evidence isolation".to_owned(),
        );
    }
    let mut rows = report
        .get("rows")
        .and_then(Value::as_array)
        .cloned()
        .ok_or("semantic live observer returned no rows")?;
    for row in &mut rows {
        row["environment_id"] = Value::String("post-microsoft-cpp-port-controls".to_owned());
        row["prompt_06_evidence_modified"] = Value::Bool(false);
        row["raw_pointer_values_recorded"] = Value::Bool(false);
    }
    merge_rows(
        &root.join("microsoft-cpp-port/semantic-live-observations.jsonl"),
        rows,
    )?;
    refresh_reports(root)?;
    let failures = report
        .get("required_failures")
        .and_then(Value::as_array)
        .ok_or("semantic live observer omitted required failures")?;
    if failures.is_empty() && excel_process_count() == 0 {
        Ok(())
    } else if !failures.is_empty() {
        Err(format!(
            "semantic live required failures: {}",
            failures
                .iter()
                .filter_map(Value::as_str)
                .collect::<Vec<_>>()
                .join(", ")
        ))
    } else {
        Err("semantic live cleanup left an Excel process; no process was terminated".to_owned())
    }
}

fn usage() -> String {
    "usage: excel-com-microsoft-sample <init|check|refresh|rust-child|rust-repeat|cpp-repeat|differential|semantic-live> --root <knowledge-root> [--profile <profile>] [--count <count>] [--exe <native-cpp-control>] [--probe <range-probe-executable>]".to_owned()
}

fn run(arguments: Vec<String>) -> Result<(), String> {
    let command = arguments.first().map(String::as_str).ok_or_else(usage)?;
    let options = options(&arguments[1..])?;
    match command {
        "init" => initialize(&root(&options)?),
        "check" => check(&root(&options)?),
        "refresh" => refresh_reports(&root(&options)?),
        "rust-child" => {
            let profile = Profile::parse(
                options
                    .get("profile")
                    .map(String::as_str)
                    .unwrap_or("baseline"),
            )?;
            println!(
                "{}",
                serde_json::to_string(&run_profile(profile)).map_err(|error| error.to_string())?
            );
            Ok(())
        }
        "rust-repeat" => {
            let count = options
                .get("count")
                .map(String::as_str)
                .unwrap_or("20")
                .parse()
                .map_err(|_| "--count must be an integer".to_owned())?;
            capture_repeat(
                &root(&options)?,
                "faithful-rust",
                options.get("profile").map(String::as_str),
                None,
                count,
            )
        }
        "cpp-repeat" => {
            let count = options
                .get("count")
                .map(String::as_str)
                .unwrap_or("20")
                .parse()
                .map_err(|_| "--count must be an integer".to_owned())?;
            let executable = options
                .get("exe")
                .map(PathBuf::from)
                .ok_or("cpp-repeat requires --exe")?;
            capture_repeat(
                &root(&options)?,
                "official-cpp",
                Some("baseline"),
                Some(&executable),
                count,
            )
        }
        "differential" => {
            let count = options
                .get("count")
                .map(String::as_str)
                .unwrap_or("1")
                .parse()
                .map_err(|_| "--count must be an integer".to_owned())?;
            let profile = options
                .get("profile")
                .ok_or("differential requires --profile")?;
            capture_differential(&root(&options)?, profile, count)
        }
        "semantic-live" => {
            let probe = options
                .get("probe")
                .map(PathBuf::from)
                .ok_or("semantic-live requires --probe")?;
            capture_semantic_live(&root(&options)?, &probe)
        }
        _ => Err(usage()),
    }
}

fn main() {
    if let Err(error) = run(env::args().skip(1).collect()) {
        eprintln!("excel-com-microsoft-sample: {error}");
        std::process::exit(1);
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn baseline_preserves_microsoft_call_fields() {
        let profile = Profile::parse("baseline").expect("baseline");
        assert_eq!(profile.add_flags, DISPATCH_PROPERTYGET);
        assert_eq!(profile.lookup_lcid, LOCALE_USER_DEFAULT);
        assert_eq!(profile.invoke_lcid, LOCALE_SYSTEM_DEFAULT);
        assert!(matches!(
            profile.initialization,
            Initialization::CoInitialize
        ));
        assert!(!profile.excepinfo_output);
        assert!(!profile.pu_arg_err_output);
        assert!(profile.set_visible);
    }

    #[test]
    fn controlled_profiles_change_one_documented_field() {
        assert_eq!(
            Profile::parse("add-method").expect("method").add_flags,
            DISPATCH_METHOD
        );
        assert_eq!(
            Profile::parse("add-method-propertyget")
                .expect("both")
                .add_flags,
            DISPATCH_METHOD | DISPATCH_PROPERTYGET
        );
        assert!(Profile::parse("invalid").is_err());
    }

    #[test]
    fn zero_argument_frame_is_null_by_construction() {
        let mut no_args: [VariantOwner; 0] = [];
        let frame = DISPPARAMS {
            rgvarg: if no_args.is_empty() {
                std::ptr::null_mut()
            } else {
                no_args.as_mut_ptr().cast()
            },
            rgdispidNamedArgs: std::ptr::null_mut(),
            cArgs: no_args.len() as u32,
            cNamedArgs: 0,
        };
        assert!(frame.rgvarg.is_null());
        assert!(frame.rgdispidNamedArgs.is_null());
        assert_eq!(frame.cArgs, 0);
    }

    #[test]
    fn property_put_uses_the_standard_named_argument() {
        let mut named = DISPID_PROPERTYPUT;
        let frame = DISPPARAMS {
            rgvarg: std::ptr::null_mut(),
            rgdispidNamedArgs: &mut named,
            cArgs: 1,
            cNamedArgs: 1,
        };
        assert_eq!(unsafe { *frame.rgdispidNamedArgs }, -3);
    }

    #[test]
    fn source_provenance_hash_is_stable() {
        assert_eq!(
            "71e1575d2a43930eb941138cad2c3ebb0b4f52b54b9316293636758a5dac51c0".len(),
            64
        );
    }

    #[test]
    fn redaction_flag_is_parsed_as_a_boolean() {
        assert!(!records_raw_pointer_values(&json!({
            "raw_pointer_values_recorded": false,
        })));
        assert!(records_raw_pointer_values(&json!({
            "raw_pointer_values_recorded": true,
        })));
    }
}
