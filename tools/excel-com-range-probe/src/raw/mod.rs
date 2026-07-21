//! Generic `windows-sys` Automation kernel used by the Prompt 05G recovery
//! probe.  This is intentionally a research-only transport layer: it declares
//! the SDK `IDispatch` vtable, never an Excel dual-interface layout.

use serde_json::{json, Value};
use std::collections::BTreeMap;
use std::fs;
use std::path::{Path, PathBuf};
use std::process::Command;

const MANIFEST: &str = "schema_version = 1\nname = \"excel-com-windows-sys-kernel\"\nclassification = \"research-only\"\nbackend_default = \"raw-windows-sys\"\nsource = \"Prompt 05E lower_level_run generic IDispatch path\"\n";

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Mode {
    L,
    S,
    X,
}

impl Mode {
    fn parse(value: &str) -> Result<Vec<Self>, String> {
        match value {
            "L" | "l" => Ok(vec![Self::L]),
            "S" | "s" => Ok(vec![Self::S]),
            "X" | "x" => Ok(vec![Self::X]),
            "all" => Ok(vec![Self::L, Self::S, Self::X]),
            _ => Err("kernel mode must be L, S, X, or all".to_owned()),
        }
    }
    fn id(self) -> &'static str {
        match self {
            Self::L => "L",
            Self::S => "S",
            Self::X => "X",
        }
    }
    fn activation(self) -> &'static str {
        match self {
            Self::L => "CoCreateInstance(CLSCTX_LOCAL_SERVER), LCID 0x0400",
            Self::S => "CoCreateInstance(CLSCTX_SERVER), LCID 0",
            Self::X => "CoCreateInstanceEx(CLSCTX_SERVER), LCID 0",
        }
    }
}

/// Creates the standalone Prompt 05G evidence tree. It is safe to call again:
/// existing runtime rows are preserved and reports are regenerated from them.
pub fn initialize(root: &Path) -> Result<(), String> {
    let mut capture = read(root).unwrap_or_else(|_| Capture::empty());
    if capture.manifest.is_empty() {
        capture.manifest = MANIFEST.to_owned();
    }
    if capture.environments.is_empty() {
        capture.environments.push(json!({
            "schema_version": 1,
            "id": "environment.05g-current-host",
            "classification": "Version-specific",
            "excel": "locally installed Automation server; exact version is read by the live kernel",
            "process_gate": "each child requires pre-existing EXCEL.EXE count = 0 and records only ownership verification, never PID/HWND/path",
            "raw_pointer_values_recorded": false
        }));
    }
    if capture.configurations.is_empty() {
        for mode in [Mode::L, Mode::S, Mode::X] {
            capture.configurations.push(json!({
                "schema_version": 1,
                "id": format!("backend.raw-windows-sys.{}", mode.id()),
                "backend": "raw-windows-sys",
                "default_for_recovery": true,
                "activation_mode": mode.id(),
                "activation": mode.activation(),
                "com_initialization": "CoInitializeEx(COINIT_APARTMENTTHREADED)",
                "dispatch": "generic SDK IDispatch vtable; no Excel-specific dual-interface declaration",
                "argument_order": "logical positional arguments are reversed into rgvarg; zero arguments use null pointers",
                "property_put": "DISPID_PROPERTYPUT with one named argument",
                "result": "VariantInit before every Invoke and VariantClear exactly once",
                "excepinfo": "deferred fill-in is called during cleanup and all returned BSTRs are freed",
                "raw_pointer_values_recorded": false
            }));
        }
        capture.configurations.push(json!({
            "schema_version": 1,
            "id": "backend.high-level-windows.diagnostic-only",
            "backend": "high-level-windows",
            "default_for_recovery": false,
            "status": "opt-in diagnostic control only",
            "raw_pointer_values_recorded": false
        }));
    }
    if capture.unresolved.is_empty() {
        capture.unresolved = vec![
            json!({"schema_version":1,"id":"unresolved.05g-runtime-matrix","classification":"Not tested","detail":"Prompt 05 scalar/rectangle Value and Value2 matrix resumes only after the required 30 fresh-process kernel runs."}),
            json!({"schema_version":1,"id":"unresolved.05g-high-level-comparison","classification":"Not tested","detail":"Clean/cold/warm/current-session high-level comparison is opt-in and must not replace raw kernel evidence."}),
        ];
    }
    write(root, &capture)
}

/// Executes one bounded experiment, a 30-child repeatability matrix, a
/// comparison control, or the narrow startup-retry policy. `raw-windows-sys`
/// is the default backend; `high-level-windows` is explicitly diagnostic.
pub fn run(
    root: &Path,
    backend: &str,
    mode: &str,
    fixture: Option<&Path>,
    action: &str,
) -> Result<String, String> {
    initialize(root)?;
    let fixture = fixture
        .map(Path::to_path_buf)
        .unwrap_or_else(default_fixture);
    if !fixture.is_file() {
        return Err(format!(
            "controlled fixture is unavailable: {}",
            fixture.display()
        ));
    }
    match action {
        "child" => {
            let mode = Mode::parse(mode)?
                .into_iter()
                .next()
                .ok_or("missing child mode")?;
            let result = execute(backend, mode, &fixture, root.parent().unwrap_or(root))?;
            serde_json::to_string(&result).map_err(|error| error.to_string())
        }
        "single" => {
            let modes = Mode::parse(mode)?;
            let mut count = 0;
            for selected in modes {
                persist_run(root, child(root, backend, selected, &fixture)?, "single")?;
                count += 1;
            }
            Ok(format!("persisted {count} fresh-process kernel run(s)"))
        }
        "repeatability" => {
            if backend != "raw-windows-sys" {
                return Err(
                    "the required repeatability matrix is limited to raw-windows-sys".to_owned(),
                );
            }
            let modes = Mode::parse(mode)?;
            if modes.len() != 3 {
                return Err("repeatability must use --mode all for exactly 10 L, 10 S, and 10 X fresh processes".to_owned());
            }
            let mut rows = Vec::new();
            for selected in modes {
                for iteration in 1..=10 {
                    let result = child(root, backend, selected, &fixture)?;
                    let success = result.get("success").and_then(Value::as_bool) == Some(true);
                    rows.push(json!({
                        "schema_version": 1,
                        "id": format!("repeatability.05g.{}.run-{:02}", selected.id(), iteration),
                        "backend": backend,
                        "activation_mode": selected.id(),
                        "iteration": iteration,
                        "fresh_process": true,
                        "complete_sequence_success": success,
                        "result": result,
                        "classification": if success { "Runtime-observed" } else { "Inconclusive" },
                        "raw_pointer_values_recorded": false
                    }));
                }
            }
            let mut capture = read(root)?;
            merge(&mut capture.repeatability, rows);
            write(root, &capture)?;
            let capture = read(root)?;
            let successful = capture
                .repeatability
                .iter()
                .filter(|row| row.get("complete_sequence_success") == Some(&Value::Bool(true)))
                .count();
            if successful != 30 {
                return Err(format!("repeatability is incomplete: {successful}/30 successful complete sequences; failures were preserved"));
            }
            Ok(
                "recorded exactly 30/30 successful raw-windows-sys fresh-process sequences"
                    .to_owned(),
            )
        }
        "compare" => {
            let selected = Mode::parse(mode)?.into_iter().next().unwrap_or(Mode::L);
            let raw = child(root, "raw-windows-sys", selected, &fixture)?;
            let high = child(root, "high-level-windows", selected, &fixture)?;
            let mut capture = read(root)?;
            merge(
                &mut capture.comparisons,
                vec![json!({
                    "schema_version":1,
                    "id":format!("backend-comparison.05g.current-session.{}", selected.id()),
                    "context":"current-session",
                    "activation_mode":selected.id(),
                    "raw_windows_sys":raw,
                    "high_level_windows":high,
                    "classification":"Runtime-observed",
                    "raw_pointer_values_recorded":false
                })],
            );
            write(root, &capture)?;
            Ok("recorded current-session raw/high-level backend comparison; cold/clean/warm rows remain explicitly unresolved".to_owned())
        }
        "retry" => {
            let selected = Mode::parse(mode)?.into_iter().next().unwrap_or(Mode::L);
            let first = child(root, backend, selected, &fixture)?;
            let retry = retry_eligible(&first);
            let second = if retry {
                Some(child(root, backend, selected, &fixture)?)
            } else {
                None
            };
            let mut capture = read(root)?;
            merge(
                &mut capture.retries,
                vec![json!({
                    "schema_version":1,
                    "id":format!("startup-retry.05g.{}", selected.id()),
                    "backend":backend,
                    "activation_mode":selected.id(),
                    "first_attempt":first,
                    "eligible":retry,
                    "retry_attempted":second.is_some(),
                    "retry":second,
                    "policy":"only Add/Open failure before a workbook is returned with EXCEPINFO scode 0x800A03EC may trigger exactly one new-instance retry",
                    "classification":"Runtime-observed",
                    "raw_pointer_values_recorded":false
                })],
            );
            write(root, &capture)?;
            Ok("recorded bounded startup-retry policy result".to_owned())
        }
        _ => {
            Err("kernel action must be single, repeatability, compare, retry, or child".to_owned())
        }
    }
}

pub fn check(root: &Path) -> Result<(), String> {
    let capture = read(root)?;
    for (path, expected) in artifacts(&capture)? {
        let actual = fs::read_to_string(root.join(&path))
            .map_err(|error| format!("cannot read {}: {error}", root.join(&path).display()))?;
        if actual != expected {
            return Err(format!(
                "kernel artifact {} is stale; rerun kernel-init or the relevant kernel action",
                path.display()
            ));
        }
        if actual.contains("\r\n") || !actual.ends_with('\n') || actual.contains("ptr=") {
            return Err(format!(
                "kernel artifact {} violates persistence rules",
                path.display()
            ));
        }
    }
    Ok(())
}

fn default_fixture() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("testdata/controlled-fixture.csv")
}

fn child(root: &Path, backend: &str, mode: Mode, fixture: &Path) -> Result<Value, String> {
    let count = excel_process_count()?;
    if count != 0 {
        return Err(format!(
            "fresh-process gate refused to run: pre-existing EXCEL.EXE process count = {count}"
        ));
    }
    let executable = std::env::current_exe().map_err(|error| error.to_string())?;
    let output = Command::new(executable)
        .args(["kernel", "--root"])
        .arg(root)
        .args(["--backend", backend, "--mode", mode.id(), "--fixture"])
        .arg(fixture)
        .args(["--action", "child"])
        .output()
        .map_err(|error| format!("cannot start fresh kernel child: {error}"))?;
    let text = String::from_utf8(output.stdout).map_err(|error| error.to_string())?;
    if !output.status.success() {
        return Err(format!(
            "kernel child failed: {}",
            String::from_utf8_lossy(&output.stderr).trim()
        ));
    }
    let result: Value = serde_json::from_str(text.trim())
        .map_err(|error| format!("kernel child emitted invalid JSON: {error}"))?;
    let post = excel_process_count()?;
    if post != 0 {
        return Err(format!("owned child cleanup did not restore EXCEL.EXE count to zero (found {post}); no process was terminated"));
    }
    Ok(result)
}

fn excel_process_count() -> Result<usize, String> {
    let output = Command::new("powershell.exe")
        .args([
            "-NoProfile",
            "-Command",
            "@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count",
        ])
        .output()
        .map_err(|error| format!("cannot query EXCEL.EXE process count: {error}"))?;
    if !output.status.success() {
        return Err("EXCEL.EXE process gate command failed".to_owned());
    }
    String::from_utf8_lossy(&output.stdout)
        .trim()
        .parse::<usize>()
        .map_err(|error| format!("invalid EXCEL.EXE process count: {error}"))
}

fn persist_run(root: &Path, result: Value, context: &str) -> Result<(), String> {
    let mode = result
        .get("activation_mode")
        .and_then(Value::as_str)
        .unwrap_or("unknown");
    let backend = result
        .get("backend")
        .and_then(Value::as_str)
        .unwrap_or("unknown");
    let mut capture = read(root)?;
    merge(
        &mut capture.workbooks,
        vec![json!({
            "schema_version":1,"id":format!("workbook.05g.{context}.{backend}.{mode}"),"backend":backend,"activation_mode":mode,
            "add":result.get("workbooks_add").cloned().unwrap_or(Value::Null),"open":result.get("workbooks_open").cloned().unwrap_or(Value::Null),
            "cleanup":result.get("cleanup").cloned().unwrap_or(Value::Null),"classification":"Runtime-observed","raw_pointer_values_recorded":false
        })],
    );
    merge(
        &mut capture.smoke,
        vec![json!({
            "schema_version":1,"id":format!("range-smoke.05g.{context}.{backend}.{mode}"),"backend":backend,"activation_mode":mode,
            "smoke":result.get("range_smoke").cloned().unwrap_or(Value::Null),"success":result.get("success").cloned().unwrap_or(Value::Bool(false)),"classification":"Runtime-observed","raw_pointer_values_recorded":false
        })],
    );
    write(root, &capture)
}

fn retry_eligible(result: &Value) -> bool {
    let failed = ["workbooks_add", "workbooks_open"].into_iter().any(|name| {
        result
            .get(name)
            .and_then(|row| row.get("hresult"))
            .and_then(Value::as_i64)
            .is_some_and(|hr| hr != 0)
    });
    failed
        && result
            .pointer("/failure/inner_scode_hex")
            .and_then(Value::as_str)
            == Some("0x800A03EC")
        && result
            .pointer("/failure/workbook_returned")
            .and_then(Value::as_bool)
            == Some(false)
}

#[derive(Clone)]
struct Capture {
    manifest: String,
    environments: Vec<Value>,
    configurations: Vec<Value>,
    workbooks: Vec<Value>,
    smoke: Vec<Value>,
    repeatability: Vec<Value>,
    retries: Vec<Value>,
    comparisons: Vec<Value>,
    unresolved: Vec<Value>,
}
impl Capture {
    fn empty() -> Self {
        Self {
            manifest: String::new(),
            environments: vec![],
            configurations: vec![],
            workbooks: vec![],
            smoke: vec![],
            repeatability: vec![],
            retries: vec![],
            comparisons: vec![],
            unresolved: vec![],
        }
    }
}
fn read(root: &Path) -> Result<Capture, String> {
    Ok(Capture {
        manifest: fs::read_to_string(root.join("SOURCE_MANIFEST.toml"))
            .map_err(|e| e.to_string())?,
        environments: read_jsonl(&root.join("environments.jsonl"))?,
        configurations: read_jsonl(&root.join("backend-configurations.jsonl"))?,
        workbooks: read_jsonl(&root.join("workbook-observations.jsonl"))?,
        smoke: read_jsonl(&root.join("range-smoke-observations.jsonl"))?,
        repeatability: read_jsonl(&root.join("repeatability.jsonl"))?,
        retries: read_jsonl(&root.join("startup-retry.jsonl"))?,
        comparisons: read_jsonl(&root.join("backend-comparisons.jsonl"))?,
        unresolved: read_jsonl(&root.join("unresolved.jsonl"))?,
    })
}
fn read_jsonl(path: &Path) -> Result<Vec<Value>, String> {
    let text = fs::read_to_string(path).map_err(|e| e.to_string())?;
    text.lines()
        .filter(|line| !line.is_empty())
        .map(|line| serde_json::from_str(line).map_err(|e| e.to_string()))
        .collect()
}
fn merge(target: &mut Vec<Value>, incoming: Vec<Value>) {
    let mut all = BTreeMap::new();
    for row in std::mem::take(target).into_iter().chain(incoming) {
        all.insert(
            row.get("id")
                .and_then(Value::as_str)
                .unwrap_or("missing")
                .to_owned(),
            row,
        );
    }
    *target = all.into_values().collect();
}
fn jsonl(rows: &[Value]) -> Result<String, String> {
    let mut rows = rows.to_vec();
    rows.sort_by_key(|row| {
        row.get("id")
            .and_then(Value::as_str)
            .unwrap_or("")
            .to_owned()
    });
    let mut text = String::new();
    for row in rows {
        text.push_str(&serde_json::to_string(&row).map_err(|e| e.to_string())?);
        text.push('\n');
    }
    if text.is_empty() {
        text.push('\n');
    }
    Ok(text)
}
fn artifacts(c: &Capture) -> Result<BTreeMap<PathBuf, String>, String> {
    let mut files = BTreeMap::new();
    files.insert(PathBuf::from("SOURCE_MANIFEST.toml"), c.manifest.clone());
    for (name, rows) in [
        ("environments.jsonl", &c.environments),
        ("backend-configurations.jsonl", &c.configurations),
        ("workbook-observations.jsonl", &c.workbooks),
        ("range-smoke-observations.jsonl", &c.smoke),
        ("repeatability.jsonl", &c.repeatability),
        ("startup-retry.jsonl", &c.retries),
        ("backend-comparisons.jsonl", &c.comparisons),
        ("unresolved.jsonl", &c.unresolved),
    ] {
        files.insert(PathBuf::from(name), jsonl(rows)?);
    }
    for (name, body) in reports(c) {
        files.insert(
            PathBuf::from("../generated/windows-sys-kernel").join(name),
            body,
        );
    }
    Ok(files)
}
fn write(root: &Path, c: &Capture) -> Result<(), String> {
    for (path, text) in artifacts(c)? {
        let full = root.join(path);
        let parent = full.parent().ok_or("artifact path has no parent")?;
        fs::create_dir_all(parent).map_err(|e| e.to_string())?;
        fs::write(full, text).map_err(|e| e.to_string())?;
    }
    Ok(())
}
fn report(title: &str, rows: &[Value]) -> String {
    let mut text=format!("# {title}\n\nGenerated by `excel-com-range-probe`; do not edit by hand. Pointer values, HWNDs, PIDs, and local fixture paths are never persisted.\n\n| ID | Classification | Summary |\n| --- | --- | --- |\n");
    if rows.is_empty() {
        text.push_str("| -- | Not tested | No row captured. |\n");
    } else {
        for row in rows {
            text.push_str(&format!(
                "| `{}` | {} | `{}` |\n",
                row.get("id").and_then(Value::as_str).unwrap_or("--"),
                row.get("classification")
                    .and_then(Value::as_str)
                    .unwrap_or("Runtime-observed"),
                serde_json::to_string(row)
                    .unwrap_or_else(|_| "{}".to_owned())
                    .replace('|', "\\|")
            ));
        }
    }
    text
}
fn reports(c: &Capture) -> BTreeMap<&'static str, String> {
    BTreeMap::from([
        (
            "kernel-design.md",
            report("Windows-sys Automation kernel", &c.configurations),
        ),
        (
            "backend-comparison.md",
            report("Backend comparison", &c.comparisons),
        ),
        (
            "workbook-add-results.md",
            report("Workbook Add results", &c.workbooks),
        ),
        (
            "workbook-open-results.md",
            report("Workbook Open results", &c.workbooks),
        ),
        (
            "range-smoke-results.md",
            report("Range smoke results", &c.smoke),
        ),
        (
            "repeatability.md",
            report("Fresh-process repeatability", &c.repeatability),
        ),
        (
            "startup-retry.md",
            report("Startup retry policy", &c.retries),
        ),
        ("cleanup.md", report("Owned Excel cleanup", &c.workbooks)),
        (
            "prompt-05-resumption.md",
            report("Prompt 05 resumption status", &c.unresolved),
        ),
        (
            "remaining-blockers.md",
            report("Remaining blockers", &c.unresolved),
        ),
    ])
}

#[cfg(windows)]
fn execute(
    backend: &str,
    mode: Mode,
    fixture: &Path,
    knowledge_root: &Path,
) -> Result<Value, String> {
    match backend {
        "raw-windows-sys" => windows::run(mode, fixture),
        "high-level-windows" => super::high_level_kernel_control(
            knowledge_root,
            Some(fixture),
            "05g-high-level-control",
        ),
        _ => Err("kernel backend must be raw-windows-sys or high-level-windows".to_owned()),
    }
}
#[cfg(not(windows))]
fn execute(
    _backend: &str,
    _mode: Mode,
    _fixture: &Path,
    _knowledge_root: &Path,
) -> Result<Value, String> {
    Err("windows-sys kernel requires Windows and locally installed Excel".to_owned())
}

#[cfg(windows)]
mod windows {
    use super::{json, Mode, Value};
    use std::ffi::c_void;
    use std::marker::PhantomData;
    use std::path::Path;
    use windows_sys::core::{IUnknown_Vtbl, GUID, HRESULT};
    use windows_sys::Win32::Foundation::{
        CloseHandle, SysAllocString, SysFreeString, FILETIME, HANDLE, HWND, WAIT_OBJECT_0,
    };
    use windows_sys::Win32::System::Com::{
        CLSIDFromProgID, CoCreateInstance, CoCreateInstanceEx, CoInitializeEx, CoUninitialize,
        CLSCTX_LOCAL_SERVER, CLSCTX_SERVER, COINIT_APARTMENTTHREADED, DISPPARAMS, EXCEPINFO,
        MULTI_QI,
    };
    use windows_sys::Win32::System::Threading::{
        GetProcessTimes, OpenProcess, WaitForSingleObject, PROCESS_QUERY_LIMITED_INFORMATION,
        PROCESS_SYNCHRONIZE,
    };
    use windows_sys::Win32::System::Variant::{
        VariantClear, VariantInit, VARIANT, VT_BOOL, VT_BSTR, VT_DISPATCH, VT_I4, VT_R8,
    };
    use windows_sys::Win32::UI::WindowsAndMessaging::GetWindowThreadProcessId;

    const IID_IDISPATCH: GUID = GUID::from_u128(0x00020400_0000_0000_c000_000000000046);
    const DISPATCH_METHOD: u16 = 1;
    const DISPATCH_PROPERTYGET: u16 = 2;
    const DISPATCH_PROPERTYPUT: u16 = 4;
    const DISPID_PROPERTYPUT: i32 = -3;
    #[repr(C)]
    struct IDispatchVtbl {
        base: IUnknown_Vtbl,
        get_type_info_count: unsafe extern "system" fn(*mut c_void, *mut u32) -> HRESULT,
        get_type_info:
            unsafe extern "system" fn(*mut c_void, u32, u32, *mut *mut c_void) -> HRESULT,
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
    struct Dispatch;
    struct ComPtr<T> {
        raw: *mut c_void,
        _marker: PhantomData<T>,
    }
    impl<T> ComPtr<T> {
        unsafe fn vtbl(&self) -> &IDispatchVtbl {
            unsafe { &**(self.raw as *const *const IDispatchVtbl) }
        }
        unsafe fn from_borrowed(raw: *mut c_void) -> Option<Self> {
            if raw.is_null() {
                None
            } else {
                unsafe { ((&**(raw as *const *const IDispatchVtbl)).base.AddRef)(raw) };
                Some(Self {
                    raw,
                    _marker: PhantomData,
                })
            }
        }
        unsafe fn dispid(&self, name: &str, lcid: u32) -> Result<i32, HRESULT> {
            let wide: Vec<u16> = name.encode_utf16().chain(Some(0)).collect();
            let names = [wide.as_ptr()];
            let mut id = 0;
            let hr = unsafe {
                (self.vtbl().get_ids_of_names)(
                    self.raw,
                    &GUID::default(),
                    names.as_ptr(),
                    1,
                    lcid,
                    &mut id,
                )
            };
            if hr == 0 {
                Ok(id)
            } else {
                Err(hr)
            }
        }
    }
    impl<T> Drop for ComPtr<T> {
        fn drop(&mut self) {
            if !self.raw.is_null() {
                unsafe {
                    (self.vtbl().base.Release)(self.raw);
                }
            }
        }
    }
    struct ComApartment;
    impl ComApartment {
        fn initialize() -> Result<Self, String> {
            let hr = unsafe { CoInitializeEx(std::ptr::null(), COINIT_APARTMENTTHREADED as u32) };
            if hr >= 0 {
                Ok(Self)
            } else {
                Err(format!("CoInitializeEx failed: {}", hex(hr)))
            }
        }
    }
    impl Drop for ComApartment {
        fn drop(&mut self) {
            unsafe { CoUninitialize() }
        }
    }
    struct OwnedBstr(*const u16);
    impl OwnedBstr {
        fn from_text(text: &str) -> Result<Self, String> {
            let wide: Vec<u16> = text.encode_utf16().chain(Some(0)).collect();
            let value = unsafe { SysAllocString(wide.as_ptr()) };
            if value.is_null() {
                Err("SysAllocString returned null".to_owned())
            } else {
                Ok(Self(value))
            }
        }
    }
    impl Drop for OwnedBstr {
        fn drop(&mut self) {
            if !self.0.is_null() {
                unsafe { SysFreeString(self.0) }
            }
        }
    }
    struct OwnedVariant(VARIANT);
    impl OwnedVariant {
        fn empty() -> Self {
            let mut v = VARIANT::default();
            unsafe { VariantInit(&mut v) };
            Self(v)
        }
        fn bool(value: bool) -> Self {
            let mut v = Self::empty();
            v.0.Anonymous.Anonymous.vt = VT_BOOL;
            v.0.Anonymous.Anonymous.Anonymous.boolVal = if value { -1 } else { 0 };
            v
        }
        fn bstr(text: &str) -> Result<Self, String> {
            let b = OwnedBstr::from_text(text)?;
            let mut v = Self::empty();
            v.0.Anonymous.Anonymous.vt = VT_BSTR;
            v.0.Anonymous.Anonymous.Anonymous.bstrVal = b.0;
            std::mem::forget(b);
            Ok(v)
        }
        fn vt(&self) -> u16 {
            unsafe { self.0.Anonymous.Anonymous.vt }
        }
        fn i4(&self) -> Option<i32> {
            if self.vt() == VT_I4 {
                Some(unsafe { self.0.Anonymous.Anonymous.Anonymous.lVal })
            } else {
                None
            }
        }
        fn is_exact_42(&self) -> bool {
            match self.vt() {
                value if value == VT_I4 => self.i4() == Some(42),
                value if value == VT_R8 => unsafe {
                    self.0.Anonymous.Anonymous.Anonymous.dblVal == 42.0
                },
                _ => false,
            }
        }
        fn dispatch(&self) -> Option<ComPtr<Dispatch>> {
            if self.vt() != VT_DISPATCH {
                None
            } else {
                unsafe { ComPtr::from_borrowed(self.0.Anonymous.Anonymous.Anonymous.pdispVal) }
            }
        }
    }
    impl Drop for OwnedVariant {
        fn drop(&mut self) {
            unsafe {
                let _ = VariantClear(&mut self.0);
            }
        }
    }
    struct OwnedExcepInfo(EXCEPINFO);
    impl OwnedExcepInfo {
        fn new() -> Self {
            Self(EXCEPINFO::default())
        }
        fn take(&mut self) -> Value {
            let deferred = self.0.pfnDeferredFillIn.is_some();
            let deferred_hr = self
                .0
                .pfnDeferredFillIn
                .map(|fill| unsafe { fill(&mut self.0) });
            let scode = self.0.scode;
            unsafe {
                for value in [
                    &mut self.0.bstrSource,
                    &mut self.0.bstrDescription,
                    &mut self.0.bstrHelpFile,
                ] {
                    if !(*value).is_null() {
                        SysFreeString(*value);
                        *value = std::ptr::null();
                    }
                }
            }
            json!({"deferred_fill_in_present":deferred,"deferred_fill_in_hresult":deferred_hr.map(hex),"scode":hex(scode),"wcode":self.0.wCode})
        }
    }
    impl Drop for OwnedExcepInfo {
        fn drop(&mut self) {
            let _ = self.take();
        }
    }
    struct Frame {
        args: Vec<OwnedVariant>,
        named: Vec<i32>,
    }
    impl Frame {
        fn empty() -> Self {
            Self {
                args: vec![],
                named: vec![],
            }
        }
        fn positional(mut args: Vec<OwnedVariant>) -> Self {
            args.reverse();
            Self {
                args,
                named: vec![],
            }
        }
        fn put(value: OwnedVariant) -> Self {
            Self {
                args: vec![value],
                named: vec![DISPID_PROPERTYPUT],
            }
        }
        fn params(&mut self) -> DISPPARAMS {
            DISPPARAMS {
                rgvarg: if self.args.is_empty() {
                    std::ptr::null_mut()
                } else {
                    self.args.as_mut_ptr().cast()
                },
                rgdispidNamedArgs: if self.named.is_empty() {
                    std::ptr::null_mut()
                } else {
                    self.named.as_mut_ptr()
                },
                cArgs: self.args.len() as u32,
                cNamedArgs: self.named.len() as u32,
            }
        }
    }
    struct Call {
        hr: HRESULT,
        result: OwnedVariant,
        exception: Value,
        arg_error: u32,
    }
    fn call(
        target: &ComPtr<Dispatch>,
        member: &str,
        flags: u16,
        lcid: u32,
        mut frame: Frame,
    ) -> Call {
        let mut result = OwnedVariant::empty();
        let mut exception = OwnedExcepInfo::new();
        let mut arg = u32::MAX;
        let params = frame.params();
        let hr = match unsafe { target.dispid(member, lcid) } {
            Ok(id) => unsafe {
                (target.vtbl().invoke)(
                    target.raw,
                    id,
                    &GUID::default(),
                    lcid,
                    flags,
                    &params,
                    &mut result.0,
                    &mut exception.0,
                    &mut arg,
                )
            },
            Err(hr) => hr,
        };
        let exception = exception.take();
        Call {
            hr,
            result,
            exception,
            arg_error: arg,
        }
    }
    struct OwnedProcess(HANDLE);
    impl OwnedProcess {
        fn from_app(app: &ComPtr<Dispatch>, lcid: u32) -> Result<Self, String> {
            let call = call(app, "Hwnd", DISPATCH_PROPERTYGET, lcid, Frame::empty());
            if call.hr != 0 {
                return Err(format!("Hwnd failed: {}", hex(call.hr)));
            }
            let hwnd = call.result.i4().ok_or("Hwnd did not return VT_I4")?;
            let mut pid = 0;
            unsafe { GetWindowThreadProcessId(hwnd as isize as HWND, &mut pid) };
            if pid == 0 {
                return Err("Hwnd ownership lookup returned zero".to_owned());
            }
            let handle = unsafe {
                OpenProcess(
                    PROCESS_QUERY_LIMITED_INFORMATION | PROCESS_SYNCHRONIZE,
                    0,
                    pid,
                )
            };
            if handle.is_null() {
                return Err("OpenProcess for owned Excel returned null".to_owned());
            }
            let mut creation = FILETIME::default();
            if unsafe {
                GetProcessTimes(
                    handle,
                    &mut creation,
                    &mut FILETIME::default(),
                    &mut FILETIME::default(),
                    &mut FILETIME::default(),
                )
            } == 0
            {
                unsafe { CloseHandle(handle) };
                return Err("GetProcessTimes for owned Excel failed".to_owned());
            }
            Ok(Self(handle))
        }
        fn wait(&self) -> bool {
            unsafe { WaitForSingleObject(self.0, 15_000) == WAIT_OBJECT_0 }
        }
    }
    impl Drop for OwnedProcess {
        fn drop(&mut self) {
            if !self.0.is_null() {
                unsafe {
                    CloseHandle(self.0);
                }
            }
        }
    }
    fn activate(mode: Mode) -> Result<ComPtr<Dispatch>, String> {
        let program: Vec<u16> = "Excel.Application".encode_utf16().chain(Some(0)).collect();
        let mut clsid = GUID::default();
        let hr = unsafe { CLSIDFromProgID(program.as_ptr(), &mut clsid) };
        if hr != 0 {
            return Err(format!("CLSIDFromProgID failed: {}", hex(hr)));
        }
        let mut raw = std::ptr::null_mut();
        let hr = match mode {
            Mode::L | Mode::S => unsafe {
                CoCreateInstance(
                    &clsid,
                    std::ptr::null_mut(),
                    if matches!(mode, Mode::L) {
                        CLSCTX_LOCAL_SERVER
                    } else {
                        CLSCTX_SERVER
                    },
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
                        &clsid,
                        std::ptr::null_mut(),
                        CLSCTX_SERVER,
                        std::ptr::null(),
                        1,
                        &mut request,
                    )
                };
                raw = request.pItf;
                if outer == 0 {
                    request.hr
                } else {
                    outer
                }
            }
        };
        if hr != 0 || raw.is_null() {
            return Err(format!("activation failed: {}", hex(hr)));
        }
        Ok(ComPtr {
            raw,
            _marker: PhantomData,
        })
    }
    fn get(
        target: &ComPtr<Dispatch>,
        name: &str,
        lcid: u32,
        args: Vec<OwnedVariant>,
    ) -> Result<ComPtr<Dispatch>, String> {
        let call = call(
            target,
            name,
            DISPATCH_PROPERTYGET,
            lcid,
            Frame::positional(args),
        );
        if call.hr != 0 {
            return Err(format!("{name} failed: {}", hex(call.hr)));
        }
        call.result
            .dispatch()
            .ok_or_else(|| format!("{name} did not return VT_DISPATCH"))
    }
    fn brief(call: &Call) -> Value {
        let applicable =
            matches!(call.hr, x if x==0x8002_0004_u32 as i32 || x==0x8002_0005_u32 as i32);
        json!({"hresult":call.hr,"hresult_hex":hex(call.hr),"result_vartype":call.result.vt(),"excepinfo":call.exception,"pu_arg_err":if applicable && call.arg_error!=u32::MAX{Value::from(call.arg_error)}else{Value::String("not-applicable".to_owned())}})
    }
    pub fn run(mode: Mode, fixture: &Path) -> Result<Value, String> {
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
                Frame::positional(vec![OwnedVariant::bool(false)]),
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
        let mut smoke = json!({"entered":false});
        let mut failure = None;
        let opened = if open.hr == 0 {
            open.result.dispatch()
        } else {
            None
        };
        if let Some(book) = opened.as_ref() {
            match get(book, "Worksheets", lcid, vec![])
                .and_then(|sheets| {
                    get(
                        &sheets,
                        "Item",
                        lcid,
                        vec![{
                            let mut v = OwnedVariant::empty();
                            v.0.Anonymous.Anonymous.vt = VT_I4;
                            v.0.Anonymous.Anonymous.Anonymous.lVal = 1;
                            v
                        }],
                    )
                })
                .and_then(|sheet| get(&sheet, "Range", lcid, vec![OwnedVariant::bstr("A1")?]))
            {
                Ok(range) => {
                    let write = call(
                        &range,
                        "Value2",
                        DISPATCH_PROPERTYPUT,
                        lcid,
                        Frame::put({
                            let mut v = OwnedVariant::empty();
                            v.0.Anonymous.Anonymous.vt = VT_I4;
                            v.0.Anonymous.Anonymous.Anonymous.lVal = 42;
                            v
                        }),
                    );
                    let read = call(&range, "Value2", DISPATCH_PROPERTYGET, lcid, Frame::empty());
                    let clear = call(
                        &range,
                        "ClearContents",
                        DISPATCH_METHOD,
                        lcid,
                        Frame::empty(),
                    );
                    let exact = read.hr == 0 && read.result.is_exact_42();
                    smoke = json!({"entered":true,"write":brief(&write),"read":brief(&read),"read_value_exactly_42":exact,"clear":brief(&clear)});
                    if !exact || write.hr != 0 || clear.hr != 0 {
                        failure = Some(
                            json!({"inner_scode_hex":read.exception.get("scode").and_then(Value::as_str).unwrap_or("--"),"workbook_returned":true}),
                        );
                    }
                }
                Err(error) => {
                    failure = Some(
                        json!({"detail":error,"inner_scode_hex":"--","workbook_returned":true}),
                    );
                }
            }
        } else {
            failure = Some(
                json!({"inner_scode_hex":open.exception.get("scode").and_then(Value::as_str).unwrap_or("--"),"workbook_returned":false}),
            );
        }
        let _ = if let Some(book) = opened {
            let _ = call(
                &book,
                "Close",
                DISPATCH_METHOD,
                lcid,
                Frame::positional(vec![OwnedVariant::bool(false)]),
            );
            Some(())
        } else {
            None
        };
        let quit = call(&app, "Quit", DISPATCH_METHOD, lcid, Frame::empty());
        drop(workbooks);
        drop(app);
        let exited = owned.wait();
        let success = add_report.get("hresult") == Some(&Value::from(0))
            && open_report.get("hresult") == Some(&Value::from(0))
            && smoke.get("read_value_exactly_42") == Some(&Value::Bool(true))
            && exited;
        Ok(
            json!({"backend":"raw-windows-sys","activation_mode":mode.id(),"activation":mode.activation(),"version":brief(&version),"workbooks_add":add_report,"workbooks_open":open_report,"range_smoke":smoke,"cleanup":{"excel_quit":brief(&quit),"owned_process_exit_verified":exited,"forced_termination":false},"failure":failure,"success":success,"raw_pointer_values_recorded":false}),
        )
    }
    fn hex(hr: i32) -> String {
        format!("0x{:08X}", hr as u32)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn modes_are_explicit() {
        assert_eq!(Mode::parse("all").unwrap().len(), 3);
        assert!(Mode::parse("Q").is_err());
    }
    #[test]
    fn empty_jsonl_has_final_lf() {
        assert_eq!(jsonl(&[]).unwrap(), "\n");
    }
    #[test]
    fn retry_is_narrow() {
        let row = json!({"workbooks_add":{"hresult":-1},"failure":{"inner_scode_hex":"0x800A03EC","workbook_returned":false}});
        assert!(retry_eligible(&row));
        assert!(!retry_eligible(&json!({})));
    }
}
