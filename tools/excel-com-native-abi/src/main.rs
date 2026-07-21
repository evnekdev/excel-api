//! Bounded research-only ABI differential for Excel Automation.
//!
//! The direct-vtable declaration below is the generic SDK `IDispatch` ABI
//! (IUnknown + GetTypeInfoCount + GetTypeInfo + GetIDsOfNames + Invoke).  It
//! deliberately contains no Excel-specific interface layout or DISPIDs.

#![cfg_attr(not(windows), allow(dead_code, unused_imports))]
#![allow(
    clippy::collapsible_if,
    clippy::too_many_arguments,
    clippy::unnecessary_to_owned
)]

use std::env;
use std::fs;
use std::path::{Path, PathBuf};

const SCHEMA_VERSION: u32 = 1;
const WINDOWS_VERSION: &str = "0.62.2";
const WINDOWS_CORE_VERSION: &str = "0.62.2";
const WINDOWS_RESULT_VERSION: &str = "0.4.1";
const WINDOWS_STRINGS_VERSION: &str = "0.5.1";
const WINDOWS_SYS_VERSION: &str = "0.61.2";

fn main() -> Result<(), String> {
    let mut args = env::args().skip(1);
    let command = args.next().ok_or_else(usage)?;
    let options = Options::parse(args.collect())?;
    match command.as_str() {
        "generate" => generate(options.root()?),
        "render" => render(options.root()?),
        "run-native" => run_native(options),
        "run-shim" => run_shim(options),
        "run-high" => run_high(options),
        "run-lower" => run_lower(options),
        "record-native-layout" => record_native_layout(options),
        "check" => check(options.root()?),
        "layout" => print_layout(),
        _ => Err(usage()),
    }
}

fn usage() -> String {
    "usage: excel-com-native-abi <generate|render|check|run-native|run-shim|run-high|run-lower|record-native-layout|layout> --root <knowledge/excel-object-model/native-abi> [--mode <mode>] [--fixture <path>] [--native-exe <path>] [--shim <path>]".to_owned()
}

#[derive(Default)]
struct Options {
    root: Option<PathBuf>,
    mode: Option<String>,
    fixture: Option<PathBuf>,
    native_exe: Option<PathBuf>,
    shim: Option<PathBuf>,
}

impl Options {
    fn parse(values: Vec<String>) -> Result<Self, String> {
        let mut result = Self::default();
        let mut index = 0;
        while index < values.len() {
            let value = values[index].as_str();
            let next = values
                .get(index + 1)
                .ok_or_else(|| format!("missing value for {value}"))?;
            match value {
                "--root" => result.root = Some(PathBuf::from(next)),
                "--mode" => result.mode = Some(next.clone()),
                "--fixture" => result.fixture = Some(PathBuf::from(next)),
                "--native-exe" => result.native_exe = Some(PathBuf::from(next)),
                "--shim" => result.shim = Some(PathBuf::from(next)),
                _ => return Err(format!("unknown option: {value}")),
            }
            index += 2;
        }
        Ok(result)
    }

    fn root(&self) -> Result<&Path, String> {
        self.root
            .as_deref()
            .ok_or_else(|| "--root is required".to_owned())
    }

    fn mode(&self) -> Result<&str, String> {
        self.mode
            .as_deref()
            .ok_or_else(|| "--mode is required".to_owned())
    }
}

fn generate(root: &Path) -> Result<(), String> {
    fs::create_dir_all(root).map_err(io_error)?;
    fs::create_dir_all(report_root(root)).map_err(io_error)?;
    write(root.join("SOURCE_MANIFEST.toml"), source_manifest())?;
    write(
        root.join("environments.jsonl"),
        format!("{}\n", environment_json()),
    )?;
    write(root.join("operation-specs.jsonl"), operation_specs())?;
    write(
        root.join("abi-layouts.jsonl"),
        format!("{}\n", layout_json("rust-generated-layout")),
    )?;
    write(root.join("version-matrix.jsonl"), version_matrix())?;
    write(root.join("unresolved.jsonl"), unresolved_records())?;
    for name in [
        "native-observations.jsonl",
        "shim-observations.jsonl",
        "rust-observations.jsonl",
        "interface-lifetimes.jsonl",
    ] {
        let path = root.join(name);
        if !path.exists() {
            write(path, String::new())?;
        }
    }
    render(root)
}

fn record_native_layout(options: Options) -> Result<(), String> {
    let root = options.root()?.to_path_buf();
    write(root.join("version-matrix.jsonl"), version_matrix())?;
    let executable = options
        .native_exe
        .ok_or_else(|| "--native-exe is required".to_owned())?;
    let output = std::process::Command::new(executable)
        .arg("--layout")
        .output()
        .map_err(io_error)?;
    if !output.status.success() {
        return Err(format!(
            "native layout control exited with {}",
            output.status
        ));
    }
    let record = String::from_utf8(output.stdout).map_err(|error| error.to_string())?;
    append(root.join("abi-layouts.jsonl"), record.trim())?;
    render(&root)
}

fn check(root: &Path) -> Result<(), String> {
    for name in [
        "SOURCE_MANIFEST.toml",
        "environments.jsonl",
        "abi-layouts.jsonl",
        "operation-specs.jsonl",
        "version-matrix.jsonl",
        "native-observations.jsonl",
        "shim-observations.jsonl",
        "rust-observations.jsonl",
        "interface-lifetimes.jsonl",
        "unresolved.jsonl",
    ] {
        if !root.join(name).is_file() {
            return Err(format!("missing evidence file: {name}"));
        }
    }
    for name in [
        "abi-layout-comparison.md",
        "calling-convention-audit.md",
        "native-vs-rust-operation-matrix.md",
        "native-shim-matrix.md",
        "interface-lifetime-comparison.md",
        "excepinfo-comparison.md",
        "workbook-add-differential.md",
        "workbook-open-differential.md",
        "root-cause-classification.md",
        "remaining-blockers.md",
    ] {
        if !report_root(root).join(name).is_file() {
            return Err(format!("missing generated report: {name}"));
        }
    }
    for name in [
        "native-observations.jsonl",
        "shim-observations.jsonl",
        "rust-observations.jsonl",
    ] {
        let content = fs::read_to_string(root.join(name)).map_err(io_error)?;
        if content.contains("C:\\") || content.contains("0x0000") || content.contains("ptr=") {
            return Err(format!("unsafe evidence persistence in {name}"));
        }
    }
    if !fs::read_to_string(root.join("abi-layouts.jsonl"))
        .map_err(io_error)?
        .contains("native-cpp-sdk-layout")
    {
        return Err("native C++ ABI layout record is missing".to_owned());
    }
    Ok(())
}

fn render(root: &Path) -> Result<(), String> {
    let reports = report_root(root);
    let read = |name: &str| fs::read_to_string(root.join(name)).unwrap_or_default();
    let native = read("native-observations.jsonl");
    let shim = read("shim-observations.jsonl");
    let rust = read("rust-observations.jsonl");
    let lifetimes = read("interface-lifetimes.jsonl");
    let layouts = read("abi-layouts.jsonl");
    let versions = read("version-matrix.jsonl");
    let unresolved = read("unresolved.jsonl");
    write(
        reports.join("abi-layout-report.md"),
        format!(
            "# ABI layout report\n\nGenerated deterministically by `excel-com-native-abi`.\n\n{}\n",
            layout_json("rust-generated-layout")
        ),
    )?;
    write(
        reports.join("native-direct-results.md"),
        format!(
            "# Native direct results\n\n{}",
            if native.is_empty() {
                "Not tested.\n"
            } else {
                &native
            }
        ),
    )?;
    write(
        reports.join("native-shim-results.md"),
        format!(
            "# Native shim results\n\n{}",
            if shim.is_empty() {
                "Not tested.\n"
            } else {
                &shim
            }
        ),
    )?;
    write(
        reports.join("rust-direct-results.md"),
        format!(
            "# Rust direct results\n\n{}",
            if rust.is_empty() {
                "Not tested.\n"
            } else {
                &rust
            }
        ),
    )?;
    write(
        reports.join("interface-lifetime-results.md"),
        format!(
            "# Interface lifetime results\n\n{}",
            if lifetimes.is_empty() {
                "Not tested.\n"
            } else {
                &lifetimes
            }
        ),
    )?;
    write(
        reports.join("reproduction-matrix.md"),
        format!(
            "# Reproduction matrix\n\n| Path | Captured rows |\n| --- | ---: |\n| Native C++ direct | {} |\n| Native C ABI shim from Rust | {} |\n| Rust high-level / lower direct | {} |\n",
            lines(&native),
            lines(&shim),
            lines(&rust)
        ),
    )?;
    write(
        reports.join("abi-layout-comparison.md"),
        format!(
            "# ABI layout comparison\n\nC++ SDK and Rust records agree on x64 `VARIANT` (24/8), `DISPPARAMS` (24/8), `EXCEPINFO` (64/8), `GUID` (16/4), all recorded offsets, and the fixed C result ABI (88/4).\n\n{layouts}"
        ),
    )?;
    write(
        reports.join("calling-convention-audit.md"),
        "# Calling-convention audit\n\nThe SDK and `windows 0.62.2` generated `IDispatch_Vtbl::Invoke` use `unsafe extern \"system\" fn(*mut c_void, i32, *const GUID, u32, DISPATCH_FLAGS/WORD, *const DISPPARAMS, *mut VARIANT, *mut EXCEPINFO, *mut u32) -> HRESULT`. `HRESULT` and `DISPID` are signed 32-bit; `LCID` and `UINT` counts are unsigned 32-bit; flags are 16-bit. `IUnknown` precedes `GetTypeInfoCount`, `GetTypeInfo`, `GetIDsOfNames`, and `Invoke`. `CoCreateInstance` and `CoCreateInstanceEx` are generated `system` imports in both windows and windows-sys.\n".to_owned(),
    )?;
    write(
        reports.join("native-vs-rust-operation-matrix.md"),
        format!(
            "# Native vs Rust operation matrix\n\nNative C ABI shim rows:\n\n{shim}\nRust rows:\n\n{rust}"
        ),
    )?;
    write(
        reports.join("native-shim-matrix.md"),
        format!(
            "# Native shim matrix\n\nThe Rust caller receives copied fixed-width fields only; it transfers neither interface pointers nor BSTR ownership.\n\n{shim}"
        ),
    )?;
    write(
        reports.join("interface-lifetime-comparison.md"),
        format!(
            "# Interface lifetime comparison\n\nEach recorded row runs clone-then-clear, retain-then-clear, and QueryInterface-then-clear followed by `Workbooks.Count`; zero is success.\n\n{lifetimes}"
        ),
    )?;
    write(
        reports.join("excepinfo-comparison.md"),
        format!(
            "# EXCEPINFO comparison\n\nSuccess rows use zero `inner_scode`. The high-level local/0x0400 failure is `DISP_E_EXCEPTION` (`-2147352567`) with copied `scode` `-2146827284` (`0x800A03EC`) and raw `puArgErr` 0, which is not applicable to a zero-argument call. BSTR text is not persisted; historical Prompt 05B records preserve its normalized diagnostics.\n\n{rust}"
        ),
    )?;
    write(
        reports.join("workbook-add-differential.md"),
        "# Workbooks.Add differential\n\nThe native C ABI shim and windows-sys generic-IDispatch paths succeed in every required activation mode. The high-level windows harness fails only for `CoCreateInstance(CLSCTX_LOCAL_SERVER)` with LCID 0x0400; its isolated minimal high-level reproduction succeeds. This is not a crate-version regression. The final standalone C++ runner's conflicting local/0x0400 failure remains unresolved.\n".to_owned(),
    )?;
    write(
        reports.join("workbook-open-differential.md"),
        "# Workbooks.Open differential\n\nThe shared Python-created fixture is opened and closed successfully by the native C ABI shim and windows-sys generic-IDispatch paths in all required activation modes. The path is redacted from evidence.\n".to_owned(),
    )?;
    write(
        reports.join("root-cause-classification.md"),
        "# Root-cause classification\n\n**Inconclusive, with a narrowed negative finding:** the native C ABI shim and lower-level Rust generic `IDispatch` path succeed, while the full high-level Rust local/0x0400 harness fails. The minimal high-level reproduction succeeds on both current and preceding released windows-rs versions. This does not confirm a windows-rs regression. The final standalone C++ runner's conflicting local/0x0400 result prevents a clean Case D classification. Prompt 05 remains blocked pending a targeted production-harness repair and revalidation.\n".to_owned(),
    )?;
    write(
        reports.join("remaining-blockers.md"),
        format!(
            "# Remaining blockers\n\n1. Isolate the full high-level local/0x0400 pre-Add sequence against the passing minimal sequence before changing production behavior.\n2. Re-run the production range probe only after that bounded repair.\n3. The current windows-rs source checkout could not compile this released-API reproduction because its feature model no longer exposes the released Win32 feature names.\n\nVersion matrix:\n\n{versions}\nAdditional unresolved rows:\n\n{unresolved}"
        ),
    )?;
    Ok(())
}

fn report_root(root: &Path) -> PathBuf {
    root.parent()
        .unwrap_or(root)
        .join("generated")
        .join("native-abi")
}

fn run_native(options: Options) -> Result<(), String> {
    let root = options.root()?.to_path_buf();
    let mode = options.mode()?.to_owned();
    let executable = options
        .native_exe
        .ok_or_else(|| "--native-exe is required".to_owned())?;
    let mut command = std::process::Command::new(executable);
    command.arg("--mode").arg(&mode);
    if let Some(fixture) = options.fixture {
        command.arg("--fixture").arg(fixture);
    }
    let output = command.output().map_err(io_error)?;
    if !output.status.success() {
        return Err(format!("native control exited with {}", output.status));
    }
    let observation = String::from_utf8(output.stdout).map_err(|error| error.to_string())?;
    append(root.join("native-observations.jsonl"), observation.trim())?;
    append(root.join("interface-lifetimes.jsonl"), observation.trim())?;
    render(&root)
}

fn run_shim(options: Options) -> Result<(), String> {
    #[cfg(windows)]
    {
        let root = options.root()?.to_path_buf();
        let mode = mode_number(options.mode()?)?;
        let fixture = options.fixture.as_deref();
        let shim = options
            .shim
            .as_deref()
            .ok_or_else(|| "--shim is required".to_owned())?;
        let result = unsafe { call_shim(shim, mode, fixture)? };
        let json = result_json("rust-c-abi-shim", mode, &result);
        append(root.join("shim-observations.jsonl"), &json)?;
        append(root.join("interface-lifetimes.jsonl"), &json)?;
        render(&root)
    }
    #[cfg(not(windows))]
    {
        let _ = options;
        Err("Windows is required".to_owned())
    }
}

fn run_high(options: Options) -> Result<(), String> {
    #[cfg(windows)]
    {
        let root = options.root()?.to_path_buf();
        let mode = mode_number(options.mode()?)?;
        let result = unsafe { high_level_run(mode, options.fixture.as_deref()) };
        let json = result_json("rust-windows-high-level", mode, &result);
        append(root.join("rust-observations.jsonl"), &json)?;
        append(root.join("interface-lifetimes.jsonl"), &json)?;
        render(&root)
    }
    #[cfg(not(windows))]
    {
        let _ = options;
        Err("Windows is required".to_owned())
    }
}

fn run_lower(options: Options) -> Result<(), String> {
    #[cfg(windows)]
    {
        let root = options.root()?.to_path_buf();
        let mode = mode_number(options.mode()?)?;
        let result = unsafe { lower_level_run(mode, options.fixture.as_deref()) };
        let json = result_json("rust-windows-sys-generic-idispatch-vtbl", mode, &result);
        append(root.join("rust-observations.jsonl"), &json)?;
        append(root.join("interface-lifetimes.jsonl"), &json)?;
        render(&root)
    }
    #[cfg(not(windows))]
    {
        let _ = options;
        Err("Windows is required".to_owned())
    }
}

fn mode_number(mode: &str) -> Result<u32, String> {
    match mode {
        "native-cocreate-local-lcid-0400" => Ok(1),
        "native-cocreate-server-lcid-0000" => Ok(2),
        "native-cocreateex-server-lcid-0000" => Ok(3),
        _ => Err("mode must be native-cocreate-local-lcid-0400, native-cocreate-server-lcid-0000, or native-cocreateex-server-lcid-0000".to_owned()),
    }
}

fn source_manifest() -> String {
    format!(
        "schema_version = {SCHEMA_VERSION}\nstarting_origin_master = \"91ab21df735d38928e27775aea8f4ecad3499821\"\nwindows = \"{WINDOWS_VERSION}\"\nwindows_core = \"{WINDOWS_CORE_VERSION}\"\nwindows_result = \"{WINDOWS_RESULT_VERSION}\"\nwindows_strings = \"{WINDOWS_STRINGS_VERSION}\"\nwindows_sys = \"{WINDOWS_SYS_VERSION}\"\nwindows_version = \"Windows 10 Enterprise 25H2 build 26200.8875\"\nexcel_file_version = \"16.0.20131.20154\"\noffice_bitness = \"64-bit\"\nmsvc = \"19.40.33812.0 (Visual Studio 2022 Community 17.10.4)\"\ncmake = \"3.30.2\"\nninja = \"1.12.1\"\nrustc = \"1.97.1\"\narchitecture = \"x86_64-pc-windows-msvc required for live rows\"\nruntime_library = \"MSVC MultiThreadedDLL\"\nsource_inspection = \"windows 0.62.2 generated IDispatch::Invoke and IDispatch_Vtbl; windows-sys 0.61.2 generated COM, DISPPARAMS, EXCEPINFO, VARIANT, CoCreateInstance, and CoCreateInstanceEx declarations\"\ncurrent_windows_rs_source = \"447078ea771a97277b710de1e3149c5146af1dc8 (isolated; incompatible Win32 feature model)\"\nenvironment_adjustments = \"initial MinGW configure rejected; reconfigured clean temporary x64 MSVC CMake/Ninja build; no Office, registration, or security changes\"\n"
    )
}

fn environment_json() -> String {
    format!(
        "{{\"schema_version\":1,\"id\":\"windows-excel-64-05e\",\"classification\":\"Version-specific\",\"windows_version\":\"Windows 10 Enterprise 25H2 build 26200.8875\",\"excel_file_version\":\"16.0.20131.20154\",\"office_bitness\":\"64-bit\",\"toolchain\":{{\"msvc\":\"19.40.33812.0\",\"cmake\":\"3.30.2\",\"ninja\":\"1.12.1\",\"rustc\":\"1.97.1\"}},\"windows_crates\":{{\"windows\":\"{WINDOWS_VERSION}\",\"windows-core\":\"{WINDOWS_CORE_VERSION}\",\"windows-result\":\"{WINDOWS_RESULT_VERSION}\",\"windows-strings\":\"{WINDOWS_STRINGS_VERSION}\",\"windows-sys\":\"{WINDOWS_SYS_VERSION}\"}},\"pointer_width\":64,\"raw_pointer_values_recorded\":false}}"
    )
}

fn operation_specs() -> String {
    [
        "{\"schema_version\":1,\"id\":\"application-version\",\"member\":\"Excel.Application.Version\",\"dispid\":392,\"flags\":2,\"lcid\":\"mode-defined\",\"args\":0}",
        "{\"schema_version\":1,\"id\":\"application-workbooks\",\"member\":\"Excel.Application.Workbooks\",\"dispid\":572,\"flags\":2,\"lcid\":\"mode-defined\",\"args\":0}",
        "{\"schema_version\":1,\"id\":\"workbooks-count\",\"member\":\"Excel.Workbooks.Count\",\"dispid\":118,\"flags\":2,\"lcid\":\"mode-defined\",\"args\":0}",
        "{\"schema_version\":1,\"id\":\"workbooks-add\",\"member\":\"Excel.Workbooks.Add\",\"dispid\":181,\"flags\":1,\"lcid\":\"mode-defined\",\"args\":0,\"omission\":\"null rgvarg and null named-DISPID pointer\"}",
        "{\"schema_version\":1,\"id\":\"workbooks-open\",\"member\":\"Excel.Workbooks.Open\",\"dispid\":1923,\"flags\":1,\"lcid\":\"mode-defined\",\"args\":1,\"fixture_path_recorded\":false}",
        "{\"schema_version\":1,\"id\":\"application-quit\",\"member\":\"Excel.Application.Quit\",\"dispid\":302,\"flags\":1,\"lcid\":\"mode-defined\",\"args\":0}",
    ].join("\n") + "\n"
}

fn version_matrix() -> String {
    [
        "{\"schema_version\":1,\"id\":\"released-current\",\"windows\":\"0.62.2\",\"windows-core\":\"0.62.2\",\"windows-result\":\"0.4.1\",\"windows-strings\":\"0.5.1\",\"windows-sys\":\"0.61.2\",\"minimal_high_level_add_hresult\":0,\"classification\":\"Runtime-observed\"}",
        "{\"schema_version\":1,\"id\":\"released-preceding\",\"windows\":\"0.62.1\",\"windows-core\":\"0.62.1\",\"windows-result\":\"0.4.0\",\"windows-strings\":\"0.5.0\",\"windows-sys\":\"0.61.1\",\"minimal_high_level_add_hresult\":0,\"classification\":\"Runtime-observed\"}",
        "{\"schema_version\":1,\"id\":\"source-head-447078ea771a97277b710de1e3149c5146af1dc8\",\"windows\":\"0.62.2 source checkout\",\"minimal_high_level_add_hresult\":null,\"classification\":\"Inconclusive\",\"blocker\":\"current source no longer exposes the 0.62 Win32 feature names required by this released-API reproduction\"}",
    ].join("\n") + "\n"
}

fn unresolved_records() -> String {
    [
        "{\"schema_version\":1,\"id\":\"native-cpp-runner-local-lcid-0400\",\"classification\":\"Inconclusive\",\"detail\":\"The final standalone C++ runner returned DISP_E_EXCEPTION/0x800A03EC although the C ABI DLL built from the same source succeeded through Rust in every mode.\",\"effect\":\"Prevents a clean Case D conclusion.\"}",
        "{\"schema_version\":1,\"id\":\"windows-rs-source-head\",\"classification\":\"Inconclusive\",\"detail\":\"The isolated source-head checkout could not compile the released Win32-feature reproduction because its feature model no longer exposes those feature names.\",\"effect\":\"No source-head runtime claim.\"}",
    ].join("\n") + "\n"
}

fn print_layout() -> Result<(), String> {
    println!("{}", layout_json("rust-generated-layout"));
    Ok(())
}

fn layout_json(path: &str) -> String {
    #[cfg(windows)]
    {
        use std::mem::{align_of, offset_of, size_of};
        use windows::Win32::System::Com::{DISPPARAMS, EXCEPINFO};
        use windows::Win32::System::Variant::VARIANT;
        format!(
            "{{\"schema_version\":1,\"path\":\"{path}\",\"pointer_width\":{},\"guid\":{{\"size\":{},\"align\":{}}},\"result\":{{\"size\":{},\"align\":{}}},\"variant\":{{\"size\":{},\"align\":{},\"vt_offset\":{},\"data_offset\":8,\"error_offset\":8,\"i4_offset\":8,\"bstr_offset\":8,\"dispatch_offset\":8}},\"dispparams\":{{\"size\":{},\"align\":{},\"rgvarg_offset\":{},\"named_offset\":{},\"args_offset\":{},\"named_count_offset\":{}}},\"excepinfo\":{{\"size\":{},\"align\":{},\"wcode_offset\":{},\"source_offset\":{},\"description_offset\":{},\"help_file_offset\":{},\"help_context_offset\":{},\"reserved_offset\":{},\"deferred_offset\":{},\"scode_offset\":{}}},\"raw_pointer_values_recorded\":false}}",
            size_of::<*const ()>() * 8,
            size_of::<windows::core::GUID>(),
            align_of::<windows::core::GUID>(),
            size_of::<RawResult>(),
            align_of::<RawResult>(),
            size_of::<VARIANT>(),
            align_of::<VARIANT>(),
            0,
            size_of::<DISPPARAMS>(),
            align_of::<DISPPARAMS>(),
            offset_of!(DISPPARAMS, rgvarg),
            offset_of!(DISPPARAMS, rgdispidNamedArgs),
            offset_of!(DISPPARAMS, cArgs),
            offset_of!(DISPPARAMS, cNamedArgs),
            size_of::<EXCEPINFO>(),
            align_of::<EXCEPINFO>(),
            offset_of!(EXCEPINFO, wCode),
            offset_of!(EXCEPINFO, bstrSource),
            offset_of!(EXCEPINFO, bstrDescription),
            offset_of!(EXCEPINFO, bstrHelpFile),
            offset_of!(EXCEPINFO, dwHelpContext),
            offset_of!(EXCEPINFO, pvReserved),
            offset_of!(EXCEPINFO, pfnDeferredFillIn),
            offset_of!(EXCEPINFO, scode),
        )
    }
    #[cfg(not(windows))]
    {
        format!("{{\"schema_version\":1,\"path\":\"{path}\",\"status\":\"Windows required\"}}")
    }
}

fn result_json(path: &str, mode: u32, result: &RawResult) -> String {
    format!(
        "{{\"schema_version\":1,\"path\":\"{path}\",\"mode\":{mode},\"activation_hresult\":{},\"version_hresult\":{},\"workbooks_hresult\":{},\"count_hresult\":{},\"add_hresult\":{},\"open_hresult\":{},\"quit_hresult\":{},\"inner_scode\":{},\"deferred_fill_in_hresult\":{},\"result_vt\":{},\"workbooks_vt\":{},\"pu_arg_err_raw\":{},\"workbook_created\":{},\"workbook_opened\":{},\"process_exited\":{},\"lifetime_clone_then_clear\":{},\"lifetime_retain_then_clear\":{},\"lifetime_query_interface_then_clear\":{},\"type_info_count\":{},\"type_info_hresult\":{},\"workbooks_query_iunknown_hresult\":{},\"workbooks_query_idispatch_hresult\":{},\"raw_paths_recorded\":false,\"raw_hwnd_recorded\":false,\"raw_pointer_values_recorded\":false}}",
        result.activation_hresult,
        result.version_hresult,
        result.workbooks_hresult,
        result.count_hresult,
        result.add_hresult,
        result.open_hresult,
        result.quit_hresult,
        result.inner_scode,
        result.deferred_fill_in_hresult,
        result.result_vt,
        result.workbooks_vt,
        result.pu_arg_err_raw,
        result.workbook_created,
        result.workbook_opened,
        result.process_exited,
        result.lifetime_clone_then_clear,
        result.lifetime_retain_then_clear,
        result.lifetime_query_interface_then_clear,
        result.type_info_count,
        result.type_info_hresult,
        result.workbooks_query_iunknown_hresult,
        result.workbooks_query_idispatch_hresult
    )
}

#[repr(C)]
#[derive(Default)]
struct RawResult {
    schema_version: u32,
    activation_hresult: i32,
    version_hresult: i32,
    workbooks_hresult: i32,
    count_hresult: i32,
    add_hresult: i32,
    open_hresult: i32,
    quit_hresult: i32,
    inner_scode: i32,
    deferred_fill_in_hresult: i32,
    result_vt: u16,
    workbooks_vt: u16,
    pu_arg_err_raw: u32,
    workbook_created: i32,
    workbook_opened: i32,
    process_exited: i32,
    lifetime_clone_then_clear: i32,
    lifetime_retain_then_clear: i32,
    lifetime_query_interface_then_clear: i32,
    type_info_count: u32,
    type_info_hresult: i32,
    workbooks_query_iunknown_hresult: i32,
    workbooks_query_idispatch_hresult: i32,
}

fn write(path: impl AsRef<Path>, content: impl AsRef<[u8]>) -> Result<(), String> {
    fs::write(path, content).map_err(io_error)
}
fn append(path: impl AsRef<Path>, content: &str) -> Result<(), String> {
    use std::io::Write;
    let mut file = fs::OpenOptions::new()
        .create(true)
        .append(true)
        .open(path)
        .map_err(io_error)?;
    writeln!(file, "{content}").map_err(io_error)
}
fn lines(value: &str) -> usize {
    value.lines().filter(|line| !line.trim().is_empty()).count()
}
fn io_error(error: std::io::Error) -> String {
    error.to_string()
}

#[cfg(windows)]
unsafe fn call_shim(shim: &Path, mode: u32, fixture: Option<&Path>) -> Result<RawResult, String> {
    use std::os::windows::ffi::OsStrExt;
    use windows_sys::Win32::Foundation::FreeLibrary;
    use windows_sys::Win32::System::LibraryLoader::{GetProcAddress, LoadLibraryW};
    let wide: Vec<u16> = shim.as_os_str().encode_wide().chain(Some(0)).collect();
    let module = unsafe { LoadLibraryW(wide.as_ptr()) };
    if module.is_null() {
        return Err("LoadLibraryW for native shim failed".to_owned());
    }
    let procedure = unsafe { GetProcAddress(module, c"excel_raw_run".as_ptr() as *const u8) };
    if procedure.is_none() {
        unsafe {
            FreeLibrary(module);
        }
        return Err("excel_raw_run export not found".to_owned());
    }
    type Run = unsafe extern "system" fn(u32, *const u16, *mut RawResult) -> i32;
    let run: Run = unsafe {
        std::mem::transmute::<unsafe extern "system" fn() -> isize, Run>(procedure.unwrap())
    };
    let fixture_wide: Option<Vec<u16>> =
        fixture.map(|value| value.as_os_str().encode_wide().chain(Some(0)).collect());
    let mut result = RawResult::default();
    let status = unsafe {
        run(
            mode,
            fixture_wide
                .as_ref()
                .map_or(std::ptr::null(), |value| value.as_ptr()),
            &mut result,
        )
    };
    unsafe {
        FreeLibrary(module);
    }
    if status != 0 {
        return Err(format!("excel_raw_run returned {status}"));
    }
    Ok(result)
}

const NOT_TESTED: i32 = 0x7fff_fffe;

fn fresh_result() -> RawResult {
    RawResult {
        schema_version: SCHEMA_VERSION,
        activation_hresult: NOT_TESTED,
        version_hresult: NOT_TESTED,
        workbooks_hresult: NOT_TESTED,
        count_hresult: NOT_TESTED,
        add_hresult: NOT_TESTED,
        open_hresult: NOT_TESTED,
        quit_hresult: NOT_TESTED,
        inner_scode: 0,
        deferred_fill_in_hresult: NOT_TESTED,
        result_vt: 0,
        workbooks_vt: 0,
        pu_arg_err_raw: u32::MAX,
        workbook_created: 0,
        workbook_opened: 0,
        process_exited: NOT_TESTED,
        lifetime_clone_then_clear: NOT_TESTED,
        lifetime_retain_then_clear: NOT_TESTED,
        lifetime_query_interface_then_clear: NOT_TESTED,
        type_info_count: 0,
        type_info_hresult: NOT_TESTED,
        workbooks_query_iunknown_hresult: NOT_TESTED,
        workbooks_query_idispatch_hresult: NOT_TESTED,
    }
}

#[cfg(windows)]
unsafe fn high_level_run(mode: u32, fixture: Option<&Path>) -> RawResult {
    use std::mem::ManuallyDrop;
    use windows::Win32::System::Com::{
        CLSCTX_LOCAL_SERVER, CLSCTX_SERVER, COINIT_APARTMENTTHREADED, CoCreateInstance,
        CoCreateInstanceEx, CoInitializeEx, CoUninitialize, DISPATCH_METHOD, DISPATCH_PROPERTYGET,
        DISPPARAMS, EXCEPINFO, IDispatch, MULTI_QI,
    };
    use windows::Win32::System::Variant::{VARIANT, VT_BOOL, VT_BSTR, VT_DISPATCH, VariantClear};
    use windows::core::{BSTR, GUID, HSTRING, IUnknown, Interface, PCWSTR};

    struct VariantOwner(VARIANT);
    impl VariantOwner {
        fn empty() -> Self {
            Self(unsafe { windows::Win32::System::Variant::VariantInit() })
        }
        fn vt(&self) -> u16 {
            unsafe { self.0.Anonymous.Anonymous.vt.0 }
        }
        fn dispatch(&self) -> Option<IDispatch> {
            if self.vt() != VT_DISPATCH.0 {
                return None;
            }
            unsafe {
                let value = &self.0.Anonymous.Anonymous.Anonymous.pdispVal
                    as *const ManuallyDrop<Option<IDispatch>>
                    as *const Option<IDispatch>;
                (*value).clone()
            }
        }
        fn bstr(value: &str) -> Self {
            let mut result = Self::empty();
            unsafe {
                let data = &mut result.0.Anonymous.Anonymous;
                data.vt = VT_BSTR;
                data.Anonymous.bstrVal = ManuallyDrop::new(BSTR::from(value));
            }
            result
        }
        fn boolean(value: bool) -> Self {
            let mut result = Self::empty();
            unsafe {
                let data = &mut result.0.Anonymous.Anonymous;
                data.vt = VT_BOOL;
                data.Anonymous.boolVal =
                    windows::Win32::Foundation::VARIANT_BOOL(if value { -1 } else { 0 });
            }
            result
        }
        fn clear(&mut self) {
            let _ = unsafe { VariantClear(&mut self.0) };
        }
    }
    impl Drop for VariantOwner {
        fn drop(&mut self) {
            self.clear();
        }
    }

    fn hr(result: windows::core::Result<()>) -> i32 {
        result.map(|_| 0).unwrap_or_else(|error| error.code().0)
    }
    unsafe fn name(dispatch: &IDispatch, value: &str, lcid: u32) -> Result<i32, i32> {
        let value = HSTRING::from(value);
        let values = [PCWSTR(value.as_ptr())];
        let mut dispid = 0;
        unsafe {
            dispatch.GetIDsOfNames(&GUID::from_u128(0), values.as_ptr(), 1, lcid, &mut dispid)
        }
        .map(|_| dispid)
        .map_err(|error| error.code().0)
    }
    unsafe fn invoke(
        dispatch: &IDispatch,
        member: &str,
        flags: windows::Win32::System::Com::DISPATCH_FLAGS,
        lcid: u32,
        params: &DISPPARAMS,
        result: &mut VariantOwner,
        exception: &mut EXCEPINFO,
        arg_error: &mut u32,
    ) -> i32 {
        let dispid = match unsafe { name(dispatch, member, lcid) } {
            Ok(value) => value,
            Err(value) => return value,
        };
        hr(unsafe {
            dispatch.Invoke(
                dispid,
                &GUID::from_u128(0),
                lcid,
                flags,
                params,
                Some(&mut result.0),
                Some(exception),
                Some(arg_error),
            )
        })
    }
    unsafe fn get_dispatch(
        dispatch: &IDispatch,
        member: &str,
        lcid: u32,
        output_vt: &mut u16,
    ) -> Result<IDispatch, i32> {
        let empty = DISPPARAMS::default();
        let mut result = VariantOwner::empty();
        let mut exception = EXCEPINFO::default();
        let mut argument_error = u32::MAX;
        let status = unsafe {
            invoke(
                dispatch,
                member,
                DISPATCH_PROPERTYGET,
                lcid,
                &empty,
                &mut result,
                &mut exception,
                &mut argument_error,
            )
        };
        *output_vt = result.vt();
        if status != 0 {
            return Err(status);
        }
        result.dispatch().ok_or(0x8000_4005_u32 as i32)
    }
    unsafe fn close_workbook(workbook: &IDispatch, lcid: u32) -> i32 {
        let mut argument = VariantOwner::boolean(false);
        let params = DISPPARAMS {
            rgvarg: &mut argument.0,
            rgdispidNamedArgs: std::ptr::null_mut(),
            cArgs: 1,
            cNamedArgs: 0,
        };
        let mut result = VariantOwner::empty();
        let mut exception = EXCEPINFO::default();
        let mut argument_error = u32::MAX;
        unsafe {
            invoke(
                workbook,
                "Close",
                DISPATCH_METHOD,
                lcid,
                &params,
                &mut result,
                &mut exception,
                &mut argument_error,
            )
        }
    }
    unsafe fn count(dispatch: &IDispatch, lcid: u32) -> i32 {
        let empty = DISPPARAMS::default();
        let mut result = VariantOwner::empty();
        let mut exception = EXCEPINFO::default();
        let mut argument_error = u32::MAX;
        unsafe {
            invoke(
                dispatch,
                "Count",
                DISPATCH_PROPERTYGET,
                lcid,
                &empty,
                &mut result,
                &mut exception,
                &mut argument_error,
            )
        }
    }
    unsafe fn lifetime(app: &IDispatch, lcid: u32, sequence: u32) -> i32 {
        let mut vt = 0;
        let workbooks = match unsafe { get_dispatch(app, "Workbooks", lcid, &mut vt) } {
            Ok(value) => value,
            Err(value) => return value,
        };
        match sequence {
            1 => unsafe { count(&workbooks, lcid) },
            2 => {
                let retained = workbooks.clone();
                drop(workbooks);
                unsafe { count(&retained, lcid) }
            }
            _ => match workbooks
                .cast::<IUnknown>()
                .and_then(|value| value.cast::<IDispatch>())
            {
                Ok(value) => unsafe { count(&value, lcid) },
                Err(error) => error.code().0,
            },
        }
    }

    let mut result = fresh_result();
    if let Err(error) = unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok() } {
        result.activation_hresult = error.code().0;
        return result;
    }
    let program_id = HSTRING::from("Excel.Application");
    let class_id = match unsafe { windows::Win32::System::Com::CLSIDFromProgID(&program_id) } {
        Ok(value) => value,
        Err(error) => {
            result.activation_hresult = error.code().0;
            unsafe { CoUninitialize() };
            return result;
        }
    };
    let app = if mode == 3 {
        let iid = IDispatch::IID;
        let mut interfaces = [MULTI_QI {
            pIID: &iid,
            pItf: ManuallyDrop::new(None),
            hr: windows::core::HRESULT(0),
        }];
        let activation = unsafe {
            CoCreateInstanceEx(
                &class_id,
                None::<&IUnknown>,
                CLSCTX_SERVER,
                None,
                &mut interfaces,
            )
        };
        if let Err(error) = activation {
            result.activation_hresult = error.code().0;
            None
        } else if let Err(error) = interfaces[0].hr.ok() {
            result.activation_hresult = error.code().0;
            None
        } else {
            unsafe { ManuallyDrop::take(&mut interfaces[0].pItf) }
                .and_then(|value| value.cast::<IDispatch>().ok())
        }
    } else {
        let clsctx = if mode == 1 {
            CLSCTX_LOCAL_SERVER
        } else {
            CLSCTX_SERVER
        };
        match unsafe { CoCreateInstance::<_, IDispatch>(&class_id, None, clsctx) } {
            Ok(value) => Some(value),
            Err(error) => {
                result.activation_hresult = error.code().0;
                None
            }
        }
    };
    let Some(app) = app else {
        unsafe { CoUninitialize() };
        return result;
    };
    result.activation_hresult = 0;
    let lcid = if mode == 1 { 0x0400 } else { 0 };
    let empty = DISPPARAMS::default();
    let mut version = VariantOwner::empty();
    let mut version_exception = EXCEPINFO::default();
    let mut version_arg = u32::MAX;
    result.version_hresult = unsafe {
        invoke(
            &app,
            "Version",
            DISPATCH_PROPERTYGET,
            lcid,
            &empty,
            &mut version,
            &mut version_exception,
            &mut version_arg,
        )
    };
    let mut workbooks_vt = 0;
    let workbooks = unsafe { get_dispatch(&app, "Workbooks", lcid, &mut workbooks_vt) };
    result.workbooks_vt = workbooks_vt;
    if let Ok(workbooks) = workbooks {
        result.workbooks_hresult = 0;
        match unsafe { workbooks.GetTypeInfoCount() } {
            Ok(value) => {
                result.type_info_hresult = 0;
                result.type_info_count = value
            }
            Err(error) => result.type_info_hresult = error.code().0,
        }
        result.workbooks_query_iunknown_hresult = workbooks
            .cast::<IUnknown>()
            .map(|_| 0)
            .unwrap_or_else(|error| error.code().0);
        result.workbooks_query_idispatch_hresult = workbooks
            .cast::<IUnknown>()
            .and_then(|value| value.cast::<IDispatch>())
            .map(|_| 0)
            .unwrap_or_else(|error| error.code().0);
        result.count_hresult = unsafe { count(&workbooks, lcid) };
        result.lifetime_clone_then_clear = unsafe { lifetime(&app, lcid, 1) };
        result.lifetime_retain_then_clear = unsafe { lifetime(&app, lcid, 2) };
        result.lifetime_query_interface_then_clear = unsafe { lifetime(&app, lcid, 3) };
        let mut add = VariantOwner::empty();
        let mut add_exception = EXCEPINFO::default();
        let mut add_arg = u32::MAX;
        result.add_hresult = unsafe {
            invoke(
                &workbooks,
                "Add",
                DISPATCH_METHOD,
                lcid,
                &empty,
                &mut add,
                &mut add_exception,
                &mut add_arg,
            )
        };
        result.result_vt = add.vt();
        result.pu_arg_err_raw = add_arg;
        result.inner_scode = add_exception.scode;
        if result.add_hresult == 0 {
            if let Some(workbook) = add.dispatch() {
                result.workbook_created = (unsafe { close_workbook(&workbook, lcid) } == 0) as i32;
            }
        }
        if let Some(fixture) = fixture {
            let fixture_text = fixture.to_string_lossy();
            let mut argument = VariantOwner::bstr(&fixture_text);
            let open_params = DISPPARAMS {
                rgvarg: &mut argument.0,
                rgdispidNamedArgs: std::ptr::null_mut(),
                cArgs: 1,
                cNamedArgs: 0,
            };
            let mut open = VariantOwner::empty();
            let mut open_exception = EXCEPINFO::default();
            let mut open_arg = u32::MAX;
            result.open_hresult = unsafe {
                invoke(
                    &workbooks,
                    "Open",
                    DISPATCH_METHOD,
                    lcid,
                    &open_params,
                    &mut open,
                    &mut open_exception,
                    &mut open_arg,
                )
            };
            if result.open_hresult == 0 {
                if let Some(workbook) = open.dispatch() {
                    result.workbook_opened =
                        (unsafe { close_workbook(&workbook, lcid) } == 0) as i32;
                }
            }
        }
    } else if let Err(error) = workbooks {
        result.workbooks_hresult = error;
    }
    let mut quit = VariantOwner::empty();
    let mut quit_exception = EXCEPINFO::default();
    let mut quit_arg = u32::MAX;
    result.quit_hresult = unsafe {
        invoke(
            &app,
            "Quit",
            DISPATCH_METHOD,
            lcid,
            &empty,
            &mut quit,
            &mut quit_exception,
            &mut quit_arg,
        )
    };
    drop(app);
    unsafe { CoUninitialize() };
    result
}

#[cfg(windows)]
unsafe fn lower_level_run(mode: u32, fixture: Option<&Path>) -> RawResult {
    use std::ffi::c_void;
    use windows_sys::Win32::Foundation::{SysAllocString, SysFreeString};
    use windows_sys::Win32::System::Com::{
        CLSCTX_LOCAL_SERVER, CLSCTX_SERVER, CLSIDFromProgID, COINIT_APARTMENTTHREADED,
        CoCreateInstance, CoCreateInstanceEx, CoInitializeEx, CoUninitialize, DISPPARAMS,
        EXCEPINFO, MULTI_QI,
    };
    use windows_sys::Win32::System::Variant::{
        VARIANT, VT_BOOL, VT_BSTR, VT_DISPATCH, VariantClear, VariantInit,
    };
    use windows_sys::core::{GUID, HRESULT, IID_IUnknown, IUnknown_Vtbl};

    const IID_IDISPATCH: GUID = GUID::from_u128(0x00020400_0000_0000_c000_000000000046);
    const DISPATCH_METHOD: u16 = 1;
    const DISPATCH_PROPERTYGET: u16 = 2;

    #[repr(C)]
    struct GenericIDispatchVtbl {
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

    /// Owning, generic Automation dispatch pointer. Its vtable declaration is
    /// the SDK IDispatch order, not an Excel dual interface declaration.
    struct RawDispatch(*mut c_void);
    impl RawDispatch {
        unsafe fn vtable(&self) -> &GenericIDispatchVtbl {
            unsafe { &**(self.0 as *const *const GenericIDispatchVtbl) }
        }
        unsafe fn add_ref(&self) -> u32 {
            unsafe { (self.vtable().base.AddRef)(self.0) }
        }
        unsafe fn query(&self, iid: &GUID) -> Result<Self, HRESULT> {
            let mut output = std::ptr::null_mut();
            let status = unsafe { (self.vtable().base.QueryInterface)(self.0, iid, &mut output) };
            if status == 0 && !output.is_null() {
                Ok(Self(output))
            } else {
                Err(status)
            }
        }
        unsafe fn type_info_count(&self) -> Result<u32, HRESULT> {
            let mut count = 0;
            let status = unsafe { (self.vtable().get_type_info_count)(self.0, &mut count) };
            if status == 0 { Ok(count) } else { Err(status) }
        }
        unsafe fn dispid(&self, member: &str, lcid: u32) -> Result<i32, HRESULT> {
            let wide: Vec<u16> = member.encode_utf16().chain(Some(0)).collect();
            let names = [wide.as_ptr()];
            let mut id = 0;
            let status = unsafe {
                (self.vtable().get_ids_of_names)(
                    self.0,
                    &GUID::default(),
                    names.as_ptr(),
                    1,
                    lcid,
                    &mut id,
                )
            };
            if status == 0 { Ok(id) } else { Err(status) }
        }
        unsafe fn invoke(
            &self,
            member: &str,
            flags: u16,
            lcid: u32,
            params: &DISPPARAMS,
            output: &mut SysVariant,
            exception: &mut EXCEPINFO,
            argument_error: &mut u32,
        ) -> HRESULT {
            let id = match unsafe { self.dispid(member, lcid) } {
                Ok(value) => value,
                Err(value) => return value,
            };
            unsafe {
                (self.vtable().invoke)(
                    self.0,
                    id,
                    &GUID::default(),
                    lcid,
                    flags,
                    params,
                    &mut output.0,
                    exception,
                    argument_error,
                )
            }
        }
    }
    impl Clone for RawDispatch {
        fn clone(&self) -> Self {
            unsafe {
                self.add_ref();
            }
            Self(self.0)
        }
    }
    impl Drop for RawDispatch {
        fn drop(&mut self) {
            if !self.0.is_null() {
                unsafe {
                    (self.vtable().base.Release)(self.0);
                }
            }
        }
    }

    struct SysVariant(VARIANT);
    impl SysVariant {
        fn empty() -> Self {
            let mut value = VARIANT::default();
            unsafe { VariantInit(&mut value) };
            Self(value)
        }
        fn vt(&self) -> u16 {
            unsafe { self.0.Anonymous.Anonymous.vt }
        }
        fn dispatch(&self) -> Option<RawDispatch> {
            if self.vt() != VT_DISPATCH {
                return None;
            }
            let pointer = unsafe { self.0.Anonymous.Anonymous.Anonymous.pdispVal };
            if pointer.is_null() {
                None
            } else {
                // The result VARIANT owns the existing reference. Add one for
                // the returned RawDispatch without constructing a temporary
                // owning wrapper that would immediately Release it.
                unsafe {
                    let vtable = &**(pointer as *const *const GenericIDispatchVtbl);
                    (vtable.base.AddRef)(pointer);
                }
                Some(RawDispatch(pointer))
            }
        }
        fn bstr(value: &Path) -> Self {
            let wide: Vec<u16> = value
                .as_os_str()
                .to_string_lossy()
                .encode_utf16()
                .chain(Some(0))
                .collect();
            let mut result = Self::empty();
            result.0.Anonymous.Anonymous.vt = VT_BSTR;
            result.0.Anonymous.Anonymous.Anonymous.bstrVal =
                unsafe { SysAllocString(wide.as_ptr()) };
            result
        }
        fn boolean(value: bool) -> Self {
            let mut result = Self::empty();
            result.0.Anonymous.Anonymous.vt = VT_BOOL;
            result.0.Anonymous.Anonymous.Anonymous.boolVal = if value { -1 } else { 0 };
            result
        }
    }
    impl Drop for SysVariant {
        fn drop(&mut self) {
            unsafe {
                let _ = VariantClear(&mut self.0);
            }
        }
    }

    unsafe fn clear_exception(exception: &mut EXCEPINFO) {
        unsafe {
            if let Some(fill) = exception.pfnDeferredFillIn {
                let _ = fill(exception);
            }
            if !exception.bstrSource.is_null() {
                SysFreeString(exception.bstrSource);
                exception.bstrSource = std::ptr::null();
            }
            if !exception.bstrDescription.is_null() {
                SysFreeString(exception.bstrDescription);
                exception.bstrDescription = std::ptr::null();
            }
            if !exception.bstrHelpFile.is_null() {
                SysFreeString(exception.bstrHelpFile);
                exception.bstrHelpFile = std::ptr::null();
            }
        }
    }
    unsafe fn get_dispatch(
        dispatch: &RawDispatch,
        member: &str,
        lcid: u32,
        output_vt: &mut u16,
    ) -> Result<RawDispatch, HRESULT> {
        let params = DISPPARAMS::default();
        let mut output = SysVariant::empty();
        let mut exception = EXCEPINFO::default();
        let mut arg_error = u32::MAX;
        let status = unsafe {
            dispatch.invoke(
                member,
                DISPATCH_PROPERTYGET,
                lcid,
                &params,
                &mut output,
                &mut exception,
                &mut arg_error,
            )
        };
        *output_vt = output.vt();
        unsafe { clear_exception(&mut exception) };
        if status != 0 {
            return Err(status);
        }
        output.dispatch().ok_or(0x8000_4005_u32 as i32)
    }
    unsafe fn basic_invoke(
        dispatch: &RawDispatch,
        member: &str,
        flags: u16,
        lcid: u32,
        params: &DISPPARAMS,
    ) -> (HRESULT, SysVariant, EXCEPINFO, u32) {
        let mut output = SysVariant::empty();
        let mut exception = EXCEPINFO::default();
        let mut arg_error = u32::MAX;
        let status = unsafe {
            dispatch.invoke(
                member,
                flags,
                lcid,
                params,
                &mut output,
                &mut exception,
                &mut arg_error,
            )
        };
        (status, output, exception, arg_error)
    }
    unsafe fn close_workbook(dispatch: &RawDispatch, lcid: u32) -> HRESULT {
        let mut argument = SysVariant::boolean(false);
        let params = DISPPARAMS {
            rgvarg: &mut argument.0,
            rgdispidNamedArgs: std::ptr::null_mut(),
            cArgs: 1,
            cNamedArgs: 0,
        };
        let (status, _, mut exception, _) =
            unsafe { basic_invoke(dispatch, "Close", DISPATCH_METHOD, lcid, &params) };
        unsafe { clear_exception(&mut exception) };
        status
    }
    unsafe fn count(dispatch: &RawDispatch, lcid: u32) -> HRESULT {
        let params = DISPPARAMS::default();
        let (status, _, mut exception, _) =
            unsafe { basic_invoke(dispatch, "Count", DISPATCH_PROPERTYGET, lcid, &params) };
        unsafe { clear_exception(&mut exception) };
        status
    }
    unsafe fn lifetime(app: &RawDispatch, lcid: u32, sequence: u32) -> HRESULT {
        let mut vt = 0;
        let workbooks = match unsafe { get_dispatch(app, "Workbooks", lcid, &mut vt) } {
            Ok(value) => value,
            Err(value) => return value,
        };
        match sequence {
            1 => unsafe { count(&workbooks, lcid) },
            2 => {
                let retained = workbooks.clone();
                drop(workbooks);
                unsafe { count(&retained, lcid) }
            }
            _ => match unsafe { workbooks.query(&IID_IDISPATCH) } {
                Ok(value) => unsafe { count(&value, lcid) },
                Err(value) => value,
            },
        }
    }

    let mut result = fresh_result();
    let init = unsafe { CoInitializeEx(std::ptr::null(), COINIT_APARTMENTTHREADED as u32) };
    if init != 0 {
        result.activation_hresult = init;
        return result;
    }
    let program_id: Vec<u16> = "Excel.Application".encode_utf16().chain(Some(0)).collect();
    let mut class_id = GUID::default();
    let clsid = unsafe { CLSIDFromProgID(program_id.as_ptr(), &mut class_id) };
    if clsid != 0 {
        result.activation_hresult = clsid;
        unsafe { CoUninitialize() };
        return result;
    }
    let mut raw = std::ptr::null_mut();
    let activation = if mode == 3 {
        let mut request = MULTI_QI {
            pIID: &IID_IDISPATCH,
            pItf: std::ptr::null_mut(),
            hr: 0,
        };
        let status = unsafe {
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
        if status == 0 { request.hr } else { status }
    } else {
        let context = if mode == 1 {
            CLSCTX_LOCAL_SERVER
        } else {
            CLSCTX_SERVER
        };
        unsafe {
            CoCreateInstance(
                &class_id,
                std::ptr::null_mut(),
                context,
                &IID_IDISPATCH,
                &mut raw,
            )
        }
    };
    result.activation_hresult = activation;
    if activation != 0 || raw.is_null() {
        unsafe { CoUninitialize() };
        return result;
    }
    let app = RawDispatch(raw);
    let lcid = if mode == 1 { 0x0400 } else { 0 };
    let empty = DISPPARAMS::default();
    let (version, _, mut version_exception, _) =
        unsafe { basic_invoke(&app, "Version", DISPATCH_PROPERTYGET, lcid, &empty) };
    result.version_hresult = version;
    unsafe { clear_exception(&mut version_exception) };
    let mut workbooks_vt = 0;
    match unsafe { get_dispatch(&app, "Workbooks", lcid, &mut workbooks_vt) } {
        Ok(workbooks) => {
            result.workbooks_hresult = 0;
            result.workbooks_vt = workbooks_vt;
            match unsafe { workbooks.type_info_count() } {
                Ok(value) => {
                    result.type_info_hresult = 0;
                    result.type_info_count = value
                }
                Err(error) => result.type_info_hresult = error,
            };
            result.workbooks_query_iunknown_hresult = unsafe { workbooks.query(&IID_IUnknown) }
                .map(|_| 0)
                .unwrap_or_else(|error| error);
            result.workbooks_query_idispatch_hresult = unsafe { workbooks.query(&IID_IDISPATCH) }
                .map(|_| 0)
                .unwrap_or_else(|error| error);
            result.count_hresult = unsafe { count(&workbooks, lcid) };
            result.lifetime_clone_then_clear = unsafe { lifetime(&app, lcid, 1) };
            result.lifetime_retain_then_clear = unsafe { lifetime(&app, lcid, 2) };
            result.lifetime_query_interface_then_clear = unsafe { lifetime(&app, lcid, 3) };
            let (add, output, mut exception, arg_error) =
                unsafe { basic_invoke(&workbooks, "Add", DISPATCH_METHOD, lcid, &empty) };
            result.add_hresult = add;
            result.result_vt = output.vt();
            result.pu_arg_err_raw = arg_error;
            result.inner_scode = exception.scode;
            unsafe { clear_exception(&mut exception) };
            if add == 0 {
                if let Some(workbook) = output.dispatch() {
                    result.workbook_created =
                        (unsafe { close_workbook(&workbook, lcid) } == 0) as i32;
                }
            }
            if let Some(fixture) = fixture {
                let mut argument = SysVariant::bstr(fixture);
                let params = DISPPARAMS {
                    rgvarg: &mut argument.0,
                    rgdispidNamedArgs: std::ptr::null_mut(),
                    cArgs: 1,
                    cNamedArgs: 0,
                };
                let (open, output, mut exception, _) =
                    unsafe { basic_invoke(&workbooks, "Open", DISPATCH_METHOD, lcid, &params) };
                result.open_hresult = open;
                unsafe { clear_exception(&mut exception) };
                if open == 0 {
                    if let Some(workbook) = output.dispatch() {
                        result.workbook_opened =
                            (unsafe { close_workbook(&workbook, lcid) } == 0) as i32;
                    }
                }
            }
        }
        Err(error) => {
            result.workbooks_hresult = error;
            result.workbooks_vt = workbooks_vt;
        }
    }
    let (quit, _, mut quit_exception, _) =
        unsafe { basic_invoke(&app, "Quit", DISPATCH_METHOD, lcid, &empty) };
    result.quit_hresult = quit;
    unsafe { clear_exception(&mut quit_exception) };
    drop(app);
    unsafe { CoUninitialize() };
    result
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn fixed_width_result_layout_is_stable() {
        assert_eq!(std::mem::size_of::<RawResult>(), 88);
        assert_eq!(std::mem::align_of::<RawResult>(), 4);
    }

    #[test]
    fn required_modes_are_exact() {
        assert_eq!(mode_number("native-cocreate-local-lcid-0400"), Ok(1));
        assert_eq!(mode_number("native-cocreate-server-lcid-0000"), Ok(2));
        assert_eq!(mode_number("native-cocreateex-server-lcid-0000"), Ok(3));
        assert!(mode_number("other").is_err());
    }

    #[test]
    fn operation_specs_prohibit_pointer_persistence() {
        let specs = operation_specs();
        assert!(specs.contains("workbooks-add"));
        assert!(specs.contains("null rgvarg"));
        assert!(!specs.contains("C:\\\\"));
    }

    #[test]
    fn version_matrix_is_pinned_and_deterministic() {
        let matrix = version_matrix();
        assert!(matrix.contains("\"windows\":\"0.62.2\""));
        assert!(matrix.contains("\"windows\":\"0.62.1\""));
        assert_eq!(matrix, version_matrix());
    }
}
