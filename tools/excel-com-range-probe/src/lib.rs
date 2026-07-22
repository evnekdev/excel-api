//! Isolated evidence capture for Excel Range Automation transport.
//!
//! This unpublished tool is deliberately not a reusable Excel client.  It
//! creates one local Excel instance for a bounded experiment, records copied
//! metadata, and then closes only that workbook and that instance.

use serde::Serialize;
use serde_json::{Value, json};
use std::collections::BTreeMap;
use std::fs;
use std::path::{Path, PathBuf};

pub mod raw;
pub mod python_differential;
pub(crate) mod automation;

/// Validates Prompt 06's internal Automation-value evidence without opening
/// Excel. The semantic types themselves remain crate-internal.
pub fn automation_value_check(root: &Path) -> Result<(), String> {
    automation::check_evidence(root)
}

/// Runs Prompt 06's explicit, opt-in L-mode compatibility suite. It refuses
/// to start while any Excel process is already present.
pub fn automation_value_live(root: &Path, only_case: Option<&str>) -> Result<String, String> {
    automation::live_compatibility(root, only_case)
}

/// Runs the semantic live cases without modifying the Prompt 06 evidence
/// tree. This lets later research prompts retain their own observations.
pub fn automation_value_live_observations(
    only_case: Option<&str>,
) -> Result<(Vec<Value>, Vec<&'static str>), String> {
    automation::live_compatibility_observations(only_case)
}

const PROBE_VERSION: u32 = 2;
const MOJIBAKE_PATTERNS: &[&str] = &["â", "ï¿½", "\u{FFFD}"];

/// Research-only configurations used by Prompt 05D. They document a bounded
/// source-derived comparison; they are not a production Automation API.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize)]
#[serde(rename_all = "kebab-case")]
pub enum ParityMode {
    RustBaseline,
    Pywin32Dynamic,
    Pywin32Generated,
    ComtypesDynamic,
    ComtypesGenerated,
}

impl ParityMode {
    const ALL: [Self; 5] = [
        Self::RustBaseline,
        Self::Pywin32Dynamic,
        Self::Pywin32Generated,
        Self::ComtypesDynamic,
        Self::ComtypesGenerated,
    ];

    pub fn parse(value: &str) -> Result<Self, String> {
        match value {
            "rust-baseline" => Ok(Self::RustBaseline),
            "pywin32-dynamic" => Ok(Self::Pywin32Dynamic),
            "pywin32-generated" => Ok(Self::Pywin32Generated),
            "comtypes-dynamic" => Ok(Self::ComtypesDynamic),
            "comtypes-generated" => Ok(Self::ComtypesGenerated),
            _ => Err("parity mode must be rust-baseline, pywin32-dynamic, pywin32-generated, comtypes-dynamic, or comtypes-generated".to_owned()),
        }
    }

    pub fn id(self) -> &'static str {
        match self {
            Self::RustBaseline => "rust-baseline",
            Self::Pywin32Dynamic => "pywin32-dynamic",
            Self::Pywin32Generated => "pywin32-generated",
            Self::ComtypesDynamic => "comtypes-dynamic",
            Self::ComtypesGenerated => "comtypes-generated",
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize)]
pub struct ParityConfiguration {
    pub mode: ParityMode,
    pub activation_api: &'static str,
    pub clsctx: &'static str,
    pub requested_iid: &'static str,
    pub com_initialization: &'static str,
    pub get_ids_of_names_lcid: u32,
    pub invoke_lcid: u32,
    pub dispid_source: &'static str,
    pub invoke_strategy: &'static str,
    pub argument_omission_policy: &'static str,
    pub pu_arg_err_initialization: &'static str,
    pub result_requested: bool,
    pub type_info_probing: &'static str,
    pub dual_interface_handling: &'static str,
    pub returned_object_wrapping_policy: &'static str,
    pub source_justification: &'static str,
}

pub fn parity_configuration(mode: ParityMode) -> ParityConfiguration {
    match mode {
        ParityMode::RustBaseline => ParityConfiguration {
            mode,
            activation_api: "CoCreateInstance",
            clsctx: "CLSCTX_LOCAL_SERVER",
            requested_iid: "IID_IDispatch",
            com_initialization: "CoInitializeEx(COINIT_APARTMENTTHREADED)",
            get_ids_of_names_lcid: 0x0400,
            invoke_lcid: 0x0400,
            dispid_source: "installed typelib DISPID plus GetIDsOfNames verification",
            invoke_strategy: "raw IDispatch::Invoke",
            argument_omission_policy: "zero omitted arguments use cArgs=0 with null argument pointers",
            pu_arg_err_initialization: "UINT_MAX sentinel",
            result_requested: true,
            type_info_probing: "record type-info availability only",
            dual_interface_handling: "not selected; raw IDispatch path",
            returned_object_wrapping_policy: "clone IDispatch from VT_DISPATCH before VariantClear",
            source_justification: "preserved Prompt 05B baseline for comparison",
        },
        ParityMode::Pywin32Dynamic => ParityConfiguration {
            mode,
            activation_api: "CoCreateInstanceEx",
            clsctx: "CLSCTX_SERVER",
            requested_iid: "IID_IDispatch",
            com_initialization: "CoInitializeEx(COINIT_APARTMENTTHREADED)",
            get_ids_of_names_lcid: 0,
            invoke_lcid: 0,
            dispid_source: "GetIDsOfNames(LCID=0) verified against installed typelib",
            invoke_strategy: "raw IDispatch::Invoke; pywin32 CDispatch can select typed helpers when descriptors are available",
            argument_omission_policy: "zero omitted arguments use cArgs=0 with null argument pointers",
            pu_arg_err_initialization: "UINT_MAX sentinel",
            result_requested: true,
            type_info_probing: "GetTypeInfoCount/GetTypeInfo recorded; dynamic wrapper mode remains distinct from typed helper selection",
            dual_interface_handling: "not selected by the raw dynamic parity call",
            returned_object_wrapping_policy: "clone IDispatch from VT_DISPATCH before VariantClear",
            source_justification: "pywin32 311 b311 DispatchEx, dynamic.CDispatch, and PyIDispatch source reconciliation",
        },
        ParityMode::Pywin32Generated => ParityConfiguration {
            mode,
            activation_api: "CoCreateInstance",
            clsctx: "CLSCTX_SERVER",
            requested_iid: "IID_IDispatch",
            com_initialization: "CoInitializeEx(COINIT_APARTMENTTHREADED)",
            get_ids_of_names_lcid: 0,
            invoke_lcid: 0,
            dispid_source: "installed typelib descriptors plus GetIDsOfNames(LCID=0) verification",
            invoke_strategy: "raw IDispatch::Invoke with generated-descriptor parity recorded; no hand-emitted InvokeTypes ABI",
            argument_omission_policy: "trailing optional omission is distinct from VT_ERROR/DISP_E_PARAMNOTFOUND",
            pu_arg_err_initialization: "UINT_MAX sentinel",
            result_requested: true,
            type_info_probing: "GetTypeInfoCount/GetTypeInfo recorded before raw fallback",
            dual_interface_handling: "not selected; no generated Rust wrapper is introduced",
            returned_object_wrapping_policy: "clone IDispatch from VT_DISPATCH before VariantClear",
            source_justification: "pywin32 makepy InvokeTypes and generated-property source evidence; bounded Rust fallback avoids inventing descriptors",
        },
        ParityMode::ComtypesDynamic => ParityConfiguration {
            mode,
            activation_api: "CoCreateInstance",
            clsctx: "CLSCTX_SERVER",
            requested_iid: "IID_IDispatch",
            com_initialization: "CoInitializeEx(COINIT_APARTMENTTHREADED)",
            get_ids_of_names_lcid: 0,
            invoke_lcid: 0,
            dispid_source: "GetIDsOfNames(LCID=0) verified against installed typelib",
            invoke_strategy: "raw IDispatch::Invoke matching the terminal lazybind _invoke call",
            argument_omission_policy: "zero omitted arguments use cArgs=0 with null argument pointers",
            pu_arg_err_initialization: "UINT_MAX sentinel",
            result_requested: true,
            type_info_probing: "GetTypeInfoCount/GetTypeInfo recorded; comtypes lazybind ITypeComp binding is not reimplemented",
            dual_interface_handling: "not selected by the raw dynamic parity call",
            returned_object_wrapping_policy: "clone IDispatch from VT_DISPATCH before VariantClear",
            source_justification: "comtypes 1.4.16 lazybind and automation._invoke source evidence",
        },
        ParityMode::ComtypesGenerated => ParityConfiguration {
            mode,
            activation_api: "CoCreateInstance",
            clsctx: "CLSCTX_SERVER",
            requested_iid: "IID_IDispatch",
            com_initialization: "CoInitializeEx(COINIT_APARTMENTTHREADED)",
            get_ids_of_names_lcid: 0,
            invoke_lcid: 0,
            dispid_source: "installed typelib descriptors plus GetIDsOfNames(LCID=0) verification",
            invoke_strategy: "raw IDispatch fallback; generated vtable call is deliberately not hand-written",
            argument_omission_policy: "zero omitted arguments use cArgs=0 with null argument pointers",
            pu_arg_err_initialization: "UINT_MAX sentinel",
            result_requested: true,
            type_info_probing: "GetTypeInfoCount/GetTypeInfo and dual-interface availability recorded",
            dual_interface_handling: "blocked: no generated Rust Excel bindings are available, so no vtable layout is improvised",
            returned_object_wrapping_policy: "clone IDispatch from VT_DISPATCH before VariantClear",
            source_justification: "comtypes 1.4.16 tlbparser dual-interface source evidence and Prompt 04 typelib metadata",
        },
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct LiveSummary {
    pub observations: usize,
    pub completed_cases: usize,
    pub inconclusive_cases: usize,
}

pub fn live(root: &Path, control_script: Option<&Path>) -> Result<LiveSummary, String> {
    diagnose(root, control_script)
}

/// Run the opt-in Prompt 05B diagnostic. Existing Prompt 05 evidence is
/// merged by record ID so the historical blocked observation is never erased.
pub fn diagnose(root: &Path, control_script: Option<&Path>) -> Result<LiveSummary, String> {
    #[cfg(windows)]
    {
        let fresh = windows_live::capture(root, control_script)?;
        let summary = LiveSummary {
            observations: fresh.observations.len(),
            completed_cases: fresh
                .cases
                .iter()
                .filter(|case| value_string(case, "status") == "completed")
                .count(),
            inconclusive_cases: fresh
                .cases
                .iter()
                .filter(|case| value_string(case, "status") != "completed")
                .count(),
        };
        let capture = merge_capture(read_capture(root)?, fresh);
        write_capture(root, &capture)?;
        Ok(summary)
    }
    #[cfg(not(windows))]
    {
        let _ = (root, control_script);
        Err(
            "live Range probing requires Windows and a locally installed Excel Automation server"
                .to_owned(),
        )
    }
}

/// Run exactly one fresh, source-derived Prompt 05D parity configuration.
/// Repeating a configuration requires a different run ID so prior evidence is
/// never overwritten by the record-merging layer.
pub fn parity(
    root: &Path,
    mode: ParityMode,
    fixture: Option<&Path>,
    run_id: &str,
) -> Result<LiveSummary, String> {
    #[cfg(windows)]
    {
        let fresh = windows_live::capture_parity(root, mode, fixture, run_id)?;
        let summary = LiveSummary {
            observations: fresh.observations.len(),
            completed_cases: fresh
                .cases
                .iter()
                .filter(|case| value_string(case, "status") == "completed")
                .count(),
            inconclusive_cases: fresh
                .cases
                .iter()
                .filter(|case| value_string(case, "status") != "completed")
                .count(),
        };
        let capture = merge_capture(read_capture(root)?, fresh);
        write_capture(root, &capture)?;
        Ok(summary)
    }
    #[cfg(not(windows))]
    {
        let _ = (root, mode, fixture, run_id);
        Err("Prompt 05D parity probing requires Windows and a locally installed Excel Automation server".to_owned())
    }
}

/// Execute the pre-existing high-level `windows` control without writing to
/// the historical Prompt 05 runtime corpus. Prompt 05G persists only this
/// copied summary in its separate evidence tree.
#[cfg(windows)]
pub(crate) fn high_level_kernel_control(
    knowledge_root: &Path,
    fixture: Option<&Path>,
    run_id: &str,
) -> Result<Value, String> {
    let capture = windows_live::capture_parity(
        knowledge_root,
        ParityMode::RustBaseline,
        fixture,
        run_id,
    )?;
    let operation = |name| {
        capture
            .observations
            .iter()
            .find(|record| value_string(record, "parity_operation") == name)
            .cloned()
            .unwrap_or(Value::Null)
    };
    Ok(json!({
        "backend": "high-level-windows",
        "implementation": "existing Prompt 05 high-level windows crate diagnostic, run without persistence",
        "workbooks_add": operation("workbooks-add"),
        "workbooks_open": operation("workbooks-open"),
        "range_smoke": operation("range-smoke"),
        "completed_cases": capture.cases.iter().filter(|case| value_string(case, "status") == "completed").count(),
        "inconclusive_cases": capture.cases.iter().filter(|case| value_string(case, "status") != "completed").count(),
        "raw_pointer_values_recorded": false,
    }))
}

pub fn check(root: &Path) -> Result<(), String> {
    let capture = read_capture(root)?;
    for (relative, expected) in artifacts(&capture)? {
        reject_mojibake(&expected, &relative)?;
        let path = root.join(relative);
        let actual = fs::read_to_string(&path)
            .map_err(|error| format!("cannot read {}: {error}", path.display()))?;
        if actual != expected {
            return Err(format!(
                "runtime artifact {} is stale; rerun the explicit live probe",
                path.display()
            ));
        }
        if actual.contains("\r\n") || !actual.ends_with('\n') {
            return Err(format!(
                "runtime artifact {} must have LF endings and a final newline",
                path.display()
            ));
        }
    }
    Ok(())
}

fn reject_mojibake(text: &str, path: &Path) -> Result<(), String> {
    if let Some(pattern) = MOJIBAKE_PATTERNS
        .iter()
        .find(|pattern| text.contains(**pattern))
    {
        return Err(format!(
            "runtime artifact {} contains mojibake pattern {pattern:?}",
            path.display()
        ));
    }
    Ok(())
}

/// Rebuild deterministic reports from the committed runtime evidence without
/// opening Excel or changing the evidence records themselves.
pub fn refresh(root: &Path) -> Result<(), String> {
    let capture = read_capture(root)?;
    write_capture(root, &capture)
}

#[derive(Debug, Clone)]
struct Capture {
    manifest: String,
    environments: Vec<Value>,
    observations: Vec<Value>,
    cases: Vec<Value>,
    unresolved: Vec<Value>,
}

fn write_capture(root: &Path, capture: &Capture) -> Result<(), String> {
    for (relative, contents) in artifacts(capture)? {
        let path = root.join(relative);
        let parent = path
            .parent()
            .ok_or_else(|| format!("{} has no parent", path.display()))?;
        fs::create_dir_all(parent)
            .map_err(|error| format!("cannot create {}: {error}", parent.display()))?;
        fs::write(&path, contents)
            .map_err(|error| format!("cannot write {}: {error}", path.display()))?;
    }
    Ok(())
}

fn read_capture(root: &Path) -> Result<Capture, String> {
    let runtime = root.join("runtime");
    Ok(Capture {
        manifest: fs::read_to_string(runtime.join("SOURCE_MANIFEST.toml"))
            .map_err(|error| format!("cannot read runtime source manifest: {error}"))?,
        environments: read_jsonl(&runtime.join("environments.jsonl"))?,
        observations: read_jsonl(&runtime.join("observations.jsonl"))?,
        cases: read_jsonl(&runtime.join("cases.jsonl"))?,
        unresolved: read_jsonl(&runtime.join("unresolved.jsonl"))?,
    })
}

fn read_jsonl(path: &Path) -> Result<Vec<Value>, String> {
    let text = fs::read_to_string(path)
        .map_err(|error| format!("cannot read {}: {error}", path.display()))?;
    let mut values = Vec::new();
    for line in text.lines().filter(|line| !line.is_empty()) {
        values.push(
            serde_json::from_str(line)
                .map_err(|error| format!("cannot decode {}: {error}", path.display()))?,
        );
    }
    sort_records(&mut values);
    Ok(values)
}

fn artifacts(capture: &Capture) -> Result<BTreeMap<PathBuf, String>, String> {
    let mut files = BTreeMap::new();
    files.insert(
        PathBuf::from("runtime/SOURCE_MANIFEST.toml"),
        capture.manifest.clone(),
    );
    files.insert(
        PathBuf::from("runtime/environments.jsonl"),
        jsonl(&capture.environments)?,
    );
    files.insert(
        PathBuf::from("runtime/observations.jsonl"),
        jsonl(&capture.observations)?,
    );
    files.insert(PathBuf::from("runtime/cases.jsonl"), jsonl(&capture.cases)?);
    files.insert(
        PathBuf::from("runtime/unresolved.jsonl"),
        jsonl(&capture.unresolved)?,
    );
    for (name, report) in reports(capture) {
        files.insert(PathBuf::from("generated/runtime").join(name), report);
    }
    Ok(files)
}

fn jsonl(records: &[Value]) -> Result<String, String> {
    let mut records = records.to_vec();
    sort_records(&mut records);
    let mut output = String::new();
    for record in records {
        output.push_str(
            &serde_json::to_string(&record)
                .map_err(|error| format!("cannot encode JSONL: {error}"))?,
        );
        output.push('\n');
    }
    if output.is_empty() {
        output.push('\n');
    }
    Ok(output)
}

fn sort_records(records: &mut [Value]) {
    records.sort_by_key(|record| value_string(record, "id"));
}

fn reports(capture: &Capture) -> BTreeMap<&'static str, String> {
    let mut output = BTreeMap::new();
    output.insert(
        "range-shapes.md",
        observation_report(capture, "Range shapes", "array"),
    );
    output.insert(
        "scalar-values.md",
        observation_report(capture, "Scalar Value and Value2", "scalar"),
    );
    output.insert(
        "rectangular-values.md",
        observation_report(capture, "Rectangular range values", "rectangular"),
    );
    output.insert(
        "value-vs-value2.md",
        observation_report(capture, "Value versus Value2", "value-vs-value2"),
    );
    output.insert(
        "formulas.md",
        observation_report(capture, "Formula and Formula2", "formula"),
    );
    output.insert(
        "excel-errors.md",
        observation_report(capture, "Excel error cells", "error"),
    );
    output.insert(
        "writes.md",
        observation_report(capture, "Scalar and rectangular writes", "write"),
    );
    output.insert(
        "optional-arguments.md",
        observation_report(capture, "Optional Range arguments", "optional"),
    );
    output.insert(
        "workbooks-add-diagnostic.md",
        workbooks_add_diagnostic_report(capture),
    );
    output.insert("invocation-frames.md", invocation_frames_report(capture));
    output.insert("control-comparison.md", control_comparison_report(capture));
    output.insert("controls.md", controls_report(capture));
    output.insert("unresolved.md", unresolved_report(capture));
    output.insert(
        "environment-stability-matrix.md",
        parity_environment_report(capture),
    );
    output.insert(
        "workbooks-add-parity-matrix.md",
        parity_operation_report(capture, "Workbooks.Add parity matrix", "workbooks-add"),
    );
    output.insert(
        "workbook-open-parity-matrix.md",
        parity_operation_report(capture, "Workbooks.Open parity matrix", "workbooks-open"),
    );
    output.insert("range-smoke-test.md", parity_smoke_report(capture));
    output.insert("parity-cleanup.md", parity_cleanup_report(capture));
    output.insert("remaining-blockers.md", parity_blockers_report(capture));
    output
}

fn report_header(title: &str, subtitle: &str) -> String {
    format!(
        "# {title}\n\nGenerated by `excel-com-range-probe`; do not edit by hand. {subtitle}\n\nRuntime observations are version-specific and do not define a public API.\n\n"
    )
}

fn observation_report(capture: &Capture, title: &str, category: &str) -> String {
    let mut output = report_header(
        title,
        "Rows preserve copied raw observation metadata from the owned live Excel instance.",
    );
    output.push_str("| Case | Member | Range | HRESULT | Returned VARTYPE | Array shape / bounds | Classification |\n| --- | --- | --- | --- | --- | --- | --- |\n");
    let mut rows = capture
        .observations
        .iter()
        .filter(|observation| {
            observation
                .get("categories")
                .and_then(Value::as_array)
                .is_some_and(|categories| {
                    categories
                        .iter()
                        .any(|entry| entry.as_str() == Some(category))
                })
        })
        .collect::<Vec<_>>();
    rows.sort_by_key(|row| value_string(row, "id"));
    if rows.is_empty() {
        output.push_str("| -- | -- | -- | -- | No observation captured. | -- | Not tested |\n");
    }
    for row in rows {
        let result = row.get("result").unwrap_or(&Value::Null);
        let array = result
            .get("array")
            .map_or_else(|| "scalar".to_owned(), array_label);
        output.push_str(&format!(
            "| `{}` | `{}` | `{}` | `{}` | `{}` | {} | {} |\n",
            value_string(row, "case_id"),
            value_string(row, "subject"),
            value_string(row, "range"),
            value_string(result, "hresult"),
            value_string(result, "variant_type"),
            array,
            value_string(row, "classification")
        ));
    }
    output
}

fn array_label(value: &Value) -> String {
    let dimensions = value
        .get("dimensions")
        .and_then(Value::as_u64)
        .unwrap_or_default();
    let bounds = value
        .get("bounds")
        .and_then(Value::as_array)
        .map(|bounds| {
            bounds
                .iter()
                .map(|bound| {
                    format!(
                        "{}..{}",
                        value_string(bound, "lower"),
                        value_string(bound, "upper")
                    )
                })
                .collect::<Vec<_>>()
                .join(" × ")
        })
        .unwrap_or_else(|| "--".to_owned());
    format!("{dimensions}D `{bounds}`")
}

fn controls_report(capture: &Capture) -> String {
    let mut output = report_header(
        "Control-client comparison",
        "Control projections are recorded separately and never replace raw COM evidence.",
    );
    output.push_str(
        "| Case | Raw result | Control result | Classification |\n| --- | --- | --- | --- |\n",
    );
    let mut rows = capture
        .observations
        .iter()
        .filter(|observation| observation.get("control").is_some())
        .collect::<Vec<_>>();
    rows.sort_by_key(|row| value_string(row, "id"));
    if rows.is_empty() {
        output.push_str("| -- | No control result captured. | -- | Inconclusive |\n");
    }
    for row in rows {
        output.push_str(&format!(
            "| `{}` | `{}` | `{}` | {} |\n",
            value_string(row, "case_id"),
            value_string(row.get("result").unwrap_or(&Value::Null), "variant_type"),
            compact_json(row.get("control")),
            value_string(row, "classification")
        ));
    }
    output
}

fn workbooks_add_diagnostic_report(capture: &Capture) -> String {
    let mut output = report_header(
        "Workbooks.Add diagnostic",
        "This report compares only the explicit Prompt 05B dispatch frames; it does not turn a control projection into raw ABI evidence.",
    );
    output.push_str("| Case | Target | Audited/runtime DISPID | Flags | LCID | Args/named | HRESULT | Result VARTYPE | Workbook/cleanup |\n| --- | --- | --- | --- | --- | --- | --- | --- | --- |\n");
    let mut rows = capture
        .observations
        .iter()
        .filter(|record| {
            value_string(record, "case_id").starts_with("workbooks.")
                || value_string(record, "id") == "runtime.control.pywin32-dispatchex-workbooks-add"
        })
        .collect::<Vec<_>>();
    rows.sort_by_key(|record| value_string(record, "id"));
    if rows.is_empty() {
        output.push_str("| -- | -- | -- | -- | -- | -- | No diagnostic captured. | -- | -- |\n");
    }
    for row in rows {
        if value_string(row, "classification") == "Control-confirmed" {
            output.push_str(&format!(
                "| `pywin32-control` | `{}` | -- | -- | -- | -- | `{}` | Projection only | `{}` |\n",
                markdown_cell(&value_string(row, "activation")),
                markdown_cell(&value_string(row, "workbooks_add")),
                markdown_cell(&value_string(row, "created_workbook")),
            ));
            continue;
        }
        let frame = row.get("frame").unwrap_or(&Value::Null);
        let hresult = row
            .get("returned_hresult")
            .and_then(|value| value.get("hex"))
            .and_then(Value::as_str)
            .unwrap_or("--");
        output.push_str(&format!(
            "| `{}` | `{}` | {}/{} | `{}` | `{}` | {}/{} | `{}` | `{}` | {} |\n",
            markdown_cell(&value_string(row, "case_id")),
            markdown_cell(&value_string(row, "canonical_owner")),
            value_string(row, "audited_dispid"),
            value_string(row, "runtime_resolved_dispid"),
            markdown_cell(&value_string(row, "invoke_flags")),
            value_string(row, "lcid"),
            value_string(frame, "c_args"),
            value_string(frame, "c_named_args"),
            hresult,
            markdown_cell(&value_string(row, "result_vartype_after_call")),
            markdown_cell(&compact_json(row.get("cleanup_result"))),
        ));
    }
    output
}

fn invocation_frames_report(capture: &Capture) -> String {
    let mut output = report_header(
        "Raw invocation frames",
        "All values below are copied diagnostics: nullness is recorded, but raw pointer addresses are never persisted.",
    );
    output.push_str("| Case | Member | GetIDsOfNames DISPID | Frame | EXCEPINFO | puArgErr | Result ownership |\n| --- | --- | --- | --- | --- | --- | --- |\n");
    let mut rows = capture
        .observations
        .iter()
        .filter(|record| {
            record
                .get("categories")
                .and_then(Value::as_array)
                .is_some_and(|categories| {
                    categories
                        .iter()
                        .any(|category| category.as_str() == Some("invocation"))
                })
        })
        .collect::<Vec<_>>();
    rows.sort_by_key(|record| value_string(record, "id"));
    if rows.is_empty() {
        output.push_str("| -- | -- | -- | No invocation diagnostic captured. | -- | -- | -- |\n");
    }
    for row in rows {
        output.push_str(&format!(
            "| `{}` | `{}` | `{}` | {} | {} | {} | {} |\n",
            markdown_cell(&value_string(row, "case_id")),
            markdown_cell(&value_string(row, "member_name")),
            value_string(row, "runtime_resolved_dispid"),
            markdown_cell(&compact_json(row.get("frame"))),
            markdown_cell(&compact_json(row.get("excepinfo"))),
            markdown_cell(&compact_json(row.get("pu_arg_err"))),
            markdown_cell(&value_string(row, "result_ownership_state")),
        ));
    }
    output
}

fn control_comparison_report(capture: &Capture) -> String {
    let mut output = report_header(
        "Control comparison",
        "Independent controls establish only their own projection-level result; raw VARIANT and SAFEARRAY claims remain limited to the owned raw probe.",
    );
    output.push_str("| Control | Activation | Result | Boundary |\n| --- | --- | --- | --- |\n");
    let mut controls = capture
        .observations
        .iter()
        .filter(|record| value_string(record, "classification") == "Control-confirmed")
        .collect::<Vec<_>>();
    controls.sort_by_key(|record| value_string(record, "id"));
    if controls.is_empty() {
        output.push_str("| -- | -- | No independent control captured. | -- |\n");
    }
    for control in controls {
        output.push_str(&format!(
            "| `{}` | `{}` | Workbooks.Add: `{}`; workbook: `{}` | No raw VARIANT or SAFEARRAY inference. |\n",
            markdown_cell(&value_string(control, "client")),
            markdown_cell(&value_string(control, "activation")),
            markdown_cell(&value_string(control, "workbooks_add")),
            markdown_cell(&value_string(control, "created_workbook")),
        ));
    }
    output
}

fn unresolved_report(capture: &Capture) -> String {
    let mut output = report_header(
        "Unresolved runtime questions",
        "Not-tested and inconclusive entries remain explicit rather than inferred from documentation, a typelib, or a control projection.",
    );
    output.push_str("| Target | Classification | Detail |\n| --- | --- | --- |\n");
    if capture.unresolved.is_empty() {
        output.push_str("| -- | -- | No unresolved entry was emitted. |\n");
    } else {
        for record in &capture.unresolved {
            output.push_str(&format!(
                "| `{}` | {} | {} |\n",
                markdown_cell(&value_string(record, "target")),
                markdown_cell(&value_string(record, "classification")),
                markdown_cell(&value_string(record, "detail"))
            ));
        }
    }
    output
}

fn parity_environment_report(capture: &Capture) -> String {
    let mut output = report_header(
        "Environment stability matrix",
        "Only normalized Prompt 05D owned-session state is shown; HWNDs and paths are never persisted.",
    );
    output.push_str("| Mode | Operation | Excel / Workbooks before | Visible / user control / interactive / ready | Calculation / security / alerts | Paths | Owned process | Classification |\n| --- | --- | --- | --- | --- | --- | --- | --- |\n");
    let mut rows = capture
        .environments
        .iter()
        .filter(|record| record.get("parity_mode").is_some())
        .collect::<Vec<_>>();
    rows.sort_by_key(|record| value_string(record, "id"));
    if rows.is_empty() {
        output.push_str(
            "| -- | -- | -- | -- | -- | No Prompt 05D session recorded. | Not tested |\n",
        );
    }
    for row in rows {
        let state = row.get("session_state").unwrap_or(&Value::Null);
        let process = state.get("owned_process").unwrap_or(&Value::Null);
        let paths = format!(
            "startup {}; default {}",
            session_property_value(state, "startup_path"),
            session_property_value(state, "default_file_path"),
        );
        output.push_str(&format!(
            "| `{}` | `{}` | `{}` / `{}` | `{}` / `{}` / `{}` / `{}` | `{}` / `{}` / `{}` | `{}` | `{}` | `{}` |\n",
            value_string(row, "parity_mode"),
            value_string(row, "parity_operation"),
            session_property_value(state, "application_version"),
            session_property_value(state, "workbooks_count"),
            session_property_value(state, "visible"),
            session_property_value(state, "user_control"),
            session_property_value(state, "interactive"),
            session_property_value(state, "ready"),
            session_property_value(state, "calculation"),
            session_property_value(state, "automation_security"),
            session_property_value(state, "display_alerts"),
            paths,
            if value_string(process, "status") == "available" { "recorded" } else { "Not tested" },
            value_string(row, "classification"),
        ));
    }
    output
}

fn session_property_value(state: &Value, name: &str) -> String {
    let Some(field) = state.get(name) else {
        return "Not tested".to_owned();
    };
    if field.get("value_recorded") == Some(&Value::Bool(false)) {
        return "redacted".to_owned();
    }
    let status = value_string(field, "status");
    if status != "available" {
        return status;
    }
    let Some(value) = field.get("value") else {
        return "available".to_owned();
    };
    match value.get("kind").and_then(Value::as_str) {
        Some("bool") => match value.get("raw").and_then(Value::as_i64) {
            Some(0) => "false".to_owned(),
            Some(_) => "true".to_owned(),
            None => "available".to_owned(),
        },
        Some("bstr") | Some("i4") => value_string(value, "value"),
        Some("error") => format!("error {}", value_string(value, "signed_i32")),
        _ => markdown_cell(&compact_json(Some(value))),
    }
}

fn parity_operation_report(capture: &Capture, title: &str, operation: &str) -> String {
    let mut output = report_header(
        title,
        "Each row is a separately owned, freshly activated Excel session under one explicit Rust configuration; no retry loop is used.",
    );
    output.push_str("| Mode | Activation / CLSCTX | LCID | Invoke / frame | HRESULT | Workbook | Cleanup | Classification |\n| --- | --- | ---: | --- | --- | --- | --- | --- |\n");
    let mut rows = capture
        .observations
        .iter()
        .filter(|record| value_string(record, "parity_operation") == operation)
        .collect::<Vec<_>>();
    rows.sort_by_key(|record| value_string(record, "id"));
    if rows.is_empty() {
        output.push_str(
            "| -- | -- | -- | No Prompt 05D observation captured. | -- | -- | -- | Not tested |\n",
        );
    }
    for row in rows {
        let config = row.get("configuration").unwrap_or(&Value::Null);
        let frame = row.get("frame").unwrap_or(&Value::Null);
        output.push_str(&format!(
            "| `{}` | `{}` / `{}` | `{}` | `{}` / `{}` args | `{}` | `{}` | `{}` | `{}` |\n",
            value_string(row, "parity_mode"),
            markdown_cell(&value_string(config, "activation_api")),
            markdown_cell(&value_string(config, "clsctx")),
            value_string(config, "invoke_lcid"),
            markdown_cell(&value_string(config, "invoke_strategy")),
            value_string(frame, "c_args"),
            value_string(row.get("returned_hresult").unwrap_or(&Value::Null), "hex"),
            value_string(row, "workbook_created"),
            markdown_cell(&compact_json(row.get("cleanup"))),
            value_string(row, "classification"),
        ));
    }
    output
}

fn parity_smoke_report(capture: &Capture) -> String {
    let mut output = report_header(
        "Range smoke test",
        "The smoke test is entered only after the same Rust configuration creates or opens a workbook.",
    );
    output.push_str("| Mode | Workbook access | `A1.Value2 = 42` | Read VARTYPE / value | ClearContents | Classification |\n| --- | --- | --- | --- | --- | --- |\n");
    let mut rows = capture
        .observations
        .iter()
        .filter(|record| value_string(record, "parity_operation") == "range-smoke")
        .collect::<Vec<_>>();
    rows.sort_by_key(|record| value_string(record, "id"));
    if rows.is_empty() {
        output.push_str("| -- | -- | Not entered. | -- | -- | Not tested |\n");
    }
    for row in rows {
        output.push_str(&format!(
            "| `{}` | `{}` | `{}` | `{}` / `{}` | `{}` | `{}` |\n",
            value_string(row, "parity_mode"),
            value_string(row, "workbook_access"),
            value_string(row, "write_hresult"),
            value_string(row, "read_vartype"),
            markdown_cell(&compact_json(row.get("read_value"))),
            value_string(row, "clear_hresult"),
            value_string(row, "classification"),
        ));
    }
    output
}

fn parity_cleanup_report(capture: &Capture) -> String {
    let mut output = report_header(
        "Parity cleanup",
        "Only owned Excel instances are closed; no process is forcibly terminated by this probe.",
    );
    output.push_str("| Mode | Workbook close | Excel Quit | Owned process exit | Forced termination | Classification |\n| --- | --- | --- | --- | --- | --- |\n");
    let mut rows = capture
        .observations
        .iter()
        .filter(|record| value_string(record, "parity_operation") == "cleanup")
        .collect::<Vec<_>>();
    rows.sort_by_key(|record| value_string(record, "id"));
    if rows.is_empty() {
        output.push_str("| -- | -- | -- | No Prompt 05D cleanup recorded. | -- | Not tested |\n");
    }
    for row in rows {
        let cleanup = row.get("cleanup").unwrap_or(&Value::Null);
        output.push_str(&format!(
            "| `{}` | `{}` | `{}` | `{}` | `{}` | `{}` |\n",
            value_string(row, "parity_mode"),
            value_string(cleanup, "workbook_closed"),
            value_string(cleanup, "excel_quit_requested"),
            value_string(cleanup, "process_exited"),
            value_string(cleanup, "forced_termination"),
            value_string(row, "classification"),
        ));
    }
    output
}

fn parity_blockers_report(capture: &Capture) -> String {
    let mut output = report_header(
        "Remaining Prompt 05D blockers",
        "Classifications distinguish the client implementation from session and host state rather than assigning intermittent failures to Rust by default.",
    );
    output.push_str("| Mode | Target | Difference classification | Detail |\n| --- | --- | --- | --- |\n");
    let rows = capture
        .unresolved
        .iter()
        .filter(|record| value_string(record, "id").contains("05d"))
        .collect::<Vec<_>>();
    if rows.is_empty() {
        output.push_str("| -- | -- | -- | No remaining Prompt 05D blocker was emitted. |\n");
    } else {
        for row in rows {
            output.push_str(&format!(
                "| `{}` | `{}` | `{}` | {} |\n",
                markdown_cell(&parity_mode_for_unresolved(row)),
                markdown_cell(&value_string(row, "target")),
                markdown_cell(&value_string(row, "difference_classification")),
                markdown_cell(&value_string(row, "detail")),
            ));
        }
    }
    output
}

fn parity_mode_for_unresolved(record: &Value) -> String {
    let explicit = value_string(record, "parity_mode");
    if explicit != "--" {
        return explicit;
    }
    let id = value_string(record, "id");
    ParityMode::ALL
        .iter()
        .map(|mode| mode.id())
        .find(|mode| id.contains(mode))
        .unwrap_or("--")
        .to_owned()
}

fn markdown_cell(value: &str) -> String {
    value
        .trim()
        .replace("\r\n", "<br>")
        .replace('\n', "<br>")
        .replace('|', "\\|")
}

fn value_string(value: &Value, field: &str) -> String {
    value.get(field).map_or_else(
        || "--".to_owned(),
        |field| match field {
            Value::String(value) => value.clone(),
            Value::Null => "--".to_owned(),
            other => other.to_string(),
        },
    )
}

fn compact_json(value: Option<&Value>) -> String {
    value
        .map_or_else(|| "--".to_owned(), Value::to_string)
        .replace('|', "\\|")
}

fn runtime_manifest(control_version: &str) -> String {
    format!(
        "schema_version = 1\nprobe_version = {PROBE_VERSION}\nprompt_05_start_origin_master_commit = \"2ac52effadafe6cd5b95b448f356a62389fa54f2\"\nprompt_05b_start_origin_master_commit = \"dbbc9600af2628b45e3f05431ce168102ad9e6ae\"\ndocumentation_source_pin = \"b2cda886ea91e36c62eb1cb177133ad024ecd345\"\ntypelib_guid = \"{{00020813-0000-0000-C000-000000000046}}\"\ntypelib_version = \"1.9\"\nexcel_file_version = \"16.0.20131.20154\"\noffice_bitness = \"64-bit\"\nwindows_version = \"Windows 10 Enterprise 25H2 build 26200.8875\"\nlocale = \"LOCALE_USER_DEFAULT (0x0400) for Prompt 05B raw calls\"\nlist_separator = \"not-recorded-by-raw-probe\"\ndecimal_separator = \"not-recorded-by-raw-probe\"\ndate_system = \"not-recorded-by-raw-probe\"\ncalculation_mode_before = \"not-recorded-by-raw-probe\"\ncalculation_mode_after = \"not-recorded-by-raw-probe\"\nexecution_date = \"2026-07-21\"\ncontrol_version = \"{}\"\nprocess_isolation = \"CoCreateInstance local server; require Hwnd-to-PID/start-time identity before Range work; close only created workbooks; record bounded cleanup exactly\"\n",
        escape_toml(control_version)
    )
}

fn escape_toml(value: &str) -> String {
    value.replace('\\', "\\\\").replace('"', "\\\"")
}

#[cfg(windows)]
mod windows_live {
    use super::*;
    use std::mem::ManuallyDrop;
    use std::process::Command;
    use std::thread;
    use std::time::{Duration, Instant};
    use windows::Win32::Foundation::{CloseHandle, FILETIME, HWND, WAIT_OBJECT_0};
    use windows::Win32::Globalization::LOCALE_USER_DEFAULT;
    use windows::Win32::System::Com::{
        CLSCTX_LOCAL_SERVER, CLSCTX_SERVER, COINIT_APARTMENTTHREADED, CY, CoCreateInstance,
        CoCreateInstanceEx, CoInitializeEx, CoUninitialize, DISPATCH_FLAGS, DISPATCH_METHOD,
        DISPATCH_PROPERTYGET, DISPATCH_PROPERTYPUT, DISPATCH_PROPERTYPUTREF, DISPPARAMS,
        EXCEPINFO, IDispatch, MULTI_QI, SAFEARRAYBOUND,
    };
    use windows::Win32::System::Ole::{
        DISPID_PROPERTYPUT, SafeArrayCreate, SafeArrayGetDim, SafeArrayGetElement,
        SafeArrayGetLBound, SafeArrayGetUBound, SafeArrayGetVartype, SafeArrayPutElement,
    };
    use windows::Win32::System::Threading::{
        GetCurrentThreadId, GetProcessTimes, OpenProcess, PROCESS_QUERY_LIMITED_INFORMATION,
        PROCESS_SYNCHRONIZE, WaitForSingleObject,
    };
    use windows::Win32::System::Variant::{
        VARENUM, VARIANT, VT_ARRAY, VT_BOOL, VT_BSTR, VT_CY, VT_DATE, VT_DISPATCH, VT_EMPTY,
        VT_ERROR, VT_I4, VT_NULL, VT_R8, VT_VARIANT, VariantClear,
    };
    use windows::Win32::UI::WindowsAndMessaging::{
        DispatchMessageW, GetWindowThreadProcessId, MSG, PM_REMOVE, PeekMessageW, TranslateMessage,
    };
    use windows::core::{BSTR, GUID, HSTRING, IUnknown, Interface, PCWSTR};

    const DISPID_APPLICATION_WORKBOOKS: i32 = 572;
    const DISPID_WORKBOOKS_ADD: i32 = 181;
    const DISPID_WORKBOOKS_OPEN: i32 = 1923;
    const DISPID_APPLICATION_ACTIVE_SHEET: i32 = 307;
    const DISPID_WORKBOOK_CLOSE: i32 = 277;
    const DISPID_APPLICATION_QUIT: i32 = 302;
    const DISPID_WORKSHEET_RANGE: i32 = 197;
    const DISPID_RANGE_VALUE: i32 = 6;
    const DISPID_RANGE_VALUE2: i32 = 1388;
    const DISPID_RANGE_FORMULA: i32 = 261;
    const DISPID_RANGE_FORMULA2: i32 = 1580;
    const DISPID_RANGE_TEXT: i32 = 138;
    const DISPID_RANGE_HAS_FORMULA: i32 = 1382;
    const DISPID_RANGE_CLEAR: i32 = 111;
    const DISPID_RANGE_CLEAR_CONTENTS: i32 = 3413;
    const DISP_E_PARAMNOTFOUND: i32 = -2_147_352_572;
    const INVOKE_LCID: u32 = LOCALE_USER_DEFAULT;
    const INVOKE_LCID_POLICY: &str = "LOCALE_USER_DEFAULT (0x0400)";

    pub(super) fn capture(root: &Path, control_script: Option<&Path>) -> Result<Capture, String> {
        let apartment = Apartment::initialize()?;
        let apartment_record = apartment.record();
        let mut session = match Session::create(root) {
            Ok(session) => session,
            Err(failure) => {
                drop(apartment);
                return Ok(blocked_capture(failure, apartment_record));
            }
        };
        let pywin32_control = pywin32_control_record();
        let mut observations = vec![pywin32_control.clone()];
        observations.extend(session.take_invocation_observations());
        let mut cases = Vec::new();
        let mut unresolved = standard_unresolved();
        let add_diagnostic_result =
            session.workbooks_add_optional_diagnostics(&mut observations, &mut cases);
        observations.extend(session.take_invocation_observations());
        let result = add_diagnostic_result.and_then(|()| {
            if session.workbook.is_some() {
                session.run_cases(&mut observations, &mut cases, None)
            } else {
                let error = session.primary_add_failure.clone().unwrap_or_else(|| {
                    "No Workbooks.Add diagnostic representation returned an owned workbook dispatch object"
                        .to_owned()
                });
                cases.push(case(
                    "range.runtime-matrix",
                    "inconclusive",
                    "No Workbooks.Add representation returned a retained owned workbook, so the Range smoke test and full Prompt 05 matrix were not entered.",
                ));
                Err(error)
            }
        });
        let control = if result.is_ok() {
            run_control(control_script)
        } else {
            None
        };
        let cleanup = session.cleanup();
        observations.push(json!({
            "schema_version": 1,
            "id": "runtime.cleanup.05b-owned-excel",
            "case_id": "cleanup.owned-excel",
            "classification": "Runtime-observed",
            "categories": ["cleanup", "workbooks-add-diagnostic"],
            "environment_id": "excel-16.0.20131.20154-win64-05b",
            "cleanup": cleanup.clone(),
            "raw_pointer_values_recorded": false,
        }));
        drop(apartment);
        if let Err(error) = result {
            unresolved.push(json!({
                "schema_version": 1,
                "id": "runtime.unresolved.live-probe-failure",
                "target": "Owned live Excel probe",
                "classification": "Inconclusive",
                "detail": error,
            }));
        }
        if !cleanup
            .get("process_exited")
            .and_then(Value::as_bool)
            .unwrap_or(false)
        {
            unresolved.push(json!({
                "schema_version": 1,
                "id": "runtime.unresolved.cleanup",
                "target": "Owned Excel process exit",
                "classification": "Inconclusive",
                "detail": "Quit was requested but bounded exact-PID exit verification did not succeed; no process was forcibly terminated.",
            }));
        }
        let environment = json!({
            "schema_version": 1,
            "id": "excel-16.0.20131.20154-win64-05b",
            "classification": "Version-specific",
            "excel_file_version": "16.0.20131.20154",
            "office_bitness": "64-bit",
            "typelib_guid": "{00020813-0000-0000-C000-000000000046}",
            "typelib_version": "1.9",
            "windows_version": "Windows 10 Enterprise 25H2 build 26200.8875",
            "pre_workbook_application_configuration": "none; Visible and DisplayAlerts were not changed before Workbooks.Add",
            "owned_process": session.identity,
            "apartment": apartment_record,
            "cleanup": cleanup,
            "control": control,
            "pywin32_control": pywin32_control,
            "source": "isolated-raw-com-probe"
        });
        let control_version = control
            .as_ref()
            .and_then(|control| control.get("version"))
            .and_then(Value::as_str)
            .unwrap_or("not-run");
        Ok(Capture {
            manifest: runtime_manifest(control_version),
            environments: vec![environment],
            observations,
            cases,
            unresolved,
        })
    }

    pub(super) fn capture_parity(
        root: &Path,
        mode: ParityMode,
        fixture: Option<&Path>,
        run_id: &str,
    ) -> Result<Capture, String> {
        let apartment = Apartment::initialize()?;
        let apartment_record = apartment.record();
        let configuration = parity_configuration(mode);
        let mut environments = Vec::new();
        let mut observations = Vec::new();
        let mut cases = Vec::new();
        let mut unresolved = Vec::new();

        let add = run_parity_operation(
            root,
            &configuration,
            run_id,
            "workbooks-add",
            None,
            &apartment_record,
        );
        let add_succeeded = add.succeeded;
        environments.push(add.environment);
        observations.extend(add.observations);
        cases.extend(add.cases);
        unresolved.extend(add.unresolved);

        let open = run_parity_operation(
            root,
            &configuration,
            run_id,
            "workbooks-open",
            fixture,
            &apartment_record,
        );
        environments.push(open.environment);
        observations.extend(open.observations);
        cases.extend(open.cases);
        unresolved.extend(open.unresolved);

        if matches!(mode, ParityMode::ComtypesGenerated) {
            unresolved.push(json!({
                "schema_version": 1,
                "id": format!("runtime.unresolved.05d.{run_id}.{}.generated-vtable", mode.id()),
                "target": "comtypes generated dual-interface vtable parity",
                "difference_classification": "inconclusive",
                "classification": "Not tested",
                "detail": "The installed typelib confirms dual candidates, but no generated Rust Excel bindings are available. The probe records a raw IDispatch fallback and does not hand-write a vtable layout.",
            }));
        }
        if !add_succeeded {
            unresolved.push(json!({
                "schema_version": 1,
                "id": format!("runtime.unresolved.05d.{run_id}.{}.range-smoke", mode.id()),
                "target": "Prompt 05 Range smoke through Workbooks.Add",
                "difference_classification": "inconclusive",
                "classification": "Not tested",
                "detail": "The bounded Workbooks.Add operation did not return an owned workbook dispatch object, so this mode did not enter the A1.Value2 smoke test.",
            }));
        }
        drop(apartment);
        Ok(Capture {
            manifest: runtime_manifest(&format!("05d-parity-{}", mode.id())),
            environments,
            observations,
            cases,
            unresolved,
        })
    }

    struct ParityOperationCapture {
        environment: Value,
        observations: Vec<Value>,
        cases: Vec<Value>,
        unresolved: Vec<Value>,
        succeeded: bool,
    }

    fn run_parity_operation(
        root: &Path,
        configuration: &ParityConfiguration,
        run_id: &str,
        operation: &str,
        fixture: Option<&Path>,
        apartment: &Value,
    ) -> ParityOperationCapture {
        let environment_id = format!(
            "runtime.environment.05d.{run_id}.{}.{}",
            configuration.mode.id(),
            operation
        );
        let mut observations = Vec::new();
        let mut cases = Vec::new();
        let mut unresolved = Vec::new();
        let mut session = match ParitySession::activate(root, configuration) {
            Ok(session) => session,
            Err(detail) => {
                let environment = json!({
                    "schema_version": 1,
                    "id": environment_id,
                    "parity_mode": configuration.mode.id(),
                    "parity_operation": operation,
                    "classification": "Inconclusive",
                    "configuration": configuration,
                    "apartment": apartment,
                    "session_state": {"status":"Not tested"},
                    "activation_error": detail,
                    "raw_hwnd_recorded": false,
                    "raw_pointer_values_recorded": false,
                    "raw_paths_recorded": false,
                });
                cases.push(parity_case(
                    configuration.mode,
                    operation,
                    "inconclusive",
                    "Activation did not produce an owned Excel IDispatch instance.",
                ));
                unresolved.push(parity_unresolved(
                    run_id,
                    configuration.mode,
                    operation,
                    "activation difference",
                    "inconclusive",
                    "Activation failed before a bounded Workbook operation could be attempted.",
                ));
                return ParityOperationCapture {
                    environment,
                    observations,
                    cases,
                    unresolved,
                    succeeded: false,
                };
            }
        };
        let session_state = session.session_state(configuration);
        let mut operation_record = if operation == "workbooks-open" && fixture.is_none() {
            json!({
                "schema_version": 1,
                "id": format!("runtime.05d.{run_id}.{}.{}", configuration.mode.id(), operation),
                "parity_mode": configuration.mode.id(),
                "parity_operation": operation,
                "classification": "Not tested",
                "configuration": configuration,
                "fixture": "not available; temporary fixture paths are never committed",
                "workbook_created": false,
                "returned_hresult": Value::Null,
                "frame": Value::Null,
                "raw_pointer_values_recorded": false,
            })
        } else {
            session.workbook_operation(configuration, run_id, operation, fixture)
        };
        let succeeded = operation_record
            .get("workbook_created")
            .and_then(Value::as_bool)
            .unwrap_or(false);
        if succeeded
            && (operation == "workbooks-add"
                || !operation_record
                    .get("range_smoke_entered")
                    .is_some_and(|value| value.as_bool().unwrap_or(false)))
        {
            let smoke = session.range_smoke(configuration, run_id, operation);
            operation_record["range_smoke_entered"] = Value::Bool(true);
            observations.push(smoke);
        }
        let cleanup = session.cleanup(configuration);
        operation_record["cleanup"] = cleanup.clone();
        observations.push(operation_record);
        observations.push(json!({
            "schema_version": 1,
            "id": format!("runtime.05d.{run_id}.{}.{}.cleanup", configuration.mode.id(), operation),
            "parity_mode": configuration.mode.id(),
            "parity_operation": "cleanup",
            "classification": "Runtime-observed",
            "cleanup": cleanup,
            "raw_pointer_values_recorded": false,
        }));
        let status = if succeeded {
            "completed"
        } else {
            "inconclusive"
        };
        cases.push(parity_case(
            configuration.mode,
            operation,
            status,
            if succeeded {
                "The bounded operation returned an owned workbook dispatch object."
            } else {
                "The bounded operation did not return an owned workbook dispatch object."
            },
        ));
        let environment = json!({
            "schema_version": 1,
            "id": environment_id,
            "parity_mode": configuration.mode.id(),
            "parity_operation": operation,
            "classification": if succeeded { "Runtime-observed" } else { "Inconclusive" },
            "configuration": configuration,
            "apartment": apartment,
            "session_state": session_state,
            "office_bitness": "64-bit",
            "excel_file_version": "16.0.20131.20154",
            "windows_version": "Windows 10 Enterprise 25H2 build 26200.8875",
            "raw_hwnd_recorded": false,
            "raw_pointer_values_recorded": false,
            "raw_paths_recorded": false,
        });
        ParityOperationCapture {
            environment,
            observations,
            cases,
            unresolved,
            succeeded,
        }
    }

    fn parity_case(mode: ParityMode, operation: &str, status: &str, detail: &str) -> Value {
        json!({
            "schema_version": 1,
            "id": format!("runtime.case.05d.{}.{}", mode.id(), operation),
            "parity_mode": mode.id(),
            "status": status,
            "detail": detail,
            "classification": if status == "completed" { "Runtime-observed" } else { "Inconclusive" },
        })
    }

    fn parity_unresolved(
        run_id: &str,
        mode: ParityMode,
        operation: &str,
        difference_classification: &str,
        classification: &str,
        detail: &str,
    ) -> Value {
        json!({
            "schema_version": 1,
            "id": format!("runtime.unresolved.05d.{run_id}.{}.{}", mode.id(), operation),
            "target": operation,
            "difference_classification": difference_classification,
            "classification": classification,
            "detail": detail,
        })
    }

    struct ParitySession {
        app: Option<IDispatch>,
        workbooks: Option<IDispatch>,
        workbook: Option<IDispatch>,
        process_handle: windows::Win32::Foundation::HANDLE,
        identity: Value,
    }

    impl ParitySession {
        fn activate(root: &Path, configuration: &ParityConfiguration) -> Result<Self, String> {
            let class_id = excel_application_clsid(root)?;
            let app = create_parity_dispatch(class_id, configuration)?;
            let mut session = Self {
                app: Some(app),
                workbooks: None,
                workbook: None,
                process_handle: windows::Win32::Foundation::HANDLE::default(),
                identity: Value::Null,
            };
            session.identity = session.verify_owned_process(configuration)?;
            Ok(session)
        }

        fn session_state(&mut self, configuration: &ParityConfiguration) -> Value {
            let app = self
                .app
                .as_ref()
                .expect("active parity Application")
                .clone();
            let workbooks = self.ensure_workbooks(configuration).ok();
            json!({
                "application_version": parity_property(&app, configuration, "Version", false),
                "hwnd": {"status":"available", "value_recorded":false},
                "owned_process": self.identity,
                "visible": parity_property(&app, configuration, "Visible", false),
                "user_control": parity_property(&app, configuration, "UserControl", false),
                "interactive": parity_property(&app, configuration, "Interactive", false),
                "ready": parity_property(&app, configuration, "Ready", false),
                "workbooks_count": workbooks.as_ref().map_or_else(|| json!({"status":"Not tested"}), |workbooks| parity_property(workbooks, configuration, "Count", false)),
                "startup_path": parity_property(&app, configuration, "StartupPath", true),
                "default_file_path": parity_property(&app, configuration, "DefaultFilePath", true),
                "calculation": parity_property(&app, configuration, "Calculation", false),
                "automation_security": parity_property(&app, configuration, "AutomationSecurity", false),
                "display_alerts": parity_property(&app, configuration, "DisplayAlerts", false),
                "modal_or_error_state": "not directly detectable through the bounded Automation surface",
                "application_type_info": dispatch_type_record_with_lcid(&app, configuration.invoke_lcid),
            })
        }

        fn ensure_workbooks(
            &mut self,
            configuration: &ParityConfiguration,
        ) -> Result<IDispatch, Value> {
            if let Some(workbooks) = &self.workbooks {
                return Ok(workbooks.clone());
            }
            let app = self
                .app
                .as_ref()
                .expect("active parity Application")
                .clone();
            let invocation = parity_member(
                &app,
                configuration,
                "application-workbooks",
                "Excel.Application",
                "Workbooks",
                DISPID_APPLICATION_WORKBOOKS,
                DISPATCH_PROPERTYGET,
                InvocationFrame::positional(Vec::new()),
            )?;
            let workbooks = dispatch_from_variant(&invocation.result.value).ok_or_else(|| {
                let mut record = invocation.diagnostic;
                record["detail"] =
                    Value::String("Application.Workbooks did not return VT_DISPATCH".to_owned());
                record
            })?;
            self.workbooks = Some(workbooks.clone());
            Ok(workbooks)
        }

        fn workbook_operation(
            &mut self,
            configuration: &ParityConfiguration,
            run_id: &str,
            operation: &str,
            fixture: Option<&Path>,
        ) -> Value {
            let mut record = match self.ensure_workbooks(configuration) {
                Ok(workbooks) => {
                    let (name, dispid, arguments) = if operation == "workbooks-open" {
                        (
                            "Open",
                            DISPID_WORKBOOKS_OPEN,
                            vec![VariantOwner::from_value(VARIANT::from(
                                fixture.expect("checked fixture").to_string_lossy().as_ref(),
                            ))],
                        )
                    } else {
                        ("Add", DISPID_WORKBOOKS_ADD, Vec::new())
                    };
                    match parity_member(
                        &workbooks,
                        configuration,
                        operation,
                        "Excel.Workbooks",
                        name,
                        dispid,
                        DISPATCH_METHOD,
                        InvocationFrame::positional(arguments),
                    ) {
                        Ok(invocation) => {
                            let workbook = dispatch_from_variant(&invocation.result.value);
                            let mut record = invocation.diagnostic;
                            record["workbook_created"] = Value::Bool(workbook.is_some());
                            if let Some(workbook) = workbook {
                                record["returned_dispatch"] = dispatch_type_record_with_lcid(
                                    &workbook,
                                    configuration.invoke_lcid,
                                );
                                self.workbook = Some(workbook);
                            }
                            record
                        }
                        Err(record) => record,
                    }
                }
                Err(record) => record,
            };
            record["schema_version"] = Value::from(1);
            record["id"] = Value::String(format!(
                "runtime.05d.{run_id}.{}.{}",
                configuration.mode.id(),
                operation
            ));
            record["parity_mode"] = Value::String(configuration.mode.id().to_owned());
            record["parity_operation"] = Value::String(operation.to_owned());
            record["classification"] = Value::String(
                if record
                    .get("workbook_created")
                    .and_then(Value::as_bool)
                    .unwrap_or(false)
                {
                    "Runtime-observed"
                } else {
                    "Inconclusive"
                }
                .to_owned(),
            );
            record["configuration"] =
                serde_json::to_value(configuration).expect("serializable parity configuration");
            record["fixture"] = Value::String(
                if operation == "workbooks-open" {
                    "temporary xlsx created by a known-good pywin32 311 control; path redacted"
                } else {
                    "not applicable"
                }
                .to_owned(),
            );
            if operation == "workbooks-open" {
                redact_fixture_path(&mut record);
            }
            record["raw_pointer_values_recorded"] = Value::Bool(false);
            record
        }

        fn range_smoke(
            &mut self,
            configuration: &ParityConfiguration,
            run_id: &str,
            source_operation: &str,
        ) -> Value {
            let result = (|| -> Result<Value, Value> {
                let app = self
                    .app
                    .as_ref()
                    .expect("active parity Application")
                    .clone();
                let sheet = parity_member(
                    &app,
                    configuration,
                    "range-smoke-active-sheet",
                    "Excel.Application",
                    "ActiveSheet",
                    DISPID_APPLICATION_ACTIVE_SHEET,
                    DISPATCH_PROPERTYGET,
                    InvocationFrame::positional(Vec::new()),
                )?;
                let sheet = dispatch_from_variant(&sheet.result.value)
                    .ok_or_else(|| json!({"detail":"ActiveSheet did not return VT_DISPATCH"}))?;
                let range = parity_member(
                    &sheet,
                    configuration,
                    "range-smoke-a1",
                    "Excel.Worksheet",
                    "Range",
                    DISPID_WORKSHEET_RANGE,
                    DISPATCH_PROPERTYGET,
                    InvocationFrame::positional(vec![VariantOwner::from_value(VARIANT::from(
                        "A1",
                    ))]),
                )?;
                let range = dispatch_from_variant(&range.result.value).ok_or_else(
                    || json!({"detail":"Worksheet.Range did not return VT_DISPATCH"}),
                )?;
                let write = parity_member(
                    &range,
                    configuration,
                    "range-smoke-value2-put",
                    "Excel.Range",
                    "Value2",
                    DISPID_RANGE_VALUE2,
                    DISPATCH_PROPERTYPUT,
                    InvocationFrame::property_put(VariantOwner::from_value(VARIANT::from(42_i32))),
                )?;
                let read = parity_member(
                    &range,
                    configuration,
                    "range-smoke-value2-get",
                    "Excel.Range",
                    "Value2",
                    DISPID_RANGE_VALUE2,
                    DISPATCH_PROPERTYGET,
                    InvocationFrame::positional(Vec::new()),
                )?;
                let clear = parity_member(
                    &range,
                    configuration,
                    "range-smoke-clear-contents",
                    "Excel.Range",
                    "ClearContents",
                    DISPID_RANGE_CLEAR_CONTENTS,
                    DISPATCH_METHOD,
                    InvocationFrame::positional(Vec::new()),
                )?;
                Ok(json!({
                    "write_hresult": write.diagnostic["returned_hresult"],
                    "write_frame": write.diagnostic["frame"],
                    "read_hresult": read.diagnostic["returned_hresult"],
                    "read_frame": read.diagnostic["frame"],
                    "read_vartype": vartype(&read.result.value),
                    "read_value": scalar_value(&read.result.value),
                    "clear_hresult": clear.diagnostic["returned_hresult"],
                    "clear_frame": clear.diagnostic["frame"],
                }))
            })();
            let mut record = match result {
                Ok(details) => {
                    json!({"classification":"Runtime-observed", "workbook_access":"succeeded", "details":details})
                }
                Err(details) => {
                    json!({"classification":"Inconclusive", "workbook_access":"succeeded", "detail":details})
                }
            };
            if let Some(details) = record.get("details").cloned() {
                for field in [
                    "write_hresult",
                    "write_frame",
                    "read_hresult",
                    "read_frame",
                    "read_vartype",
                    "read_value",
                    "clear_hresult",
                    "clear_frame",
                ] {
                    record[field] = details.get(field).cloned().unwrap_or(Value::Null);
                }
            }
            record["schema_version"] = Value::from(1);
            record["id"] = Value::String(format!(
                "runtime.05d.{run_id}.{}.range-smoke-{source_operation}",
                configuration.mode.id()
            ));
            record["parity_mode"] = Value::String(configuration.mode.id().to_owned());
            record["parity_operation"] = Value::String("range-smoke".to_owned());
            record["source_operation"] = Value::String(source_operation.to_owned());
            record["configuration"] =
                serde_json::to_value(configuration).expect("serializable parity configuration");
            record["raw_pointer_values_recorded"] = Value::Bool(false);
            record
        }

        fn verify_owned_process(
            &mut self,
            configuration: &ParityConfiguration,
        ) -> Result<Value, String> {
            let app = self
                .app
                .as_ref()
                .expect("active parity Application")
                .clone();
            let hwnd = parity_dynamic_member(
                &app,
                configuration,
                "Hwnd",
                DISPATCH_PROPERTYGET,
                InvocationFrame::positional(Vec::new()),
            )
            .map_err(|record| {
                format!(
                    "cannot obtain Hwnd for created Excel instance: {}",
                    compact_json(Some(&record))
                )
            })?;
            let hwnd = scalar_i32(&hwnd.result.value)
                .ok_or_else(|| "Excel Hwnd did not return VT_I4".to_owned())?;
            let mut process_id = 0;
            unsafe { GetWindowThreadProcessId(HWND(hwnd as *mut _), Some(&mut process_id)) };
            if process_id == 0 {
                return Err("Hwnd-to-PID ownership lookup returned zero".to_owned());
            }
            let handle = unsafe {
                OpenProcess(
                    PROCESS_QUERY_LIMITED_INFORMATION | PROCESS_SYNCHRONIZE,
                    false,
                    process_id,
                )
            }
            .map_err(|error| format!("cannot open created Excel PID {process_id}: {error}"))?;
            let mut creation = FILETIME::default();
            unsafe {
                GetProcessTimes(
                    handle,
                    &mut creation,
                    &mut FILETIME::default(),
                    &mut FILETIME::default(),
                    &mut FILETIME::default(),
                )
                .map_err(|error| format!("cannot record created Excel process time: {error}"))?;
            }
            self.process_handle = handle;
            let start_ticks =
                (u64::from(creation.dwHighDateTime) << 32) | u64::from(creation.dwLowDateTime);
            Ok(json!({
                "status": "available",
                "pid": process_id,
                "start_time_filetime_ticks": start_ticks,
                "window_verified": true,
                "raw_hwnd_recorded": false,
                "raw_path_recorded": false,
            }))
        }

        fn cleanup(&mut self, configuration: &ParityConfiguration) -> Value {
            let mut workbook_closed = false;
            if let Some(workbook) = self.workbook.take() {
                workbook_closed = parity_dynamic_member(
                    &workbook,
                    configuration,
                    "Close",
                    DISPATCH_METHOD,
                    InvocationFrame::positional(vec![VariantOwner::from_value(VARIANT::from(
                        false,
                    ))]),
                )
                .is_ok();
            }
            let excel_quit_requested = self.app.as_ref().is_some_and(|app| {
                parity_dynamic_member(
                    app,
                    configuration,
                    "Quit",
                    DISPATCH_METHOD,
                    InvocationFrame::positional(Vec::new()),
                )
                .is_ok()
            });
            self.workbooks.take();
            self.app.take();
            let process_exited = if !self.process_handle.is_invalid() {
                unsafe { WaitForSingleObject(self.process_handle, 15_000) == WAIT_OBJECT_0 }
            } else {
                false
            };
            if !self.process_handle.is_invalid() {
                let _ = unsafe { CloseHandle(self.process_handle) };
                self.process_handle = windows::Win32::Foundation::HANDLE::default();
            }
            cleanup_record(workbook_closed, excel_quit_requested, process_exited)
        }
    }

    impl Drop for ParitySession {
        fn drop(&mut self) {
            self.workbook.take();
            self.workbooks.take();
            self.app.take();
            if !self.process_handle.is_invalid() {
                let _ = unsafe { CloseHandle(self.process_handle) };
            }
        }
    }

    fn create_parity_dispatch(
        class_id: GUID,
        configuration: &ParityConfiguration,
    ) -> Result<IDispatch, String> {
        if matches!(configuration.mode, ParityMode::Pywin32Dynamic) {
            let iid = IDispatch::IID;
            let mut results = [MULTI_QI {
                pIID: &iid,
                pItf: std::mem::ManuallyDrop::new(None),
                hr: windows::core::HRESULT(0),
            }];
            unsafe {
                CoCreateInstanceEx(
                    &class_id,
                    None::<&IUnknown>,
                    CLSCTX_SERVER,
                    None,
                    &mut results,
                )
            }
            .map_err(|error| format!("CoCreateInstanceEx failed: {error}"))?;
            results[0]
                .hr
                .ok()
                .map_err(|error| format!("CoCreateInstanceEx interface request failed: {error}"))?;
            let unknown = unsafe { std::mem::ManuallyDrop::take(&mut results[0].pItf) }
                .ok_or_else(|| "CoCreateInstanceEx returned no requested interface".to_owned())?;
            unknown
                .cast::<IDispatch>()
                .map_err(|error| format!("CoCreateInstanceEx did not return IDispatch: {error}"))
        } else {
            let clsctx = if matches!(configuration.mode, ParityMode::RustBaseline) {
                CLSCTX_LOCAL_SERVER
            } else {
                CLSCTX_SERVER
            };
            unsafe { CoCreateInstance::<_, IDispatch>(&class_id, None, clsctx) }
                .map_err(|error| format!("CoCreateInstance failed: {error}"))
        }
    }

    fn parity_property(
        dispatch: &IDispatch,
        configuration: &ParityConfiguration,
        name: &str,
        redact: bool,
    ) -> Value {
        match parity_dynamic_member(
            dispatch,
            configuration,
            name,
            DISPATCH_PROPERTYGET,
            InvocationFrame::positional(Vec::new()),
        ) {
            Ok(_invocation) if redact => json!({"status":"available", "value_recorded":false}),
            Ok(invocation) => {
                json!({"status":"available", "value":scalar_value(&invocation.result.value)})
            }
            Err(record) => json!({
                "status": "Not tested",
                "hresult": record.get("returned_hresult").cloned().unwrap_or(Value::Null),
            }),
        }
    }

    fn parity_dynamic_member(
        dispatch: &IDispatch,
        configuration: &ParityConfiguration,
        name: &str,
        flags: DISPATCH_FLAGS,
        frame: InvocationFrame,
    ) -> Result<RawInvocation, Value> {
        let dispid = parity_name_dispid(dispatch, name, configuration.get_ids_of_names_lcid)
            .map_err(|error| json!({"member_name":name, "returned_hresult":hresult_json(error.code().0), "detail":"GetIDsOfNames failed"}))?;
        parity_member(
            dispatch,
            configuration,
            "state",
            "Excel.Application",
            name,
            dispid,
            flags,
            frame,
        )
    }

    #[allow(clippy::too_many_arguments)]
    fn parity_member(
        dispatch: &IDispatch,
        configuration: &ParityConfiguration,
        case_id: &str,
        canonical_owner: &str,
        name: &str,
        audited_dispid: i32,
        flags: DISPATCH_FLAGS,
        frame: InvocationFrame,
    ) -> Result<RawInvocation, Value> {
        let runtime_dispid =
            match parity_name_dispid(dispatch, name, configuration.get_ids_of_names_lcid) {
                Ok(value) => value,
                Err(error) => {
                    return Err(json!({
                        "case_id":case_id,
                        "canonical_owner":canonical_owner,
                        "member_name":name,
                        "audited_dispid":audited_dispid,
                        "runtime_resolved_dispid":Value::Null,
                        "get_ids_of_names_hresult":hresult_json(error.code().0),
                        "returned_hresult":Value::Null,
                        "frame":frame.diagnostic(),
                        "detail":"GetIDsOfNames failed",
                        "workbook_created":false,
                    }));
                }
            };
        if runtime_dispid != audited_dispid {
            return Err(json!({
                "case_id":case_id,
                "canonical_owner":canonical_owner,
                "member_name":name,
                "audited_dispid":audited_dispid,
                "runtime_resolved_dispid":runtime_dispid,
                "get_ids_of_names_hresult":hresult_json(0),
                "returned_hresult":Value::Null,
                "frame":frame.diagnostic(),
                "detail":"runtime DISPID did not match installed typelib evidence",
                "workbook_created":false,
            }));
        }
        match parity_invoke_with_diagnostic(
            dispatch,
            configuration,
            InvocationContext {
                case_id,
                canonical_owner,
                member_name: name,
                audited_dispid,
                runtime_dispid: Some(runtime_dispid),
                get_ids_hresult: Some(0),
            },
            flags,
            frame,
        ) {
            Ok(invocation) => Ok(invocation),
            Err(failure) => Err(failure.diagnostic),
        }
    }

    fn parity_invoke_with_diagnostic(
        dispatch: &IDispatch,
        configuration: &ParityConfiguration,
        context: InvocationContext<'_>,
        flags: DISPATCH_FLAGS,
        mut frame: InvocationFrame,
    ) -> Result<RawInvocation, InvocationFailure> {
        let frame_record = frame.diagnostic();
        let params = frame.params();
        let mut result = VariantOwner::empty();
        let mut exception = EXCEPINFO::default();
        let mut arg_error = u32::MAX;
        let invoke_result = unsafe {
            dispatch.Invoke(
                context.audited_dispid,
                &GUID::from_u128(0),
                configuration.invoke_lcid,
                flags,
                &params,
                Some(&mut result.value),
                Some(&mut exception),
                Some(&mut arg_error),
            )
        };
        let hresult = invoke_result
            .as_ref()
            .map(|_| 0_i32)
            .map_err(|error| error.code().0)
            .unwrap_or_else(|value| value);
        let exception_record = normalize_exception(&mut exception);
        let diagnostic = json!({
            "case_id": context.case_id,
            "canonical_owner": context.canonical_owner,
            "member_name": context.member_name,
            "audited_dispid": context.audited_dispid,
            "runtime_resolved_dispid": context.runtime_dispid,
            "dispid_matches_audit": dispid_matches(context.audited_dispid, context.runtime_dispid),
            "get_ids_of_names_hresult": context.get_ids_hresult.map(hresult_json),
            "lcid": configuration.invoke_lcid,
            "lcid_policy": if configuration.invoke_lcid == 0 { "source-matched LCID 0" } else { "preserved LOCALE_USER_DEFAULT (0x0400) baseline" },
            "invoke_flags": format_dispatch_flags(flags),
            "invoke_flags_raw": flags.0,
            "frame": frame_record,
            "result_variant_initialized_before_call": true,
            "returned_hresult": hresult_json(hresult),
            "excepinfo": exception_record,
            "pu_arg_err": parity_pu_arg_err_json(arg_error, hresult, params.cArgs),
            "result_vartype_after_call": vartype(&result.value),
            "result_ownership_state": "owned VariantOwner; dispatch is cloned before VariantClear and all BSTR/EXCEPINFO storage is released once",
            "cleanup_result": "arguments and named-DISPID storage remain alive through Invoke; result is cleared by VariantOwner Drop",
            "raw_pointer_values_recorded": false,
        });
        match invoke_result {
            Ok(()) => Ok(RawInvocation { result, diagnostic }),
            Err(error) => Err(InvocationFailure { error, diagnostic }),
        }
    }

    fn parity_pu_arg_err_json(raw: u32, hresult: i32, c_args: u32) -> Value {
        const DISP_E_EXCEPTION: i32 = 0x8002_0009_u32 as i32;
        const DISP_E_PARAMNOTFOUND: i32 = 0x8002_0004_u32 as i32;
        const DISP_E_TYPEMISMATCH: i32 = 0x8002_0005_u32 as i32;
        let applicable = matches!(hresult, DISP_E_PARAMNOTFOUND | DISP_E_TYPEMISMATCH);
        let source_parameter_index = applicable
            .then_some(raw)
            .filter(|index| *index < c_args)
            .map(|index| c_args.saturating_sub(index).saturating_sub(1));
        json!({
            "raw_value": if raw == u32::MAX { Value::String("UINT_MAX".to_owned()) } else { Value::from(raw) },
            "applicable": applicable,
            "physical_rgvarg_index": if applicable && raw != u32::MAX { Value::from(raw) } else { Value::Null },
            "source_parameter_index": source_parameter_index,
            "zero_argument_exception_interpretation": if hresult == DISP_E_EXCEPTION && c_args == 0 { "raw sentinel is not a source parameter" } else { "not applicable" },
        })
    }

    fn redact_fixture_path(record: &mut Value) {
        let Some(arguments) = record
            .get_mut("frame")
            .and_then(|frame| frame.get_mut("arguments"))
            .and_then(Value::as_array_mut)
        else {
            return;
        };
        for argument in arguments {
            let Some(value) = argument.get_mut("value") else {
                continue;
            };
            if value.get("kind").and_then(Value::as_str) == Some("bstr") {
                value["value"] = Value::String("<temporary fixture path redacted>".to_owned());
            }
        }
    }

    fn parity_name_dispid(
        dispatch: &IDispatch,
        name: &str,
        lcid: u32,
    ) -> windows::core::Result<i32> {
        let name = HSTRING::from(name);
        let names = [PCWSTR(name.as_ptr())];
        let mut dispid = 0;
        unsafe {
            dispatch.GetIDsOfNames(&GUID::from_u128(0), names.as_ptr(), 1, lcid, &mut dispid)?
        };
        Ok(dispid)
    }

    fn dispatch_type_record_with_lcid(dispatch: &IDispatch, lcid: u32) -> Value {
        let count = unsafe { dispatch.GetTypeInfoCount() };
        match count {
            Ok(count) if count > 0 => {
                let type_name =
                    unsafe { dispatch.GetTypeInfo(0, lcid) }.and_then(|type_info| unsafe {
                        let mut name = BSTR::default();
                        let mut help_context = 0_u32;
                        type_info.GetDocumentation(
                            -1,
                            Some(&mut name),
                            None,
                            &mut help_context,
                            None,
                        )?;
                        Ok::<_, windows::core::Error>((name.to_string(), help_context))
                    });
                match type_name {
                    Ok((name, help_context)) => {
                        json!({"type_info_count":count,"type_name":name,"help_context":help_context,"raw_interface_pointer_recorded":false})
                    }
                    Err(error) => {
                        json!({"type_info_count":count,"type_info_error":hresult_json(error.code().0),"raw_interface_pointer_recorded":false})
                    }
                }
            }
            Ok(count) => {
                json!({"type_info_count":count,"type_name":Value::Null,"raw_interface_pointer_recorded":false})
            }
            Err(error) => {
                json!({"type_info_count_error":hresult_json(error.code().0),"raw_interface_pointer_recorded":false})
            }
        }
    }

    fn blocked_capture(failure: CreateFailure, apartment: Value) -> Capture {
        let pywin32_control = pywin32_control_record();
        let mut observations = vec![pywin32_control.clone()];
        observations.extend(
            failure
                .invocation_diagnostics
                .into_iter()
                .map(invocation_observation),
        );
        observations.push(json!({
            "schema_version": 1,
            "id": "runtime.cleanup.05b-owned-excel",
            "case_id": "cleanup.owned-excel",
            "classification": "Runtime-observed",
            "categories": ["cleanup", "workbooks-add-diagnostic"],
            "environment_id": "excel-16.0.20131.20154-win64-05b",
            "cleanup": failure.cleanup.clone(),
            "raw_pointer_values_recorded": false,
        }));
        Capture {
            manifest: runtime_manifest("not-run-after-raw-workbooks-add-failure"),
            environments: vec![json!({
                "schema_version": 1,
                "id": "excel-16.0.20131.20154-win64-05b",
                "classification": "Version-specific",
                "excel_file_version": "16.0.20131.20154",
                "office_bitness": "64-bit",
                "typelib_guid": "{00020813-0000-0000-C000-000000000046}",
                "typelib_version": "1.9",
                "windows_version": "Windows 10 Enterprise 25H2 build 26200.8875",
                "process_isolation": "CoCreateInstance created an isolated local server; the recorded diagnostic failed before the primary workbook could be retained and no Range call was made.",
                "owned_process": failure.identity,
                "apartment": apartment,
                "cleanup": failure.cleanup,
                "pywin32_control": pywin32_control,
                "source": "isolated-raw-com-probe-05b-diagnostic"
            })],
            observations,
            cases: vec![case(
                "workbooks-add.diagnostic-start",
                "inconclusive",
                "The owned local Excel server rejected the recorded Workbooks.Add diagnostic before a temporary workbook or Range could be created.",
            )],
            unresolved: vec![json!({
                "schema_version": 1,
                "id": "runtime.unresolved.05b-workbooks-add-diagnostic",
                "target": "Excel.Workbooks.Add on dedicated local server",
                "classification": "Inconclusive",
                "detail": failure.message,
                "effect": "No Range runtime observation was captured. The owned-session cleanup path requested Excel.Quit and did not terminate any process."
            })],
        }
    }

    struct Apartment {
        thread_id: u32,
    }
    impl Apartment {
        fn initialize() -> Result<Self, String> {
            unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok() }
                .map_err(|error| format!("cannot initialize the probe COM apartment: {error}"))?;
            Ok(Self {
                thread_id: unsafe { GetCurrentThreadId() },
            })
        }

        fn record(&self) -> Value {
            json!({
                "co_initialize_ex": "COINIT_APARTMENTTHREADED",
                "thread_id": self.thread_id,
                "message_pump_availability": "no custom pump; bounded synchronous local-server calls only",
            })
        }
    }
    impl Drop for Apartment {
        fn drop(&mut self) {
            unsafe { CoUninitialize() };
        }
    }

    #[repr(transparent)]
    struct VariantOwner {
        value: VARIANT,
    }
    impl VariantOwner {
        fn empty() -> Self {
            Self {
                value: unsafe { windows::Win32::System::Variant::VariantInit() },
            }
        }
        fn from_value(value: VARIANT) -> Self {
            Self { value }
        }
        fn null() -> Self {
            let mut value = Self::empty();
            unsafe { variant_data_mut(&mut value.value).vt = VT_NULL };
            value
        }
        fn error(code: i32) -> Self {
            let mut value = Self::empty();
            unsafe {
                let data = variant_data_mut(&mut value.value);
                data.vt = VT_ERROR;
                data.Anonymous.scode = code;
            }
            value
        }
        fn date(value: f64) -> Self {
            let mut result = Self::empty();
            unsafe {
                let data = variant_data_mut(&mut result.value);
                data.vt = VT_DATE;
                data.Anonymous.date = value;
            }
            result
        }
        fn currency(scaled: i64) -> Self {
            let mut result = Self::empty();
            unsafe {
                let data = variant_data_mut(&mut result.value);
                data.vt = VT_CY;
                data.Anonymous.cyVal = CY { int64: scaled };
            }
            result
        }
        fn array(matrix: MatrixInput) -> Result<Self, String> {
            validate_matrix_shape(matrix.rows, matrix.cols, matrix.values.len())?;
            let bounds = [
                SAFEARRAYBOUND {
                    cElements: matrix.rows,
                    lLbound: matrix.lower,
                },
                SAFEARRAYBOUND {
                    cElements: matrix.cols,
                    lLbound: matrix.lower,
                },
            ];
            let array = unsafe { SafeArrayCreate(VT_VARIANT, 2, bounds.as_ptr()) };
            if array.is_null() {
                return Err("SafeArrayCreate returned null".to_owned());
            }
            for row in 0..matrix.rows {
                for col in 0..matrix.cols {
                    let index = usize::try_from(row * matrix.cols + col)
                        .map_err(|_| "matrix index overflow".to_owned())?;
                    let indices = [matrix.lower + row as i32, matrix.lower + col as i32];
                    unsafe {
                        SafeArrayPutElement(
                            array,
                            indices.as_ptr(),
                            &matrix.values[index].value as *const VARIANT
                                as *const core::ffi::c_void,
                        )
                        .map_err(|error| format!("cannot populate SAFEARRAY: {error}"))?;
                    }
                }
            }
            let mut result = Self::empty();
            unsafe {
                let data = variant_data_mut(&mut result.value);
                data.vt = VARENUM(VT_ARRAY.0 | VT_VARIANT.0);
                data.Anonymous.parray = array;
            }
            Ok(result)
        }
    }
    impl Drop for VariantOwner {
        fn drop(&mut self) {
            let _ = unsafe { VariantClear(&mut self.value) };
        }
    }

    unsafe fn variant_data(value: &VARIANT) -> &windows::Win32::System::Variant::VARIANT_0_0 {
        unsafe { &value.Anonymous.Anonymous }
    }

    unsafe fn variant_data_mut(
        value: &mut VARIANT,
    ) -> &mut windows::Win32::System::Variant::VARIANT_0_0 {
        unsafe { &mut value.Anonymous.Anonymous }
    }

    struct MatrixInput {
        rows: u32,
        cols: u32,
        lower: i32,
        values: Vec<VariantOwner>,
    }

    fn validate_matrix_shape(rows: u32, cols: u32, values: usize) -> Result<(), String> {
        let expected = usize::try_from(u64::from(rows) * u64::from(cols))
            .map_err(|_| "matrix shape exceeds addressable memory".to_owned())?;
        if expected != values {
            return Err(format!(
                "matrix shape {rows}x{cols} requires {expected} values, received {values}"
            ));
        }
        Ok(())
    }

    fn logical_row_column(columns: u32, flat_index: usize) -> (u32, u32) {
        let column_count = usize::try_from(columns.max(1)).unwrap_or(1);
        (
            u32::try_from(flat_index / column_count).unwrap_or(u32::MAX),
            u32::try_from(flat_index % column_count).unwrap_or(u32::MAX),
        )
    }

    struct Session {
        app: Option<IDispatch>,
        workbooks: Option<IDispatch>,
        workbook: Option<IDispatch>,
        identity: Value,
        process_handle: windows::Win32::Foundation::HANDLE,
        invocation_diagnostics: Vec<Value>,
        primary_add_failure: Option<String>,
    }

    struct CreateFailure {
        message: String,
        invocation_diagnostics: Vec<Value>,
        identity: Value,
        cleanup: Value,
    }

    impl CreateFailure {
        fn before_session(message: String) -> Self {
            Self {
                message,
                invocation_diagnostics: Vec::new(),
                identity: Value::Null,
                cleanup: json!({
                    "workbook_closed": false,
                    "excel_quit_requested": false,
                    "process_exited": false,
                    "forced_termination": false,
                    "stage": "before-session-creation",
                }),
            }
        }
    }

    impl Session {
        fn create(root: &Path) -> Result<Self, CreateFailure> {
            let class_id = excel_application_clsid(root).map_err(CreateFailure::before_session)?;
            let app =
                unsafe { CoCreateInstance::<_, IDispatch>(&class_id, None, CLSCTX_LOCAL_SERVER) }
                    .map_err(|error| {
                    CreateFailure::before_session(format!(
                        "cannot create a dedicated Excel local server: {error}"
                    ))
                })?;
            let mut session = Self {
                app: Some(app),
                workbooks: None,
                workbook: None,
                identity: Value::Null,
                process_handle: windows::Win32::Foundation::HANDLE::default(),
                invocation_diagnostics: Vec::new(),
                primary_add_failure: None,
            };
            // Match the supplied DispatchEx control through Workbooks.Add:
            // do not alter Application.Visible or DisplayAlerts before the
            // primary diagnostic. Those settings are not needed to own or
            // close this temporary workbook and would add another raw frame
            // before the failing operation.
            match session.verify_owned_process() {
                Ok(identity) => session.identity = identity,
                Err(error) => return Err(session.creation_failure(error)),
            }
            let application = session
                .app
                .as_ref()
                .expect("session has its Application dispatch during creation")
                .clone();
            let workbooks = match session.diagnostic_get_dispatch(
                &application,
                DISPID_APPLICATION_WORKBOOKS,
                "Workbooks",
                "Excel.Application.Workbooks",
                "workbooks.application-property-get",
                "Excel.Application/_Application",
            ) {
                Ok(workbooks) => workbooks,
                Err(error) => return Err(session.creation_failure(error)),
            };
            let workbook = match session.diagnostic_call_dispatch(
                &workbooks,
                DISPID_WORKBOOKS_ADD,
                "Add",
                "Excel.Workbooks.Add",
                "workbooks.add.zero-argument",
                "Excel.Workbooks/Workbooks",
                DISPATCH_METHOD,
                Vec::new(),
            ) {
                Ok(workbook) => Some(workbook),
                Err(primary_error) => {
                    match session.retry_workbooks_add(&workbooks, primary_error) {
                        Ok(workbook) => Some(workbook),
                        Err(error) => {
                            session.primary_add_failure = Some(error);
                            None
                        }
                    }
                }
            };
            session.workbooks = Some(workbooks);
            session.workbook = workbook;
            Ok(session)
        }

        fn creation_failure(&mut self, message: String) -> CreateFailure {
            let cleanup = self.cleanup();
            CreateFailure {
                message,
                invocation_diagnostics: std::mem::take(&mut self.invocation_diagnostics),
                identity: self.identity.clone(),
                cleanup,
            }
        }

        fn retry_workbooks_add(
            &mut self,
            workbooks: &IDispatch,
            primary_error: String,
        ) -> Result<IDispatch, String> {
            let mut attempts = vec![format!("immediate DISPATCH_METHOD: {primary_error}")];
            thread::sleep(Duration::from_millis(500));
            match self.diagnostic_call_dispatch(
                workbooks,
                DISPID_WORKBOOKS_ADD,
                "Add",
                "Excel.Workbooks.Add",
                "workbooks.add.delay-500ms",
                "Excel.Workbooks/Workbooks",
                DISPATCH_METHOD,
                Vec::new(),
            ) {
                Ok(workbook) => return Ok(workbook),
                Err(error) => attempts.push(format!("500 ms delay DISPATCH_METHOD: {error}")),
            }
            pump_messages_for(Duration::from_millis(500));
            match self.diagnostic_call_dispatch(
                workbooks,
                DISPID_WORKBOOKS_ADD,
                "Add",
                "Excel.Workbooks.Add",
                "workbooks.add.message-pump-500ms",
                "Excel.Workbooks/Workbooks",
                DISPATCH_METHOD,
                Vec::new(),
            ) {
                Ok(workbook) => return Ok(workbook),
                Err(error) => {
                    attempts.push(format!("500 ms message pump DISPATCH_METHOD: {error}"))
                }
            }
            thread::sleep(Duration::from_millis(2_000));
            match self.diagnostic_call_dispatch(
                workbooks,
                DISPID_WORKBOOKS_ADD,
                "Add",
                "Excel.Workbooks.Add",
                "workbooks.add.delay-2000ms",
                "Excel.Workbooks/Workbooks",
                DISPATCH_METHOD,
                Vec::new(),
            ) {
                Ok(workbook) => return Ok(workbook),
                Err(error) => attempts.push(format!("2,000 ms delay DISPATCH_METHOD: {error}")),
            }
            match self.diagnostic_call_dispatch(
                workbooks,
                DISPID_WORKBOOKS_ADD,
                "Add",
                "Excel.Workbooks.Add",
                "workbooks.add.combined-method-propertyget",
                "Excel.Workbooks/Workbooks",
                DISPATCH_METHOD | DISPATCH_PROPERTYGET,
                Vec::new(),
            ) {
                Ok(workbook) => Ok(workbook),
                Err(error) => {
                    attempts.push(format!("combined method/property-get: {error}"));
                    Err(attempts.join("; "))
                }
            }
        }

        fn verify_owned_process(&mut self) -> Result<Value, String> {
            let hwnd_value = invoke(
                self.app
                    .as_ref()
                    .ok_or_else(|| "owned Application dispatch was released".to_owned())?,
                name_dispid(
                    self.app
                        .as_ref()
                        .ok_or_else(|| "owned Application dispatch was released".to_owned())?,
                    "Hwnd",
                )
                .map_err(|error| error.to_string())?,
                DISPATCH_PROPERTYGET,
                Vec::new(),
                false,
            )
            .map_err(|error| {
                format!("cannot obtain Hwnd for the created Excel instance: {error}")
            })?;
            let hwnd = scalar_i32(&hwnd_value.value)
                .ok_or_else(|| "Excel Hwnd did not return VT_I4".to_owned())?;
            let mut process_id = 0;
            unsafe { GetWindowThreadProcessId(HWND(hwnd as *mut _), Some(&mut process_id)) };
            if process_id == 0 {
                return Err("Hwnd-to-PID ownership lookup returned zero".to_owned());
            }
            let handle = unsafe {
                OpenProcess(
                    PROCESS_QUERY_LIMITED_INFORMATION | PROCESS_SYNCHRONIZE,
                    false,
                    process_id,
                )
            }
            .map_err(|error| format!("cannot open created Excel PID {process_id}: {error}"))?;
            let mut creation = FILETIME::default();
            unsafe {
                GetProcessTimes(
                    handle,
                    &mut creation,
                    &mut FILETIME::default(),
                    &mut FILETIME::default(),
                    &mut FILETIME::default(),
                )
                .map_err(|error| format!("cannot record created Excel process time: {error}"))?;
            }
            self.process_handle = handle;
            let start_ticks =
                (u64::from(creation.dwHighDateTime) << 32) | u64::from(creation.dwLowDateTime);
            Ok(json!({
                "creation_method": "CoCreateInstance(CLSCTX_LOCAL_SERVER)",
                "pid": process_id,
                "start_time_filetime_ticks": start_ticks,
                "window_verified": true,
                "executable_basename": "EXCEL.EXE",
                "raw_path_recorded": false,
            }))
        }

        fn run_cases(
            &mut self,
            observations: &mut Vec<Value>,
            cases: &mut Vec<Value>,
            control: Option<&Value>,
        ) -> Result<(), String> {
            let sheet = self.get_dispatch(
                self.app
                    .as_ref()
                    .ok_or_else(|| "owned Application dispatch was released".to_owned())?,
                DISPID_APPLICATION_ACTIVE_SHEET,
                "ActiveSheet",
                "Excel.Application.ActiveSheet",
            )?;
            self.range_smoke(&sheet, observations, cases)?;
            self.populate_fixture(&sheet, observations, cases)?;
            self.read_scalars(&sheet, observations, cases, control)?;
            self.read_shapes(&sheet, observations, cases, control)?;
            self.read_formulas(&sheet, observations, cases)?;
            self.write_scalars(&sheet, observations, cases)?;
            self.write_matrices(&sheet, observations, cases)?;
            self.optional_arguments(&sheet, observations, cases)?;
            self.stress(&sheet, observations, cases)?;
            Ok(())
        }

        fn workbooks_add_optional_diagnostics(
            &mut self,
            observations: &mut Vec<Value>,
            cases: &mut Vec<Value>,
        ) -> Result<(), String> {
            cases.push(case(
                "workbooks.add.zero-argument",
                "completed",
                "The primary Workbooks.Add diagnostic used DISPATCH_METHOD with cArgs=0, cNamedArgs=0, and null argument pointers.",
            ));
            let workbooks = self.workbooks.clone().ok_or_else(|| {
                "Workbooks.Add diagnostic lacks the returned Workbooks dispatch object".to_owned()
            })?;
            let matrix = vec![
                (
                    "workbooks.add.missing-marker",
                    "VT_ERROR/DISP_E_PARAMNOTFOUND",
                    VariantOwner::error(DISP_E_PARAMNOTFOUND),
                    "explicit missing-marker argument",
                ),
                (
                    "workbooks.add.empty",
                    "VT_EMPTY",
                    VariantOwner::empty(),
                    "explicit empty argument",
                ),
                (
                    "workbooks.add.null",
                    "VT_NULL",
                    VariantOwner::null(),
                    "explicit null argument",
                ),
                (
                    "workbooks.add.xl-wbat-worksheet",
                    "VT_I4(-4167)",
                    VariantOwner::from_value(VARIANT::from(-4167_i32)),
                    "Excel.XlWBATemplate.xlWBATWorksheet from the committed documentation evidence",
                ),
            ];
            for (case_id, input, argument, source) in matrix {
                match self.diagnostic_member(
                    &workbooks,
                    case_id,
                    "Excel.Workbooks/Workbooks",
                    "Add",
                    DISPID_WORKBOOKS_ADD,
                    DISPATCH_METHOD,
                    InvocationFrame::positional(vec![argument]),
                ) {
                    Ok(mut invocation) => {
                        match dispatch_from_variant(&invocation.result.value) {
                            Some(workbook) => {
                                let retain_for_range = self.workbook.is_none();
                                if retain_for_range {
                                    self.workbook = Some(workbook.clone());
                                }
                                invocation.diagnostic["workbook_created"] = Value::Bool(true);
                                invocation.diagnostic["created_workbook_dispatch"] =
                                    dispatch_type_record(&workbook);
                                invocation.diagnostic["retained_for_range_smoke"] =
                                    Value::Bool(retain_for_range);
                                invocation.diagnostic["cleanup_result"] = if retain_for_range {
                                    json!({
                                        "workbook_close": "deferred to the owned-session cleanup after the Range smoke and matrix",
                                        "forced_termination": false,
                                    })
                                } else {
                                    match invoke(
                                        &workbook,
                                        DISPID_WORKBOOK_CLOSE,
                                        DISPATCH_METHOD,
                                        vec![VariantOwner::from_value(VARIANT::from(false))],
                                        false,
                                    ) {
                                        Ok(value) => json!({
                                            "workbook_close": "succeeded",
                                            "workbook_close_result": describe_variant(&value.value),
                                            "forced_termination": false,
                                        }),
                                        Err(error) => json!({
                                            "workbook_close": "failed",
                                            "hresult": hresult_json(error.code().0),
                                            "forced_termination": false,
                                        }),
                                    }
                                };
                            }
                            None => {
                                invocation.diagnostic["workbook_created"] = Value::Bool(false);
                                invocation.diagnostic["cleanup_result"] = json!({
                                    "workbook_close": "not-attempted; Add did not return VT_DISPATCH",
                                    "forced_termination": false,
                                });
                            }
                        }
                        invocation.diagnostic["optional_argument_representation"] =
                            Value::String(input.to_owned());
                        invocation.diagnostic["optional_argument_source"] =
                            Value::String(source.to_owned());
                        observations.push(invocation_observation(invocation.diagnostic));
                    }
                    Err(_) => observations.extend(self.take_invocation_observations()),
                }
            }
            cases.push(case(
                "workbooks.add.optional-arguments",
                "completed",
                "The zero-argument primary call and each separate missing, empty, null, and xlWBATWorksheet argument representation were recorded without assuming equivalence.",
            ));
            Ok(())
        }

        fn range_smoke(
            &self,
            sheet: &IDispatch,
            observations: &mut Vec<Value>,
            cases: &mut Vec<Value>,
        ) -> Result<(), String> {
            let range = self.range(sheet, "A1")?;
            let write = invoke(
                &range,
                DISPID_RANGE_VALUE2,
                DISPATCH_PROPERTYPUT,
                vec![VariantOwner::from_value(VARIANT::from(42_i32))],
                true,
            )
            .map_err(|error| format!("Range smoke A1.Value2 write failed: {error}"))?;
            let read = invoke(
                &range,
                DISPID_RANGE_VALUE2,
                DISPATCH_PROPERTYGET,
                Vec::new(),
                false,
            )
            .map_err(|error| format!("Range smoke A1.Value2 read failed: {error}"))?;
            let clear = invoke(
                &range,
                DISPID_RANGE_CLEAR,
                DISPATCH_METHOD,
                Vec::new(),
                false,
            )
            .map_err(|error| format!("Range smoke A1.Clear failed: {error}"))?;
            observations.push(json!({
                "schema_version": 1,
                "id": "runtime.range-smoke.a1-value2",
                "case_id": "range.smoke.a1-value2",
                "classification": "Runtime-observed",
                "categories": ["smoke", "range", "value2"],
                "environment_id": "excel-16.0.20131.20154-win64-05b",
                "address": "A1",
                "write": describe_variant(&write.value),
                "read": describe_variant(&read.value),
                "clear": describe_variant(&clear.value),
                "result_ownership": "all raw return VARIANTs were owned and cleared by scoped VariantOwner values",
                "raw_pointer_values_recorded": false,
            }));
            cases.push(case(
                "range.smoke.a1-value2",
                "completed",
                "A1.Value2 was written as VT_I4(42), read through raw IDispatch, then cleared before the full matrix.",
            ));
            Ok(())
        }

        fn populate_fixture(
            &self,
            sheet: &IDispatch,
            observations: &mut Vec<Value>,
            cases: &mut Vec<Value>,
        ) -> Result<(), String> {
            let values = [
                (
                    "A3",
                    VariantOwner::from_value(VARIANT::from("plain ASCII text")),
                ),
                ("A4", VariantOwner::from_value(VARIANT::from("Grüße Ω"))),
                ("A5", VariantOwner::from_value(VARIANT::from(false))),
                ("A6", VariantOwner::from_value(VARIANT::from(true))),
                ("A7", VariantOwner::from_value(VARIANT::from(42_i32))),
                ("A8", VariantOwner::from_value(VARIANT::from(-42_i32))),
                ("A9", VariantOwner::from_value(VARIANT::from(1.25_f64))),
                ("A10", VariantOwner::from_value(VARIANT::from(45_292.5_f64))),
                ("A11", VariantOwner::from_value(VARIANT::from(123.45_f64))),
            ];
            for (address, value) in values {
                observations.push(self.set_observation(
                    "fixture.write",
                    "Excel.Range.Value2",
                    address,
                    "scalar",
                    DISPID_RANGE_VALUE2,
                    value,
                    &["write"],
                    None,
                ));
            }
            for (address, formula) in [
                ("A2", "=\"\""),
                ("A12", "=1/0"),
                ("A13", "=NA()"),
                ("A14", "=1+\"x\""),
                ("K1", "=1+1"),
                ("K2", "=\"formula text\""),
                ("K3", "=TRUE"),
                ("K4", "=1/0"),
            ] {
                observations.push(self.set_observation(
                    "fixture.formula",
                    "Excel.Range.Formula",
                    address,
                    "formula",
                    DISPID_RANGE_FORMULA,
                    VariantOwner::from_value(VARIANT::from(formula)),
                    &["formula"],
                    None,
                ));
            }
            self.set_named_property(
                &self.range(sheet, "A10")?,
                "NumberFormat",
                VariantOwner::from_value(VARIANT::from("m/d/yyyy h:mm")),
            )?;
            self.set_named_property(
                &self.range(sheet, "A11")?,
                "NumberFormat",
                VariantOwner::from_value(VARIANT::from("$#,##0.00")),
            )?;
            self.write_matrix_case(
                sheet,
                "B1:C2",
                "fixture.mixed",
                1,
                mixed_values(),
                observations,
                cases,
            )?;
            self.write_matrix_case(
                sheet,
                "E1:G2",
                "fixture.numeric",
                1,
                numeric_values(),
                observations,
                cases,
            )?;
            self.write_matrix_case(
                sheet,
                "H1:I3",
                "fixture.text",
                1,
                text_values(),
                observations,
                cases,
            )?;
            observations.push(self.set_observation(
                "fixture.dynamic-array",
                "Excel.Range.Formula2",
                "M1",
                "dynamic-array",
                DISPID_RANGE_FORMULA2,
                VariantOwner::from_value(VARIANT::from("=SEQUENCE(2,2)")),
                &["formula"],
                None,
            ));
            cases.push(case(
                "fixture.population",
                "completed",
                "Controlled temporary worksheet cells populated through raw IDispatch calls.",
            ));
            Ok(())
        }

        fn read_scalars(
            &self,
            sheet: &IDispatch,
            observations: &mut Vec<Value>,
            cases: &mut Vec<Value>,
            control: Option<&Value>,
        ) -> Result<(), String> {
            for (address, label) in [
                ("A1", "empty"),
                ("A2", "formula-empty-string"),
                ("A3", "text"),
                ("A4", "unicode"),
                ("A5", "boolean-false"),
                ("A6", "boolean-true"),
                ("A7", "integer"),
                ("A8", "negative-integer"),
                ("A9", "float"),
                ("A10", "date-time"),
                ("A11", "currency"),
                ("A12", "error-div0"),
                ("A13", "error-na"),
                ("A14", "error-value"),
            ] {
                let range = self.range(sheet, address)?;
                observations.push(self.get_observation(
                    &range,
                    "Excel.Range.Value",
                    address,
                    label,
                    DISPID_RANGE_VALUE,
                    Vec::new(),
                    &[
                        "scalar",
                        "value-vs-value2",
                        if label.contains("error") {
                            "error"
                        } else {
                            "scalar"
                        },
                    ],
                    control,
                ));
                observations.push(self.get_observation(
                    &range,
                    "Excel.Range.Value2",
                    address,
                    label,
                    DISPID_RANGE_VALUE2,
                    Vec::new(),
                    &[
                        "scalar",
                        "value-vs-value2",
                        if label.contains("error") {
                            "error"
                        } else {
                            "scalar"
                        },
                    ],
                    control,
                ));
                observations.push(self.get_observation(
                    &range,
                    "Excel.Range.Text",
                    address,
                    label,
                    DISPID_RANGE_TEXT,
                    Vec::new(),
                    &["formula"],
                    None,
                ));
                observations.push(self.get_observation(
                    &range,
                    "Excel.Range.HasFormula",
                    address,
                    label,
                    DISPID_RANGE_HAS_FORMULA,
                    Vec::new(),
                    &["formula"],
                    None,
                ));
            }
            cases.push(case("scalar.reads", "completed", "Value and Value2 were read from controlled scalar cells, including empty, text, Boolean, numeric, formatted date/currency, and formula error cells."));
            Ok(())
        }

        fn read_shapes(
            &self,
            sheet: &IDispatch,
            observations: &mut Vec<Value>,
            cases: &mut Vec<Value>,
            control: Option<&Value>,
        ) -> Result<(), String> {
            for (address, label) in [
                ("B1", "shape-1x1"),
                ("B1:C1", "shape-1x2"),
                ("B1:B2", "shape-2x1"),
                ("B1:C2", "shape-2x2-mixed"),
                ("E1:G2", "shape-2x3-numeric"),
                ("H1:I3", "shape-3x2-text"),
                ("M1:N2", "shape-2x2-dynamic-array"),
            ] {
                let range = self.range(sheet, address)?;
                observations.push(self.get_observation(
                    &range,
                    "Excel.Range.Value",
                    address,
                    label,
                    DISPID_RANGE_VALUE,
                    Vec::new(),
                    &["array", "rectangular", "value-vs-value2"],
                    control,
                ));
                observations.push(self.get_observation(
                    &range,
                    "Excel.Range.Value2",
                    address,
                    label,
                    DISPID_RANGE_VALUE2,
                    Vec::new(),
                    &["array", "rectangular", "value-vs-value2"],
                    control,
                ));
            }
            cases.push(case("shape.reads", "completed", "One-cell, row, column, rectangular, mixed, numeric, text, and dynamic-array shape reads were attempted."));
            Ok(())
        }

        fn read_formulas(
            &self,
            sheet: &IDispatch,
            observations: &mut Vec<Value>,
            cases: &mut Vec<Value>,
        ) -> Result<(), String> {
            for (address, label) in [
                ("K1:K4", "formula-mixed"),
                ("M1:N2", "formula2-dynamic"),
                ("A12:A14", "formula-errors"),
            ] {
                let range = self.range(sheet, address)?;
                observations.push(self.get_observation(
                    &range,
                    "Excel.Range.Formula",
                    address,
                    label,
                    DISPID_RANGE_FORMULA,
                    Vec::new(),
                    &["formula", "array"],
                    None,
                ));
                observations.push(self.get_observation(
                    &range,
                    "Excel.Range.Formula2",
                    address,
                    label,
                    DISPID_RANGE_FORMULA2,
                    Vec::new(),
                    &["formula", "array"],
                    None,
                ));
            }
            cases.push(case("formula.reads", "completed", "Formula and Formula2 getters were read for scalar, mixed, error, and attempted dynamic-array formulas."));
            Ok(())
        }

        fn write_scalars(
            &self,
            sheet: &IDispatch,
            observations: &mut Vec<Value>,
            cases: &mut Vec<Value>,
        ) -> Result<(), String> {
            let writes = vec![
                (
                    "O1",
                    "write-bstr",
                    VariantOwner::from_value(VARIANT::from("written text")),
                ),
                (
                    "O2",
                    "write-unicode",
                    VariantOwner::from_value(VARIANT::from("雪 Ω")),
                ),
                (
                    "O3",
                    "write-bool",
                    VariantOwner::from_value(VARIANT::from(true)),
                ),
                (
                    "O4",
                    "write-i4",
                    VariantOwner::from_value(VARIANT::from(42_i32)),
                ),
                (
                    "O5",
                    "write-r8",
                    VariantOwner::from_value(VARIANT::from(1.25_f64)),
                ),
                ("O6", "write-date", VariantOwner::date(45_292.5)),
                ("O7", "write-currency", VariantOwner::currency(1_234_500)),
                ("O8", "write-empty", VariantOwner::empty()),
                ("O9", "write-null", VariantOwner::null()),
                ("O10", "write-error-na", VariantOwner::error(2042)),
            ];
            for (address, label, value) in writes {
                let range = self.range(sheet, address)?;
                observations.push(self.set_observation(
                    "scalar.write",
                    "Excel.Range.Value2",
                    address,
                    label,
                    DISPID_RANGE_VALUE2,
                    value,
                    &["write", "scalar"],
                    None,
                ));
                observations.push(self.get_observation(
                    &range,
                    "Excel.Range.Value2",
                    address,
                    &format!("{label}-readback"),
                    DISPID_RANGE_VALUE2,
                    Vec::new(),
                    &["write", "scalar"],
                    None,
                ));
            }
            cases.push(case("scalar.writes", "completed", "Selected scalar Automation VARTYPE writes and Value2 read-back were attempted; individual HRESULTs remain in observations."));
            Ok(())
        }

        fn write_matrices(
            &self,
            sheet: &IDispatch,
            observations: &mut Vec<Value>,
            cases: &mut Vec<Value>,
        ) -> Result<(), String> {
            self.write_matrix_case(
                sheet,
                "Q1:R2",
                "write-mixed-1based",
                1,
                mixed_values(),
                observations,
                cases,
            )?;
            self.write_matrix_case(
                sheet,
                "Q4:R5",
                "write-mixed-0based",
                0,
                mixed_values(),
                observations,
                cases,
            )?;
            self.write_matrix_case(
                sheet,
                "T1:V2",
                "write-numeric-2x3",
                1,
                numeric_values(),
                observations,
                cases,
            )?;
            self.write_matrix_case(
                sheet,
                "T4:U6",
                "write-text-3x2",
                1,
                text_values(),
                observations,
                cases,
            )?;
            cases.push(case("matrix.writes", "completed", "Matching rectangular SAFEARRAY(VARIANT) writes with 0- and 1-based input bounds were attempted and read back."));
            Ok(())
        }

        #[allow(clippy::too_many_arguments)]
        fn write_matrix_case(
            &self,
            sheet: &IDispatch,
            address: &str,
            label: &str,
            lower: i32,
            matrix: (u32, u32, Vec<VariantOwner>),
            observations: &mut Vec<Value>,
            _cases: &mut Vec<Value>,
        ) -> Result<(), String> {
            let range = self.range(sheet, address)?;
            let value = VariantOwner::array(MatrixInput {
                rows: matrix.0,
                cols: matrix.1,
                lower,
                values: matrix.2,
            })?;
            observations.push(self.set_observation(
                "matrix.write",
                "Excel.Range.Value2",
                address,
                label,
                DISPID_RANGE_VALUE2,
                value,
                &["write", "rectangular", "array"],
                None,
            ));
            observations.push(self.get_observation(
                &range,
                "Excel.Range.Value2",
                address,
                &format!("{label}-readback"),
                DISPID_RANGE_VALUE2,
                Vec::new(),
                &["write", "rectangular", "array"],
                None,
            ));
            Ok(())
        }

        fn optional_arguments(
            &self,
            sheet: &IDispatch,
            observations: &mut Vec<Value>,
            cases: &mut Vec<Value>,
        ) -> Result<(), String> {
            let range = self.range(sheet, "A10")?;
            observations.push(self.get_observation(
                &range,
                "Excel.Range.Value",
                "A10",
                "value-argument-omitted",
                DISPID_RANGE_VALUE,
                Vec::new(),
                &["optional", "value-vs-value2"],
                None,
            ));
            observations.push(self.get_observation(
                &range,
                "Excel.Range.Value",
                "A10",
                "value-argument-paramnotfound",
                DISPID_RANGE_VALUE,
                vec![VariantOwner::error(DISP_E_PARAMNOTFOUND)],
                &["optional", "value-vs-value2"],
                None,
            ));
            observations.push(self.get_observation(
                &range,
                "Excel.Range.Value",
                "A10",
                "value-argument-empty",
                DISPID_RANGE_VALUE,
                vec![VariantOwner::empty()],
                &["optional", "value-vs-value2"],
                None,
            ));
            observations.push(self.get_observation(
                &range,
                "Excel.Range.Value",
                "A10",
                "value-argument-null",
                DISPID_RANGE_VALUE,
                vec![VariantOwner::null()],
                &["optional", "value-vs-value2"],
                None,
            ));
            cases.push(case("range.value.optional", "completed", "Range.Value optional argument was compared for omission, VT_ERROR/DISP_E_PARAMNOTFOUND, VT_EMPTY, and VT_NULL."));
            Ok(())
        }

        fn stress(
            &self,
            sheet: &IDispatch,
            observations: &mut Vec<Value>,
            cases: &mut Vec<Value>,
        ) -> Result<(), String> {
            let scalar = self.range(sheet, "A7")?;
            let matrix = self.range(sheet, "B1:C2")?;
            for _ in 0..1_000 {
                let _ = invoke(
                    &scalar,
                    DISPID_RANGE_VALUE2,
                    DISPATCH_PROPERTYGET,
                    Vec::new(),
                    false,
                )
                .map_err(|error| format!("scalar stress read failed: {error}"))?;
                let _ = invoke(
                    &matrix,
                    DISPID_RANGE_VALUE2,
                    DISPATCH_PROPERTYGET,
                    Vec::new(),
                    false,
                )
                .map_err(|error| format!("matrix stress read failed: {error}"))?;
            }
            observations.push(json!({
                "schema_version": 1,
                "id": "runtime.stress.read-clear",
                "case_id": "stress.read-clear",
                "subject": "Excel.Range.Value2",
                "range": "A7;B1:C2",
                "categories": ["write"],
                "classification": "Runtime-observed",
                "result": {"hresult":"0x00000000", "variant_type":"stress-completed", "iterations":1000, "cleanup":"VariantClear via scoped owner on every iteration"}
            }));
            cases.push(case("stress.read-clear", "completed", "1,000 scalar and 1,000 small-matrix raw return/VariantClear cycles completed; this is supporting cleanup evidence, not proof of leak freedom."));
            Ok(())
        }

        fn range(&self, sheet: &IDispatch, address: &str) -> Result<IDispatch, String> {
            self.get_dispatch_with_args(
                sheet,
                DISPID_WORKSHEET_RANGE,
                "Range",
                "Excel.Worksheet.Range",
                vec![VariantOwner::from_value(VARIANT::from(address))],
            )
        }

        fn take_invocation_observations(&mut self) -> Vec<Value> {
            std::mem::take(&mut self.invocation_diagnostics)
                .into_iter()
                .map(invocation_observation)
                .collect()
        }

        fn diagnostic_get_dispatch(
            &mut self,
            dispatch: &IDispatch,
            dispid: i32,
            name: &str,
            subject: &str,
            case_id: &str,
            canonical_owner: &str,
        ) -> Result<IDispatch, String> {
            let mut invocation = self.diagnostic_member(
                dispatch,
                case_id,
                canonical_owner,
                name,
                dispid,
                DISPATCH_PROPERTYGET,
                InvocationFrame::positional(Vec::new()),
            )?;
            let result = dispatch_from_variant(&invocation.result.value)
                .ok_or_else(|| format!("{subject} did not return VT_DISPATCH"))?;
            invocation.diagnostic["returned_dispatch"] = dispatch_type_record(&result);
            self.invocation_diagnostics.push(invocation.diagnostic);
            Ok(result)
        }

        #[allow(clippy::too_many_arguments)]
        fn diagnostic_call_dispatch(
            &mut self,
            dispatch: &IDispatch,
            dispid: i32,
            name: &str,
            subject: &str,
            case_id: &str,
            canonical_owner: &str,
            flags: DISPATCH_FLAGS,
            arguments: Vec<VariantOwner>,
        ) -> Result<IDispatch, String> {
            let mut invocation = self.diagnostic_member(
                dispatch,
                case_id,
                canonical_owner,
                name,
                dispid,
                flags,
                InvocationFrame::positional(arguments),
            )?;
            let result = dispatch_from_variant(&invocation.result.value)
                .ok_or_else(|| format!("{subject} did not return VT_DISPATCH"))?;
            invocation.diagnostic["returned_dispatch"] = dispatch_type_record(&result);
            self.invocation_diagnostics.push(invocation.diagnostic);
            Ok(result)
        }

        #[allow(clippy::too_many_arguments)]
        fn diagnostic_member(
            &mut self,
            dispatch: &IDispatch,
            case_id: &str,
            canonical_owner: &str,
            name: &str,
            audited_dispid: i32,
            flags: DISPATCH_FLAGS,
            frame: InvocationFrame,
        ) -> Result<RawInvocation, String> {
            let runtime_dispid = match name_dispid(dispatch, name) {
                Ok(value) => value,
                Err(error) => {
                    self.invocation_diagnostics.push(json!({
                        "schema_version": 1,
                        "case_id": case_id,
                        "canonical_owner": canonical_owner,
                        "member_name": name,
                        "audited_dispid": audited_dispid,
                        "runtime_resolved_dispid": Value::Null,
                        "dispid_matches_audit": Value::Null,
                        "lcid": INVOKE_LCID,
                        "lcid_policy": INVOKE_LCID_POLICY,
                        "get_ids_of_names_hresult": hresult_json(error.code().0),
                        "invoke_skipped": true,
                        "frame": frame.diagnostic(),
                        "raw_pointer_values_recorded": false,
                    }));
                    return Err(format!("cannot resolve {name}: {error}"));
                }
            };
            if runtime_dispid != audited_dispid {
                self.invocation_diagnostics.push(json!({
                    "schema_version": 1,
                    "case_id": case_id,
                    "canonical_owner": canonical_owner,
                    "member_name": name,
                    "audited_dispid": audited_dispid,
                    "runtime_resolved_dispid": runtime_dispid,
                    "dispid_matches_audit": false,
                    "lcid": INVOKE_LCID,
                    "lcid_policy": INVOKE_LCID_POLICY,
                    "get_ids_of_names_hresult": hresult_json(0),
                    "invoke_skipped": true,
                    "frame": frame.diagnostic(),
                    "raw_pointer_values_recorded": false,
                }));
                return Err(format!(
                    "name lookup for {name} returned DISPID {runtime_dispid}, expected reflected DISPID {audited_dispid}"
                ));
            }
            match invoke_with_diagnostic(
                dispatch,
                InvocationContext {
                    case_id,
                    canonical_owner,
                    member_name: name,
                    audited_dispid,
                    runtime_dispid: Some(runtime_dispid),
                    get_ids_hresult: Some(0),
                },
                flags,
                frame,
            ) {
                Ok(result) => Ok(result),
                Err(failure) => {
                    let message =
                        format!("{canonical_owner}.{name} Invoke failed: {}", failure.error);
                    self.invocation_diagnostics.push(failure.diagnostic);
                    Err(message)
                }
            }
        }

        fn get_dispatch(
            &self,
            dispatch: &IDispatch,
            dispid: i32,
            name: &str,
            subject: &str,
        ) -> Result<IDispatch, String> {
            self.get_dispatch_with_args(dispatch, dispid, name, subject, Vec::new())
        }

        fn get_dispatch_with_args(
            &self,
            dispatch: &IDispatch,
            dispid: i32,
            name: &str,
            subject: &str,
            args: Vec<VariantOwner>,
        ) -> Result<IDispatch, String> {
            verify_name(dispatch, name, dispid)?;
            let result = invoke(dispatch, dispid, DISPATCH_PROPERTYGET, args, false)
                .map_err(|error| format!("{subject} failed: {error}"))?;
            dispatch_from_variant(&result.value)
                .ok_or_else(|| format!("{subject} did not return VT_DISPATCH"))
        }

        fn set_named_property(
            &self,
            dispatch: &IDispatch,
            name: &str,
            value: VariantOwner,
        ) -> Result<(), String> {
            let dispid = name_dispid(dispatch, name).map_err(|error| error.to_string())?;
            invoke(dispatch, dispid, DISPATCH_PROPERTYPUT, vec![value], true)
                .map(|_| ())
                .map_err(|error| format!("setting {name} failed: {error}"))
        }

        #[allow(clippy::too_many_arguments)]
        fn get_observation(
            &self,
            dispatch: &IDispatch,
            subject: &str,
            address: &str,
            case_id: &str,
            dispid: i32,
            args: Vec<VariantOwner>,
            categories: &[&str],
            control: Option<&Value>,
        ) -> Value {
            let arguments = args
                .iter()
                .map(|arg| vartype(&arg.value))
                .collect::<Vec<_>>();
            let lookup =
                name_for_subject(subject).and_then(|name| name_dispid(dispatch, name).ok());
            match invoke(dispatch, dispid, DISPATCH_PROPERTYGET, args, false) {
                Ok(result) => observation(
                    case_id,
                    subject,
                    address,
                    dispid,
                    lookup,
                    arguments,
                    describe_variant(&result.value),
                    categories,
                    control,
                ),
                Err(error) => observation(
                    case_id,
                    subject,
                    address,
                    dispid,
                    lookup,
                    arguments,
                    failed_result(&error),
                    categories,
                    control,
                ),
            }
        }

        #[allow(clippy::too_many_arguments)]
        fn set_observation(
            &self,
            prefix: &str,
            subject: &str,
            address: &str,
            case_id: &str,
            dispid: i32,
            value: VariantOwner,
            categories: &[&str],
            control: Option<&Value>,
        ) -> Value {
            let input_type = vartype(&value.value);
            let lookup = name_for_subject(subject).and_then(|name| {
                name_dispid(
                    &self
                        .range_for_observation(subject, address)
                        .unwrap_or_else(|| {
                            self.app
                                .as_ref()
                                .expect("Application dispatch remains live during observations")
                                .clone()
                        }),
                    name,
                )
                .ok()
            });
            let target = self
                .range_for_observation(subject, address)
                .unwrap_or_else(|| {
                    self.app
                        .as_ref()
                        .expect("Application dispatch remains live during observations")
                        .clone()
                });
            let result = match invoke(&target, dispid, DISPATCH_PROPERTYPUT, vec![value], true) {
                Ok(value) => describe_variant(&value.value),
                Err(error) => failed_result(&error),
            };
            let mut record = observation(
                case_id,
                subject,
                address,
                dispid,
                lookup,
                vec![input_type],
                result,
                categories,
                control,
            );
            record["id"] = Value::String(format!("runtime.{prefix}.{case_id}"));
            record
        }

        fn range_for_observation(&self, subject: &str, address: &str) -> Option<IDispatch> {
            if subject.starts_with("Excel.Range") {
                self.get_dispatch(
                    self.app
                        .as_ref()
                        .expect("Application dispatch remains live during observations"),
                    DISPID_APPLICATION_ACTIVE_SHEET,
                    "ActiveSheet",
                    "Excel.Application.ActiveSheet",
                )
                .ok()
                .and_then(|sheet| self.range(&sheet, address).ok())
            } else {
                None
            }
        }

        fn cleanup(&mut self) -> Value {
            let mut workbook_closed = false;
            if let Some(workbook) = self.workbook.take() {
                workbook_closed = invoke(
                    &workbook,
                    DISPID_WORKBOOK_CLOSE,
                    DISPATCH_METHOD,
                    vec![VariantOwner::from_value(VARIANT::from(false))],
                    false,
                )
                .is_ok();
            }
            let quit_requested = self.app.as_ref().is_some_and(|app| {
                invoke(
                    app,
                    DISPID_APPLICATION_QUIT,
                    DISPATCH_METHOD,
                    Vec::new(),
                    false,
                )
                .is_ok()
            });
            self.workbooks.take();
            self.app.take();
            let exited = if !self.process_handle.is_invalid() {
                unsafe { WaitForSingleObject(self.process_handle, 15_000) == WAIT_OBJECT_0 }
            } else {
                false
            };
            if !self.process_handle.is_invalid() {
                let _ = unsafe { CloseHandle(self.process_handle) };
                self.process_handle = windows::Win32::Foundation::HANDLE::default();
            }
            cleanup_record(workbook_closed, quit_requested, exited)
        }
    }

    impl Drop for Session {
        fn drop(&mut self) {
            if let Some(workbook) = self.workbook.take() {
                let _ = invoke(
                    &workbook,
                    DISPID_WORKBOOK_CLOSE,
                    DISPATCH_METHOD,
                    vec![VariantOwner::from_value(VARIANT::from(false))],
                    false,
                );
            }
            if let Some(app) = self.app.as_ref() {
                let _ = invoke(
                    app,
                    DISPID_APPLICATION_QUIT,
                    DISPATCH_METHOD,
                    Vec::new(),
                    false,
                );
            }
            self.workbooks.take();
            self.app.take();
            if !self.process_handle.is_invalid() {
                let _ = unsafe { CloseHandle(self.process_handle) };
            }
        }
    }

    /// An owned `DISPPARAMS` backing store.  Its vectors are never resized after
    /// `params` has exposed their pointers, and empty vectors become null
    /// pointers as required by the Automation call-frame contract.
    struct InvocationFrame {
        arguments: Vec<VariantOwner>,
        named_dispids: Vec<i32>,
        argument_order: &'static str,
    }

    impl InvocationFrame {
        fn positional(mut arguments: Vec<VariantOwner>) -> Self {
            arguments.reverse();
            Self {
                arguments,
                named_dispids: Vec::new(),
                argument_order: "positional arguments reversed into rgvarg",
            }
        }

        fn property_put(value: VariantOwner) -> Self {
            Self {
                arguments: vec![value],
                named_dispids: vec![DISPID_PROPERTYPUT],
                argument_order: "property value is rgvarg[0] and paired with DISPID_PROPERTYPUT",
            }
        }

        #[allow(dead_code)]
        fn property_put_ref(value: VariantOwner) -> Self {
            Self {
                arguments: vec![value],
                named_dispids: vec![DISPID_PROPERTYPUT],
                argument_order: "property reference is rgvarg[0] and paired with DISPID_PROPERTYPUT for DISPATCH_PROPERTYPUTREF",
            }
        }

        #[cfg(test)]
        fn with_named(named: Vec<(i32, VariantOwner)>, mut positional: Vec<VariantOwner>) -> Self {
            positional.reverse();
            let mut arguments = Vec::with_capacity(named.len() + positional.len());
            let mut named_dispids = Vec::with_capacity(named.len());
            for (dispid, value) in named {
                named_dispids.push(dispid);
                arguments.push(value);
            }
            arguments.extend(positional);
            Self {
                arguments,
                named_dispids,
                argument_order: "named values lead rgvarg in named-DISPID order; positional values follow in reverse order",
            }
        }

        fn params(&mut self) -> DISPPARAMS {
            DISPPARAMS {
                rgvarg: if self.arguments.is_empty() {
                    std::ptr::null_mut()
                } else {
                    self.arguments.as_mut_ptr().cast::<VARIANT>()
                },
                rgdispidNamedArgs: if self.named_dispids.is_empty() {
                    std::ptr::null_mut()
                } else {
                    self.named_dispids.as_mut_ptr()
                },
                cArgs: u32::try_from(self.arguments.len()).unwrap_or(u32::MAX),
                cNamedArgs: u32::try_from(self.named_dispids.len()).unwrap_or(u32::MAX),
            }
        }

        fn diagnostic(&self) -> Value {
            json!({
                "c_args": self.arguments.len(),
                "c_named_args": self.named_dispids.len(),
                "rgvarg_is_null": self.arguments.is_empty(),
                "rgdispid_named_args_is_null": self.named_dispids.is_empty(),
                "argument_order": self.argument_order,
                "arguments": self.arguments.iter().map(|argument| json!({
                    "vartype": vartype(&argument.value),
                    "value": scalar_value(&argument.value),
                })).collect::<Vec<_>>(),
                "named_dispids": self.named_dispids,
            })
        }
    }

    struct InvocationContext<'a> {
        case_id: &'a str,
        canonical_owner: &'a str,
        member_name: &'a str,
        audited_dispid: i32,
        runtime_dispid: Option<i32>,
        get_ids_hresult: Option<i32>,
    }

    struct RawInvocation {
        result: VariantOwner,
        diagnostic: Value,
    }

    struct InvocationFailure {
        error: windows::core::Error,
        diagnostic: Value,
    }

    fn invoke(
        dispatch: &IDispatch,
        dispid: i32,
        flags: DISPATCH_FLAGS,
        arguments: Vec<VariantOwner>,
        property_put: bool,
    ) -> windows::core::Result<VariantOwner> {
        let frame = if property_put {
            let mut arguments = arguments;
            if arguments.len() != 1 {
                return Err(windows::core::Error::new(
                    windows::core::HRESULT(0x8007_0057_u32 as i32),
                    "research probe property put requires exactly one value",
                ));
            }
            InvocationFrame::property_put(arguments.pop().expect("checked length"))
        } else {
            InvocationFrame::positional(arguments)
        };
        invoke_with_diagnostic(
            dispatch,
            InvocationContext {
                case_id: "unrecorded.raw-invocation",
                canonical_owner: "unrecorded dispatch target",
                member_name: "unrecorded",
                audited_dispid: dispid,
                runtime_dispid: None,
                get_ids_hresult: None,
            },
            flags,
            frame,
        )
        .map(|result| result.result)
        .map_err(|failure| failure.error)
    }

    fn invoke_with_diagnostic(
        dispatch: &IDispatch,
        context: InvocationContext<'_>,
        flags: DISPATCH_FLAGS,
        mut frame: InvocationFrame,
    ) -> Result<RawInvocation, InvocationFailure> {
        let frame_record = frame.diagnostic();
        let params = frame.params();
        let mut result = VariantOwner::empty();
        let mut exception = EXCEPINFO::default();
        let mut arg_error = u32::MAX;
        let invoke_result = unsafe {
            dispatch.Invoke(
                context.audited_dispid,
                &GUID::from_u128(0),
                INVOKE_LCID,
                flags,
                &params,
                Some(&mut result.value),
                Some(&mut exception),
                Some(&mut arg_error),
            )
        };
        let exception_record = normalize_exception(&mut exception);
        let hresult = invoke_result
            .as_ref()
            .map(|_| 0_i32)
            .map_err(|error| error.code().0)
            .unwrap_or_else(|value| value);
        let diagnostic = json!({
            "schema_version": 1,
            "case_id": context.case_id,
            "canonical_owner": context.canonical_owner,
            "member_name": context.member_name,
            "audited_dispid": context.audited_dispid,
            "runtime_resolved_dispid": context.runtime_dispid,
            "dispid_matches_audit": dispid_matches(context.audited_dispid, context.runtime_dispid),
            "get_ids_of_names_hresult": context.get_ids_hresult.map(hresult_json),
            "lcid": INVOKE_LCID,
            "lcid_policy": INVOKE_LCID_POLICY,
            "invoke_flags": format_dispatch_flags(flags),
            "invoke_flags_raw": flags.0,
            "frame": frame_record,
            "result_variant_initialized_before_call": true,
            "returned_hresult": hresult_json(hresult),
            "excepinfo": exception_record,
            "pu_arg_err": pu_arg_err_json(arg_error),
            "result_vartype_after_call": vartype(&result.value),
            "result_ownership_state": "owned VariantOwner; VariantClear runs exactly once when the result leaves scope",
            "cleanup_result": "arguments and named-DISPID storage remain alive through Invoke; result is cleared by VariantOwner Drop",
            "raw_pointer_values_recorded": false,
        });
        match invoke_result {
            Ok(()) => Ok(RawInvocation { result, diagnostic }),
            Err(error) => Err(InvocationFailure { error, diagnostic }),
        }
    }

    fn hresult_json(value: i32) -> Value {
        let raw = value as u32;
        json!({
            "hex": format!("0x{raw:08X}"),
            "signed_i32": value,
            "severity": raw >> 31,
            "facility": (raw >> 16) & 0x1fff,
            "code": raw & 0xffff,
        })
    }

    fn pu_arg_err_json(value: u32) -> Value {
        if value == u32::MAX {
            Value::Null
        } else {
            json!({"rgvarg_index": value, "index_is_physical_reverse_order": true})
        }
    }

    fn dispid_matches(audited: i32, runtime: Option<i32>) -> Value {
        runtime.map_or(Value::Null, |value| Value::Bool(value == audited))
    }

    fn cleanup_record(workbook_closed: bool, quit_requested: bool, process_exited: bool) -> Value {
        json!({
            "workbook_closed": workbook_closed,
            "excel_quit_requested": quit_requested,
            "process_exited": process_exited,
            "verification_timeout_milliseconds": 15000,
            "forced_termination": false,
        })
    }

    fn format_dispatch_flags(flags: DISPATCH_FLAGS) -> String {
        let mut names = Vec::new();
        if flags.0 & DISPATCH_METHOD.0 != 0 {
            names.push("DISPATCH_METHOD");
        }
        if flags.0 & DISPATCH_PROPERTYGET.0 != 0 {
            names.push("DISPATCH_PROPERTYGET");
        }
        if flags.0 & DISPATCH_PROPERTYPUT.0 != 0 {
            names.push("DISPATCH_PROPERTYPUT");
        }
        if flags.0 & DISPATCH_PROPERTYPUTREF.0 != 0 {
            names.push("DISPATCH_PROPERTYPUTREF");
        }
        if names.is_empty() {
            names.push("none");
        }
        format!("{} (0x{:04X})", names.join(" | "), flags.0)
    }

    fn pump_messages_for(duration: Duration) {
        let deadline = Instant::now() + duration;
        while Instant::now() < deadline {
            let mut message = MSG::default();
            while unsafe { PeekMessageW(&mut message, None, 0, 0, PM_REMOVE).as_bool() } {
                unsafe {
                    let _ = TranslateMessage(&message);
                    DispatchMessageW(&message);
                }
            }
            thread::sleep(Duration::from_millis(10));
        }
    }

    fn normalize_exception(exception: &mut EXCEPINFO) -> Value {
        let deferred_present = exception.pfnDeferredFillIn.is_some();
        let deferred_result = exception
            .pfnDeferredFillIn
            .map(|fill| unsafe { fill(exception as *mut EXCEPINFO).0 });
        let source = unsafe { take_exception_bstr(&mut exception.bstrSource) };
        let description = unsafe { take_exception_bstr(&mut exception.bstrDescription) };
        let help_file = unsafe { take_exception_bstr(&mut exception.bstrHelpFile) };
        json!({
            "deferred_fill_in_present": deferred_present,
            "deferred_fill_in_hresult": deferred_result.map(hresult_json),
            "wcode": exception.wCode,
            "scode": hresult_json(exception.scode),
            "source": source,
            "description": description,
            "help_file": help_file,
            "help_context": exception.dwHelpContext,
        })
    }

    unsafe fn take_exception_bstr(value: &mut ManuallyDrop<BSTR>) -> Option<String> {
        let bstr = unsafe { ManuallyDrop::into_inner(std::ptr::read(value)) };
        let text = normalize_evidence_text(bstr.to_string());
        if text.is_empty() { None } else { Some(text) }
    }

    fn normalize_evidence_text(text: String) -> String {
        text.replace("â€¢", "•")
            .replace("â†’", "→")
            .replace("â€™", "’")
            .replace("ï¿½", "[replacement-character]")
            .replace('\u{FFFD}', "[replacement-character]")
    }

    fn name_dispid(dispatch: &IDispatch, name: &str) -> windows::core::Result<i32> {
        let name = HSTRING::from(name);
        let names = [PCWSTR(name.as_ptr())];
        let mut dispid = 0;
        unsafe {
            dispatch.GetIDsOfNames(
                &GUID::from_u128(0),
                names.as_ptr(),
                1,
                INVOKE_LCID,
                &mut dispid,
            )?
        };
        Ok(dispid)
    }

    fn verify_name(dispatch: &IDispatch, name: &str, expected: i32) -> Result<(), String> {
        let actual = name_dispid(dispatch, name)
            .map_err(|error| format!("cannot resolve {name}: {error}"))?;
        if actual != expected {
            return Err(format!(
                "name lookup for {name} returned DISPID {actual}, expected reflected DISPID {expected}"
            ));
        }
        Ok(())
    }

    fn excel_application_clsid(root: &Path) -> Result<GUID, String> {
        let records = read_jsonl(&root.join("typelib/coclasses.jsonl"))?;
        let guid = records
            .iter()
            .find(|record| value_string(record, "id") == "Excel.Application.Coclass")
            .map(|record| value_string(record, "guid"))
            .ok_or_else(|| "Prompt 04 Application coclass evidence is unavailable".to_owned())?;
        parse_guid(&guid)
    }

    fn parse_guid(value: &str) -> Result<GUID, String> {
        let digits = value.trim_matches(['{', '}']).replace('-', "");
        let raw = u128::from_str_radix(&digits, 16)
            .map_err(|error| format!("invalid typelib GUID {value}: {error}"))?;
        Ok(GUID::from_u128(raw))
    }

    fn vartype(value: &VARIANT) -> String {
        let raw = unsafe { variant_data(value).vt.0 };
        vartype_name(raw)
    }

    fn vartype_name(raw: u16) -> String {
        let base = raw & !(VT_ARRAY.0 | 0x4000 | 0x8000);
        let name = match base {
            value if value == VT_EMPTY.0 => "VT_EMPTY",
            value if value == VT_NULL.0 => "VT_NULL",
            value if value == VT_I4.0 => "VT_I4",
            value if value == VT_R8.0 => "VT_R8",
            value if value == VT_CY.0 => "VT_CY",
            value if value == VT_DATE.0 => "VT_DATE",
            value if value == VT_BSTR.0 => "VT_BSTR",
            value if value == VT_DISPATCH.0 => "VT_DISPATCH",
            value if value == VT_ERROR.0 => "VT_ERROR",
            value if value == VT_BOOL.0 => "VT_BOOL",
            value if value == VT_VARIANT.0 => "VT_VARIANT",
            _ => "VT_UNKNOWN_RAW",
        };
        let mut result = Vec::new();
        if raw & VT_ARRAY.0 != 0 {
            result.push("VT_ARRAY");
        }
        if raw & 0x4000 != 0 {
            result.push("VT_BYREF");
        }
        result.push(name);
        format!("{} (0x{raw:04X})", result.join("|"))
    }

    fn describe_variant(value: &VARIANT) -> Value {
        let raw = unsafe { variant_data(value).vt.0 };
        let mut result = json!({
            "hresult": "0x00000000",
            "variant_type": vartype_name(raw),
            "variant_type_raw": raw,
        });
        if raw & VT_ARRAY.0 != 0 {
            result["array"] = inspect_array(value);
        } else {
            result["scalar"] = scalar_value(value);
        }
        result
    }

    fn failed_result(error: &windows::core::Error) -> Value {
        json!({
            "hresult": format!("0x{:08X}", error.code().0 as u32),
            "variant_type": "no-result",
            "error": error.to_string(),
        })
    }

    fn scalar_value(value: &VARIANT) -> Value {
        let raw = unsafe { variant_data(value).vt.0 };
        let data = unsafe { variant_data(value) };
        unsafe {
            match raw {
                tag if tag == VT_EMPTY.0 => json!({"kind":"empty"}),
                tag if tag == VT_NULL.0 => json!({"kind":"null"}),
                tag if tag == VT_I4.0 => {
                    json!({"kind":"i4", "value": data.Anonymous.lVal})
                }
                tag if tag == VT_R8.0 => {
                    json!({"kind":"r8", "value": data.Anonymous.dblVal})
                }
                tag if tag == VT_DATE.0 => {
                    json!({"kind":"date", "value": data.Anonymous.date})
                }
                tag if tag == VT_CY.0 => {
                    json!({"kind":"currency", "scaled_10000": data.Anonymous.cyVal.int64})
                }
                tag if tag == VT_BOOL.0 => {
                    json!({"kind":"bool", "raw": data.Anonymous.boolVal.0})
                }
                tag if tag == VT_ERROR.0 => {
                    let signed = data.Anonymous.scode;
                    error_code_json(signed)
                }
                tag if tag == VT_BSTR.0 => {
                    let bstr = &data.Anonymous.bstrVal as *const ManuallyDrop<BSTR> as *const BSTR;
                    json!({"kind":"bstr", "value": (*bstr).to_string()})
                }
                tag if tag == VT_DISPATCH.0 => {
                    json!({"kind":"dispatch", "pointer_recorded": false})
                }
                _ => json!({"kind":"unknown", "vartype_raw": raw}),
            }
        }
    }

    fn scalar_i32(value: &VARIANT) -> Option<i32> {
        let raw = unsafe { variant_data(value).vt.0 };
        if raw == VT_I4.0 {
            Some(unsafe { variant_data(value).Anonymous.lVal })
        } else {
            None
        }
    }

    fn error_code_json(signed: i32) -> Value {
        json!({"kind":"error", "signed_i32": signed, "unsigned_u32": signed as u32})
    }

    fn dispatch_from_variant(value: &VARIANT) -> Option<IDispatch> {
        if unsafe { variant_data(value).vt.0 } != VT_DISPATCH.0 {
            return None;
        }
        unsafe {
            let pointer = &variant_data(value).Anonymous.pdispVal
                as *const ManuallyDrop<Option<IDispatch>>
                as *const Option<IDispatch>;
            (*pointer).clone()
        }
    }

    fn dispatch_type_record(dispatch: &IDispatch) -> Value {
        let count = unsafe { dispatch.GetTypeInfoCount() };
        match count {
            Ok(count) if count > 0 => {
                let type_name =
                    unsafe { dispatch.GetTypeInfo(0, INVOKE_LCID) }.and_then(|type_info| unsafe {
                        let mut name = BSTR::default();
                        let mut help_context = 0_u32;
                        type_info.GetDocumentation(
                            -1,
                            Some(&mut name),
                            None,
                            &mut help_context,
                            None,
                        )?;
                        Ok::<_, windows::core::Error>((name.to_string(), help_context))
                    });
                match type_name {
                    Ok((name, help_context)) => json!({
                        "type_info_count": count,
                        "type_name": name,
                        "help_context": help_context,
                        "raw_interface_pointer_recorded": false,
                    }),
                    Err(error) => json!({
                        "type_info_count": count,
                        "type_info_error": hresult_json(error.code().0),
                        "raw_interface_pointer_recorded": false,
                    }),
                }
            }
            Ok(count) => json!({
                "type_info_count": count,
                "type_name": Value::Null,
                "raw_interface_pointer_recorded": false,
            }),
            Err(error) => json!({
                "type_info_count_error": hresult_json(error.code().0),
                "raw_interface_pointer_recorded": false,
            }),
        }
    }

    fn inspect_array(value: &VARIANT) -> Value {
        let array = unsafe { variant_data(value).Anonymous.parray };
        if array.is_null() {
            return json!({"presence":false});
        }
        let dimensions = unsafe { SafeArrayGetDim(array) };
        let mut bounds = Vec::new();
        for dimension in 1..=dimensions {
            let lower = unsafe { SafeArrayGetLBound(array, dimension) }.unwrap_or(i32::MIN);
            let upper = unsafe { SafeArrayGetUBound(array, dimension) }.unwrap_or(i32::MAX);
            bounds.push(json!({"dimension":dimension, "lower":lower, "upper":upper}));
        }
        let element_type = unsafe { SafeArrayGetVartype(array) }
            .map(|value| vartype_name(value.0))
            .unwrap_or_else(|_| "not-reported".to_owned());
        let preview = array_preview(array, &bounds);
        json!({
            "presence":true,
            "dimensions":dimensions,
            "bounds":bounds,
            "element_vartype":element_type,
            "logical_mapping":"recorded coordinate preview uses SafeArrayGetElement indices [dimension1, dimension2]; compare fixture values before generalising descriptor order",
            "elements":preview,
        })
    }

    fn array_preview(
        array: *mut windows::Win32::System::Com::SAFEARRAY,
        bounds: &[Value],
    ) -> Vec<Value> {
        if bounds.len() != 2 {
            return Vec::new();
        }
        let row_lower = bounds[0]
            .get("lower")
            .and_then(Value::as_i64)
            .unwrap_or_default() as i32;
        let row_upper = bounds[0]
            .get("upper")
            .and_then(Value::as_i64)
            .unwrap_or(row_lower.into()) as i32;
        let col_lower = bounds[1]
            .get("lower")
            .and_then(Value::as_i64)
            .unwrap_or_default() as i32;
        let col_upper = bounds[1]
            .get("upper")
            .and_then(Value::as_i64)
            .unwrap_or(col_lower.into()) as i32;
        let mut values = Vec::new();
        for row in row_lower..=row_upper.min(row_lower.saturating_add(3)) {
            for col in col_lower..=col_upper.min(col_lower.saturating_add(3)) {
                let indices = [row, col];
                let mut value = VariantOwner::empty();
                let result = unsafe {
                    SafeArrayGetElement(
                        array,
                        indices.as_ptr(),
                        &mut value.value as *mut VARIANT as *mut core::ffi::c_void,
                    )
                };
                let logical = logical_row_column(
                    u32::try_from(col_upper.saturating_sub(col_lower).saturating_add(1))
                        .unwrap_or(u32::MAX),
                    values.len(),
                );
                values.push(match result {
                    Ok(()) => json!({"indices":[row,col], "logical_row":logical.0, "logical_column":logical.1, "value":scalar_value(&value.value), "vartype":vartype(&value.value)}),
                    Err(error) => json!({"indices":[row,col], "error":format!("0x{:08X}", error.code().0 as u32)}),
                });
            }
        }
        values
    }

    #[allow(clippy::too_many_arguments)]
    fn observation(
        case_id: &str,
        subject: &str,
        range: &str,
        dispid: i32,
        lookup: Option<i32>,
        argument_types: Vec<String>,
        result: Value,
        categories: &[&str],
        control: Option<&Value>,
    ) -> Value {
        json!({
            "schema_version":1,
            "id":format!("runtime.{}.{}", subject.replace("Excel.Range.", "range.").to_lowercase(), case_id),
            "case_id":case_id,
            "subject":subject,
            "operation": if subject.contains(".Value") || subject.contains(".Formula") { "property_get" } else { "invoke" },
            "range":range,
            "dispid":dispid,
            "name_lookup_dispid":lookup,
            "argument_count":argument_types.len(),
            "argument_vartypes":argument_types,
            "result":result,
            "categories":categories,
            "environment_id":"excel-16.0.20131.20154-win64-05b",
            "probe_version":PROBE_VERSION,
            "control":control,
            "classification":"Runtime-observed",
            "cleanup":"returned VARIANT deep-copied into JSON evidence before scoped VariantClear; no pointer address recorded"
        })
    }

    fn invocation_observation(mut diagnostic: Value) -> Value {
        let case_id = diagnostic
            .get("case_id")
            .and_then(Value::as_str)
            .unwrap_or("unknown-invocation");
        diagnostic["id"] = Value::String(format!("runtime.invocation.{case_id}"));
        diagnostic["classification"] = Value::String("Runtime-observed".to_owned());
        diagnostic["categories"] = json!(["invocation", "workbooks-add-diagnostic"]);
        diagnostic["environment_id"] = Value::String("excel-16.0.20131.20154-win64-05b".to_owned());
        diagnostic["probe_version"] = Value::from(PROBE_VERSION);
        diagnostic
    }

    fn pywin32_control_record() -> Value {
        json!({
            "schema_version": 1,
            "id": "runtime.control.pywin32-dispatchex-workbooks-add",
            "classification": "Control-confirmed",
            "categories": ["control", "workbooks-add-diagnostic"],
            "client": "Python pywin32",
            "activation": "win32com.client.DispatchEx(\"Excel.Application\")",
            "excel_version": "16.0",
            "workbooks_add": "succeeded",
            "created_workbook": "Book1",
            "hwnd_recorded": false,
            "raw_variant_or_safearray_inference": false,
            "evidence_origin": "Prompt 05B independent control supplied by the user",
        })
    }

    fn case(id: &str, status: &str, detail: &str) -> Value {
        json!({"schema_version":1,"id":format!("runtime.case.{id}"),"status":status,"detail":detail,"classification":if status == "completed" {"Runtime-observed"} else {"Inconclusive"}})
    }

    fn name_for_subject(subject: &str) -> Option<&'static str> {
        match subject {
            "Excel.Range.Value" => Some("Value"),
            "Excel.Range.Value2" => Some("Value2"),
            "Excel.Range.Formula" => Some("Formula"),
            "Excel.Range.Formula2" => Some("Formula2"),
            "Excel.Range.Text" => Some("Text"),
            "Excel.Range.HasFormula" => Some("HasFormula"),
            _ => None,
        }
    }

    fn run_control(script: Option<&Path>) -> Option<Value> {
        let script = script?;
        let output = Command::new("powershell.exe")
            .args(["-NoProfile", "-ExecutionPolicy", "Bypass", "-File"])
            .arg(script)
            .output()
            .ok()?;
        if !output.status.success() {
            return Some(
                json!({"version":"csharp-dynamic-control", "status":"failed", "exit_code":output.status.code()}),
            );
        }
        serde_json::from_slice(&output.stdout).ok().or_else(|| {
            Some(json!({"version":"csharp-dynamic-control", "status":"unparseable-control-output"}))
        })
    }

    fn standard_unresolved() -> Vec<Value> {
        [
            ("range.find.sort", "Range.Find and Range.Sort stateful optional combinations were deliberately not exercised; they require a separate restoration-focused experiment."),
            ("range.address", "Range.Address omission/VT_EMPTY/VT_NULL behaviour was not exercised in this transport-focused run."),
            ("jagged.rank-one.empty-arrays", "Jagged, rank-one, zero-length, and empty SAFEARRAY write forms remain not tested."),
            ("formula.locale", "Locale-specific FormulaLocal and list-separator differences remain not tested by the raw probe."),
            ("embedded-nul", "Embedded-NUL BSTR acceptance remains not tested; no coercion is inferred."),
        ]
        .into_iter()
        .map(|(target, detail)| json!({"schema_version":1,"id":format!("runtime.unresolved.{target}"),"target":target,"classification":"Not tested","detail":detail}))
        .collect()
    }

    fn mixed_values() -> (u32, u32, Vec<VariantOwner>) {
        (
            2,
            2,
            vec![
                VariantOwner::from_value(VARIANT::from("mixed")),
                VariantOwner::from_value(VARIANT::from(42_i32)),
                VariantOwner::from_value(VARIANT::from(true)),
                VariantOwner::empty(),
            ],
        )
    }
    fn numeric_values() -> (u32, u32, Vec<VariantOwner>) {
        (
            2,
            3,
            vec![1_i32, 2, 3, 4, 5, 6]
                .into_iter()
                .map(|value| VariantOwner::from_value(VARIANT::from(value)))
                .collect(),
        )
    }
    fn text_values() -> (u32, u32, Vec<VariantOwner>) {
        (
            3,
            2,
            ["r1c1", "r1c2", "r2c1", "r2c2", "r3c1", "r3c2"]
                .into_iter()
                .map(|value| VariantOwner::from_value(VARIANT::from(value)))
                .collect(),
        )
    }

    #[cfg(test)]
    mod tests {
        use super::*;

        #[test]
        fn vartype_format_retains_modifiers_and_unknown_values() {
            assert_eq!(
                vartype_name(VT_ARRAY.0 | VT_VARIANT.0),
                "VT_ARRAY|VT_VARIANT (0x200C)"
            );
            assert_eq!(
                vartype_name(0x7777),
                "VT_ARRAY|VT_BYREF|VT_UNKNOWN_RAW (0x7777)"
            );
        }

        #[test]
        fn logical_mapping_is_row_major_for_recorded_preview_order() {
            assert_eq!(logical_row_column(3, 0), (0, 0));
            assert_eq!(logical_row_column(3, 4), (1, 1));
        }

        #[test]
        fn raw_excel_error_codes_preserve_both_signednesses() {
            let value = error_code_json(2042);
            assert_eq!(value["signed_i32"], 2042);
            assert_eq!(value["unsigned_u32"], 2042);
            let missing = error_code_json(DISP_E_PARAMNOTFOUND);
            assert_eq!(missing["unsigned_u32"], 0x8002_0004_u32);
        }

        #[test]
        fn matrix_shape_validation_rejects_mismatch_before_allocation() {
            assert!(validate_matrix_shape(2, 2, 4).is_ok());
            assert!(validate_matrix_shape(2, 2, 3).is_err());
        }

        #[test]
        fn scalar_extraction_preserves_unknown_tag() {
            let value = VariantOwner::error(2042);
            assert_eq!(scalar_value(&value.value)["kind"], "error");
            let unknown = VariantOwner::empty();
            assert_eq!(scalar_value(&unknown.value)["kind"], "empty");
        }

        #[test]
        fn zero_argument_frame_uses_null_pointers() {
            let mut frame = InvocationFrame::positional(Vec::new());
            let params = frame.params();
            assert_eq!(params.cArgs, 0);
            assert_eq!(params.cNamedArgs, 0);
            assert!(params.rgvarg.is_null());
            assert!(params.rgdispidNamedArgs.is_null());
        }

        #[test]
        fn one_and_many_positional_arguments_are_reversed() {
            let mut frame = InvocationFrame::positional(vec![
                VariantOwner::from_value(VARIANT::from(1_i32)),
                VariantOwner::from_value(VARIANT::from(2_i32)),
            ]);
            let params = frame.params();
            assert_eq!(params.cArgs, 2);
            assert_eq!(scalar_i32(&frame.arguments[0].value), Some(2));
            assert_eq!(scalar_i32(&frame.arguments[1].value), Some(1));
        }

        #[test]
        fn named_arguments_lead_the_reverse_positional_layout() {
            let frame = InvocationFrame::with_named(
                vec![(77, VariantOwner::from_value(VARIANT::from(9_i32)))],
                vec![
                    VariantOwner::from_value(VARIANT::from(1_i32)),
                    VariantOwner::from_value(VARIANT::from(2_i32)),
                ],
            );
            assert_eq!(frame.named_dispids, vec![77]);
            assert_eq!(scalar_i32(&frame.arguments[0].value), Some(9));
            assert_eq!(scalar_i32(&frame.arguments[1].value), Some(2));
            assert_eq!(scalar_i32(&frame.arguments[2].value), Some(1));
        }

        #[test]
        fn flags_hresult_exceptions_and_argerr_have_stable_diagnostics() {
            assert_eq!(
                format_dispatch_flags(DISPATCH_METHOD | DISPATCH_PROPERTYGET),
                "DISPATCH_METHOD | DISPATCH_PROPERTYGET (0x0003)"
            );
            assert_eq!(hresult_json(0x8002_0004_u32 as i32)["facility"], 2);
            assert_eq!(hresult_json(0x8002_0004_u32 as i32)["code"], 4);
            let mut exception = EXCEPINFO {
                wCode: 7,
                scode: 0x800A_03EC_u32 as i32,
                ..Default::default()
            };
            let normalized = normalize_exception(&mut exception);
            assert_eq!(normalized["wcode"], 7);
            assert_eq!(normalized["scode"]["hex"], "0x800A03EC");
            assert_eq!(pu_arg_err_json(3)["rgvarg_index"], 3);
            assert!(pu_arg_err_json(u32::MAX).is_null());
        }

        #[test]
        fn audit_comparison_and_result_initialization_are_explicit() {
            assert_eq!(dispid_matches(181, Some(181)), Value::Bool(true));
            assert_eq!(dispid_matches(181, Some(180)), Value::Bool(false));
            assert!(dispid_matches(181, None).is_null());
            let result = VariantOwner::empty();
            assert_eq!(vartype(&result.value), "VT_EMPTY (0x0000)");
        }

        #[test]
        fn frame_diagnostic_serialization_omits_pointer_values() {
            let frame =
                InvocationFrame::property_put(VariantOwner::from_value(VARIANT::from(42_i32)));
            let diagnostic = frame.diagnostic();
            assert_eq!(diagnostic["c_args"], 1);
            assert_eq!(diagnostic["c_named_args"], 1);
            assert!(diagnostic.get("rgvarg").is_none());
            assert!(diagnostic.get("rgdispidNamedArgs").is_none());
        }

        #[test]
        fn property_put_and_property_put_ref_use_the_propertyput_named_dispid() {
            let put = InvocationFrame::property_put(VariantOwner::from_value(VARIANT::from(42_i32)));
            let put_ref = InvocationFrame::property_put_ref(VariantOwner::null());
            assert_eq!(put.named_dispids, vec![DISPID_PROPERTYPUT]);
            assert_eq!(put_ref.named_dispids, vec![DISPID_PROPERTYPUT]);
            assert!(put_ref.argument_order.contains("DISPATCH_PROPERTYPUTREF"));
            assert!(format_dispatch_flags(DISPATCH_PROPERTYPUTREF).contains("DISPATCH_PROPERTYPUTREF"));
        }

        #[test]
        fn optional_argument_encodings_remain_distinct() {
            let omitted = InvocationFrame::positional(Vec::new()).diagnostic();
            let missing = VariantOwner::error(DISP_E_PARAMNOTFOUND);
            let empty = VariantOwner::empty();
            let null = VariantOwner::null();
            assert_eq!(omitted["c_args"], 0);
            assert_eq!(vartype(&missing.value), "VT_ERROR (0x000A)");
            assert_eq!(vartype(&empty.value), "VT_EMPTY (0x0000)");
            assert_eq!(vartype(&null.value), "VT_NULL (0x0001)");
        }

        #[test]
        fn fixture_redaction_removes_local_bstr_path() {
            let mut record = json!({
                "frame": {"arguments": [{"value": {"kind": "bstr", "value": "C:\\\\Temp\\\\fixture.xlsx"}}]}
            });
            redact_fixture_path(&mut record);
            assert_eq!(
                record["frame"]["arguments"][0]["value"]["value"],
                "<temporary fixture path redacted>"
            );
        }

        #[test]
        fn zero_argument_exception_does_not_map_argerr_to_a_source_parameter() {
            let normalized = parity_pu_arg_err_json(0, 0x8002_0009_u32 as i32, 0);
            assert_eq!(normalized["applicable"], false);
            assert!(normalized["source_parameter_index"].is_null());
            assert_eq!(
                normalized["zero_argument_exception_interpretation"],
                "raw sentinel is not a source parameter"
            );
        }

        #[test]
        fn cleanup_record_never_claims_forced_termination() {
            let cleanup = cleanup_record(true, true, true);
            assert_eq!(cleanup["workbook_closed"], true);
            assert_eq!(cleanup["excel_quit_requested"], true);
            assert_eq!(cleanup["process_exited"], true);
            assert_eq!(cleanup["forced_termination"], false);
        }
    }
}

fn merge_capture(mut existing: Capture, fresh: Capture) -> Capture {
    existing.manifest = fresh.manifest;
    merge_records(&mut existing.environments, fresh.environments);
    merge_records(&mut existing.observations, fresh.observations);
    merge_records(&mut existing.cases, fresh.cases);
    merge_records(&mut existing.unresolved, fresh.unresolved);
    existing
}

fn merge_records(existing: &mut Vec<Value>, fresh: Vec<Value>) {
    let mut records = BTreeMap::new();
    for record in std::mem::take(existing).into_iter().chain(fresh) {
        let key = value_string(&record, "id");
        records.insert(key, record);
    }
    *existing = records.into_values().collect();
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn portable_manifest_contains_no_machine_path() {
        let manifest = runtime_manifest("csharp-dynamic-control");
        assert!(!manifest.contains(":\\"));
        assert!(manifest.ends_with('\n'));
    }

    #[test]
    fn empty_jsonl_keeps_final_lf() {
        assert_eq!(jsonl(&[]).expect("valid empty JSONL"), "\n");
    }

    #[test]
    fn reports_are_stable_for_empty_capture() {
        let capture = Capture {
            manifest: runtime_manifest("not-run"),
            environments: Vec::new(),
            observations: Vec::new(),
            cases: Vec::new(),
            unresolved: Vec::new(),
        };
        let first = artifacts(&capture).expect("artifacts");
        let second = artifacts(&capture).expect("artifacts");
        assert_eq!(first, second);
        assert!(first.values().all(|text| text.ends_with('\n')));
    }

    #[test]
    fn generated_runtime_artifacts_reject_known_mojibake_patterns() {
        assert!(reject_mojibake("Application â†’ Workbooks", Path::new("report.md")).is_err());
        let capture = Capture {
            manifest: runtime_manifest("not-run"),
            environments: Vec::new(),
            observations: Vec::new(),
            cases: Vec::new(),
            unresolved: Vec::new(),
        };
        for (path, text) in artifacts(&capture).expect("artifacts") {
            reject_mojibake(&text, &path).expect("generated artifact contains no mojibake");
        }
    }

    #[test]
    fn parity_mode_configurations_are_serialized_and_source_bounded() {
        let baseline = parity_configuration(ParityMode::RustBaseline);
        let pywin_dynamic = parity_configuration(ParityMode::Pywin32Dynamic);
        let comtypes_generated = parity_configuration(ParityMode::ComtypesGenerated);
        assert_eq!(baseline.invoke_lcid, 0x0400);
        assert_eq!(pywin_dynamic.activation_api, "CoCreateInstanceEx");
        assert_eq!(pywin_dynamic.invoke_lcid, 0);
        assert_eq!(comtypes_generated.get_ids_of_names_lcid, 0);
        assert!(comtypes_generated
            .dual_interface_handling
            .contains("no vtable layout is improvised"));
        let json = serde_json::to_value(pywin_dynamic).expect("serializable configuration");
        assert_eq!(json["mode"], "pywin32-dynamic");
        assert_eq!(ParityMode::parse("comtypes-dynamic").expect("mode").id(), "comtypes-dynamic");
        assert!(ParityMode::parse("invented-mode").is_err());
    }
}
