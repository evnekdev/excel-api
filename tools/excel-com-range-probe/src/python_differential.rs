//! Deterministic storage for Prompt 05I Python-client differential evidence.
//!
//! The companion Python harness is the only component that drives pywin32 or
//! comtypes.  This module invokes it as an explicit opt-in command, retains
//! its client-visible values separately from raw VARIANT observations, and
//! checks/rebuilds reports without opening Excel or Python.

use std::collections::{BTreeMap, BTreeSet};
use std::fs;
use std::path::{Path, PathBuf};
use std::process::Command;

use serde_json::{json, Value};

const FILES: &[&str] = &[
    "environments.jsonl",
    "python-environments.jsonl",
    "case-definitions.jsonl",
    "mixed-array-rust-observations.jsonl",
    "mixed-array-python-observations.jsonl",
    "date-rust-observations.jsonl",
    "date-python-observations.jsonl",
    "shape-mismatch-observations.jsonl",
    "dynamic-array-observations.jsonl",
    "cross-mode-observations.jsonl",
    "source-audit.jsonl",
    "conversion-policy-comparison.jsonl",
    "unresolved.jsonl",
];

const MANIFEST: &str = "schema_version = 1\nname = \"excel-com-python-client-differential\"\nclassification = \"research-only\"\nraw_backend = \"windows-sys generic IDispatch\"\npython_clients = [\"pywin32\", \"comtypes\"]\nvalue_boundary = \"Python-visible values are never inferred to be raw returned VARTYPEs.\"\n";

const CASES: &[(&str, &str)] = &[
    ("mixed-array", "Fixed-position heterogeneous inputs, isolating Empty, Null, I4, error, date, and currency candidates."),
    ("date", "OA -1, 0, 1, fraction, modern, and negative-fraction comparisons through Value and Value2."),
    ("shape-mismatch", "Natural Python sequence rank and rectangular target-shape controls."),
    ("dynamic-array", "Formula2 SEQUENCE, text-spill, and blocked-spill controls."),
    ("stability", "Raw L/S/X repetitions are retained separately from client-wrapper conversion observations."),
];

pub fn initialize(root: &Path) -> Result<(), String> {
    fs::create_dir_all(root).map_err(|error| format!("cannot create {}: {error}", root.display()))?;
    let manifest = root.join("SOURCE_MANIFEST.toml");
    if !manifest.exists() {
        fs::write(&manifest, MANIFEST).map_err(|error| error.to_string())?;
    }
    let cases = root.join("case-definitions.jsonl");
    if !cases.exists() {
        let rows = CASES
            .iter()
            .map(|(id, detail)| {
                json!({"schema_version":1,"id":format!("case.05i.{id}"),"family":id,"detail":detail,"classification":"planned-or-observed","raw_pointer_values_recorded":false})
            })
            .collect::<Vec<_>>();
        write_jsonl(&cases, &rows)?;
    }
    for name in FILES {
        let path = root.join(name);
        if !path.exists() {
            write_jsonl(&path, &[])?;
        }
    }
    refresh(root)
}

/// Runs the checked-in Python harness once.  The Python executable and cache
/// directory are explicit so venvs and generated wrappers remain disposable
/// and outside the repository.
pub fn capture(
    root: &Path,
    python: &Path,
    client: &str,
    wrapper: &str,
    family: &str,
    environment_id: &str,
    cache_dir: &Path,
) -> Result<String, String> {
    initialize(root)?;
    if !matches!(client, "pywin32" | "comtypes") {
        return Err("python client must be pywin32 or comtypes".to_owned());
    }
    if !matches!(wrapper, "dynamic" | "generated") {
        return Err("python wrapper must be dynamic or generated".to_owned());
    }
    if !matches!(family, "all" | "mixed" | "date" | "shape" | "dynamic") {
        return Err("python family must be all, mixed, date, shape, or dynamic".to_owned());
    }
    let script = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("python")
        .join("python_differential.py");
    let output = Command::new(python)
        .arg(&script)
        .args(["--client", client, "--wrapper", wrapper, "--family", family])
        .args(["--environment-id", environment_id, "--cache-dir"])
        .arg(cache_dir)
        .output()
        .map_err(|error| format!("cannot start {}: {error}", python.display()))?;
    if !output.status.success() {
        return Err(format!(
            "Python differential failed: {}",
            String::from_utf8_lossy(&output.stderr).trim()
        ));
    }
    let result: Value = serde_json::from_slice(&output.stdout).map_err(|error| {
        format!(
            "Python differential emitted invalid JSON: {error}; stderr: {}",
            String::from_utf8_lossy(&output.stderr).trim()
        )
    })?;
    if let Some(error) = result.get("fatal_error") {
        return Err(format!("Python differential fatal error: {error}"));
    }
    persist_run(root, &result, environment_id, family)?;
    refresh(root)?;
    let count = result
        .get("records")
        .and_then(Value::as_array)
        .map_or(0, Vec::len);
    Ok(format!(
        "recorded {count} Python-client observations for {environment_id} ({client}/{wrapper}, {family})"
    ))
}

/// Persists a raw windows-sys differential envelope returned by
/// `raw::differential`.  The raw observations remain explicitly distinct from
/// Python-visible rows even when they share a case family and target range.
pub fn capture_raw(
    root: &Path,
    family: &str,
    result: &Value,
    run_id: Option<&str>,
) -> Result<String, String> {
    initialize(root)?;
    let mode = result
        .get("activation_mode")
        .and_then(Value::as_str)
        .ok_or_else(|| "raw differential omitted activation mode".to_owned())?;
    let environment_id = format!("raw-windows-sys-{mode}");
    merge_jsonl(
        root,
        "environments.jsonl",
        vec![json!({
            "schema_version":1,
            "id":format!("environment.05i.{environment_id}"),
            "environment_id":environment_id,
            "backend":"raw-windows-sys",
            "activation_mode":mode,
            "dispatch":"generic SDK IDispatch; no Excel-specific vtable",
            "classification":"Runtime-observed",
            "raw_pointer_values_recorded":false,
        })],
    )?;
    let file = match family {
        "mixed" => "mixed-array-rust-observations.jsonl",
        "date" => "date-rust-observations.jsonl",
        "shape" | "rectangles" => "shape-mismatch-observations.jsonl",
        "dynamic" => "dynamic-array-observations.jsonl",
        _ => return Err("raw differential family must be mixed, date, shape, rectangles, or dynamic".to_owned()),
    };
    let records = result
        .get("observations")
        .and_then(Value::as_array)
        .cloned()
        .ok_or_else(|| "raw differential omitted observations".to_owned())?;
    let mut rows = Vec::with_capacity(records.len());
    for mut record in records {
        let source_id = record.get("id").and_then(Value::as_str).unwrap_or("unnamed").to_owned();
        let id = run_id
            .map(|run_id| format!("raw.05i.{mode}.{run_id}.{source_id}"))
            .unwrap_or_else(|| format!("raw.05i.{mode}.{source_id}"));
        record["id"] = Value::String(id);
        record["environment_id"] = Value::String(environment_id.clone());
        record["backend"] = Value::String("raw-windows-sys".to_owned());
        if let Some(run_id) = run_id {
            record["run_id"] = Value::String(run_id.to_owned());
        }
        record["raw_pointer_values_recorded"] = Value::Bool(false);
        rows.push(record);
    }
    merge_jsonl(root, file, rows)?;
    merge_jsonl(
        root,
        "cross-mode-observations.jsonl",
        vec![json!({
            "schema_version":1,
            "id":run_id.map(|run_id| format!("run.05i.{environment_id}.{family}.{run_id}")).unwrap_or_else(|| format!("run.05i.{environment_id}.{family}")),
            "environment_id":environment_id,
            "family":family,
            "backend":"raw-windows-sys",
            "cleanup":result.get("cleanup").cloned().unwrap_or(Value::Null),
            "success":result.get("success").cloned().unwrap_or(Value::Bool(false)),
            "classification":"Runtime-observed",
            "run_id":run_id,
            "raw_pointer_values_recorded":false,
        })],
    )?;
    refresh(root)?;
    Ok(format!("recorded raw windows-sys {family} differential for {mode}"))
}

pub fn refresh(root: &Path) -> Result<(), String> {
    require_source_files(root)?;
    let generated = root
        .parent()
        .ok_or_else(|| format!("{} has no knowledge parent", root.display()))?
        .join("generated")
        .join("python-client-differential");
    fs::create_dir_all(&generated).map_err(|error| error.to_string())?;
    for (name, body) in reports(root)? {
        fs::write(generated.join(name), body).map_err(|error| error.to_string())?;
    }
    Ok(())
}

pub fn check(root: &Path) -> Result<(), String> {
    let manifest = fs::read_to_string(root.join("SOURCE_MANIFEST.toml"))
        .map_err(|error| error.to_string())?;
    if manifest != MANIFEST || manifest.contains("\r\n") || !manifest.ends_with('\n') {
        return Err("Python differential source manifest is stale or not LF-terminated".to_owned());
    }
    require_source_files(root)?;
    for name in FILES {
        let path = root.join(name);
        let text = fs::read_to_string(&path).map_err(|error| error.to_string())?;
        if text.contains("\r\n") || !text.ends_with('\n') {
            return Err(format!("{} must use LF endings and a final newline", path.display()));
        }
        if text.contains("\\\\?\\") || text.contains("\"pid\"") || text.contains("ptr=") {
            return Err(format!("{} contains machine-specific data", path.display()));
        }
        let _ = read_jsonl(&path)?;
    }
    let python_environments = read_jsonl(&root.join("python-environments.jsonl"))?;
    let mixed_raw = read_jsonl(&root.join("mixed-array-rust-observations.jsonl"))?;
    let mixed_python = read_jsonl(&root.join("mixed-array-python-observations.jsonl"))?;
    let date_raw = read_jsonl(&root.join("date-rust-observations.jsonl"))?;
    let date_python = read_jsonl(&root.join("date-python-observations.jsonl"))?;
    let shape = read_jsonl(&root.join("shape-mismatch-observations.jsonl"))?;
    let dynamic = read_jsonl(&root.join("dynamic-array-observations.jsonl"))?;
    let source_audit = read_jsonl(&root.join("source-audit.jsonl"))?;
    let conversion = read_jsonl(&root.join("conversion-policy-comparison.jsonl"))?;
    require_minimum("Python environments", &python_environments, 5)?;
    require_minimum("raw mixed-array rows", &mixed_raw, 10)?;
    require_minimum("Python mixed-array rows", &mixed_python, 35)?;
    require_minimum("raw date rows", &date_raw, 24)?;
    require_minimum("Python date rows", &date_python, 100)?;
    require_minimum("shape rows", &shape, 60)?;
    require_minimum("dynamic rows", &dynamic, 25)?;
    require_minimum("source-audit rows", &source_audit, 4)?;
    require_minimum("conversion-policy rows", &conversion, 4)?;
    let raw_modes = mixed_raw
        .iter()
        .filter_map(|row| row.get("activation_mode").and_then(Value::as_str))
        .collect::<BTreeSet<_>>();
    if raw_modes != BTreeSet::from(["L", "S", "X"]) {
        return Err(format!("raw mixed-array modes are incomplete: {raw_modes:?}"));
    }
    for mode in ["S", "X"] {
        for run_id in ["run-01", "run-02"] {
            let exists = mixed_raw.iter().any(|row| {
                row.get("activation_mode").and_then(Value::as_str) == Some(mode)
                    && row.get("run_id").and_then(Value::as_str) == Some(run_id)
            });
            if !exists {
                return Err(format!("missing {mode} mixed-array repeat {run_id}"));
            }
        }
    }
    let raw_rectangles = shape
        .iter()
        .filter(|row| {
            row.get("backend").and_then(Value::as_str) == Some("raw-windows-sys")
                && row.get("category").and_then(Value::as_str)
                    == Some("rectangular-differential")
        })
        .collect::<Vec<_>>();
    if raw_rectangles.len() < 50 {
        return Err(format!(
            "raw rectangular rows are incomplete: {}/50",
            raw_rectangles.len()
        ));
    }
    for mode in ["S", "X"] {
        for run_id in ["run-01", "run-02"] {
            let has_2x3 = raw_rectangles.iter().any(|row| {
                row.get("activation_mode").and_then(Value::as_str) == Some(mode)
                    && row.get("run_id").and_then(Value::as_str) == Some(run_id)
                    && row.get("rows").and_then(Value::as_u64) == Some(2)
                    && row.get("columns").and_then(Value::as_u64) == Some(3)
            });
            if !has_2x3 {
                return Err(format!("missing {mode} 2x3 rectangular repeat {run_id}"));
            }
        }
    }
    let generated = root
        .parent()
        .ok_or_else(|| format!("{} has no knowledge parent", root.display()))?
        .join("generated")
        .join("python-client-differential");
    for (name, expected) in reports(root)? {
        let path = generated.join(name);
        let actual = fs::read_to_string(&path).map_err(|error| error.to_string())?;
        if actual != expected {
            return Err(format!("{} is stale; rerun python-differential-refresh", path.display()));
        }
        if actual.contains("\r\n") || !actual.ends_with('\n') {
            return Err(format!("{} must use LF endings and a final newline", path.display()));
        }
    }
    Ok(())
}

fn require_minimum(label: &str, rows: &[Value], minimum: usize) -> Result<(), String> {
    if rows.len() < minimum {
        Err(format!("{label} are incomplete: {}/{} rows", rows.len(), minimum))
    } else {
        Ok(())
    }
}

fn require_source_files(root: &Path) -> Result<(), String> {
    for name in FILES {
        if !root.join(name).is_file() {
            return Err(format!("missing Python differential evidence {}", root.join(name).display()));
        }
    }
    Ok(())
}

fn persist_run(root: &Path, result: &Value, environment_id: &str, family: &str) -> Result<(), String> {
    let environment = result
        .get("environment")
        .cloned()
        .ok_or_else(|| "Python differential omitted environment".to_owned())?;
    let mut environment_row = environment;
    environment_row["id"] = Value::String(format!("environment.05i.{environment_id}"));
    environment_row["classification"] = Value::String("Runtime-observed".to_owned());
    environment_row["raw_pointer_values_recorded"] = Value::Bool(false);
    merge_jsonl(root, "environments.jsonl", vec![environment_row.clone()])?;
    environment_row["id"] = Value::String(format!("python.05i.{environment_id}"));
    merge_jsonl(root, "python-environments.jsonl", vec![environment_row])?;

    let cleanup = result.get("cleanup").cloned().unwrap_or(Value::Null);
    let success = result.get("success").cloned().unwrap_or(Value::Bool(false));
    let setup_error = result.get("setup_error").cloned().unwrap_or(Value::Null);
    merge_jsonl(
        root,
        "cross-mode-observations.jsonl",
        vec![json!({
            "schema_version":1,
            "id":format!("run.05i.{environment_id}.{family}"),
            "environment_id":environment_id,
            "family":family,
            "backend":"python-client",
            "cleanup":cleanup,
            "setup_error":setup_error,
            "success":success,
            "classification":"Runtime-observed",
            "raw_pointer_values_recorded":false,
        })],
    )?;

    let records = result.get("records").and_then(Value::as_array).cloned().unwrap_or_default();
    let mut grouped: BTreeMap<&str, Vec<Value>> = BTreeMap::new();
    for mut record in records {
        let source_id = record.get("id").and_then(Value::as_str).unwrap_or("unnamed").to_owned();
        let family_name = record
            .get("family")
            .and_then(Value::as_str)
            .unwrap_or("smoke")
            .to_owned();
        record["id"] = Value::String(format!("python.05i.{environment_id}.{source_id}"));
        record["environment_id"] = Value::String(environment_id.to_owned());
        record["client_visible_only"] = Value::Bool(true);
        record["raw_return_vartype_observable"] = Value::Bool(false);
        record["raw_pointer_values_recorded"] = Value::Bool(false);
        let file = match family_name.as_str() {
            "mixed-array" => "mixed-array-python-observations.jsonl",
            "date" => "date-python-observations.jsonl",
            "shape-mismatch" => "shape-mismatch-observations.jsonl",
            "dynamic-array" => "dynamic-array-observations.jsonl",
            _ => "cross-mode-observations.jsonl",
        };
        grouped.entry(file).or_default().push(record);
    }
    for (file, rows) in grouped {
        merge_jsonl(root, file, rows)?;
    }
    Ok(())
}

fn merge_jsonl(root: &Path, name: &str, incoming: Vec<Value>) -> Result<(), String> {
    let path = root.join(name);
    let mut rows = read_jsonl(&path)?;
    let mut merged = BTreeMap::new();
    for row in rows.drain(..).chain(incoming) {
        let id = row.get("id").and_then(Value::as_str).unwrap_or("missing").to_owned();
        merged.insert(id, row);
    }
    write_jsonl(&path, &merged.into_values().collect::<Vec<_>>())
}

fn read_jsonl(path: &Path) -> Result<Vec<Value>, String> {
    let text = fs::read_to_string(path).map_err(|error| format!("cannot read {}: {error}", path.display()))?;
    text.lines()
        .filter(|line| !line.is_empty())
        .map(|line| serde_json::from_str(line).map_err(|error| format!("{}: {error}", path.display())))
        .collect()
}

fn write_jsonl(path: &Path, rows: &[Value]) -> Result<(), String> {
    let mut rows = rows.to_vec();
    rows.sort_by_key(|row| row.get("id").and_then(Value::as_str).unwrap_or("").to_owned());
    let mut text = String::new();
    for row in rows {
        text.push_str(&serde_json::to_string(&row).map_err(|error| error.to_string())?);
        text.push('\n');
    }
    if text.is_empty() {
        text.push('\n');
    }
    fs::write(path, text).map_err(|error| format!("cannot write {}: {error}", path.display()))
}

fn reports(root: &Path) -> Result<BTreeMap<&'static str, String>, String> {
    let environments = read_jsonl(&root.join("python-environments.jsonl"))?;
    let mixed_python = read_jsonl(&root.join("mixed-array-python-observations.jsonl"))?;
    let mixed_raw = read_jsonl(&root.join("mixed-array-rust-observations.jsonl"))?;
    let date_python = read_jsonl(&root.join("date-python-observations.jsonl"))?;
    let date_raw = read_jsonl(&root.join("date-rust-observations.jsonl"))?;
    let shape = read_jsonl(&root.join("shape-mismatch-observations.jsonl"))?;
    let dynamic = read_jsonl(&root.join("dynamic-array-observations.jsonl"))?;
    let cross = read_jsonl(&root.join("cross-mode-observations.jsonl"))?;
    let source = read_jsonl(&root.join("source-audit.jsonl"))?;
    let conversion = read_jsonl(&root.join("conversion-policy-comparison.jsonl"))?;
    let unresolved = read_jsonl(&root.join("unresolved.jsonl"))?;
    let mut output = BTreeMap::new();
    output.insert("python-environments.md", table_report("Python environments", &environments));
    output.insert("mixed-array-differential.md", differential_report("Mixed SAFEARRAY differential", &mixed_raw, &mixed_python));
    output.insert("date-differential.md", differential_report("Date differential", &date_raw, &date_python));
    output.insert("shape-mismatch.md", table_report("Shape mismatch observations", &shape));
    output.insert("dynamic-arrays.md", table_report("Formula2 and dynamic-array observations", &dynamic));
    output.insert("cross-mode.md", table_report("Cross-mode and cleanup observations", &cross));
    output.insert("source-audit.md", table_report("Python client source audit", &source));
    output.insert("conversion-policies.md", table_report("Conversion-policy comparison", &conversion));
    output.insert("unresolved.md", table_report("Unresolved observations", &unresolved));
    output.insert("requirements.md", requirements_report(&mixed_raw, &mixed_python, &date_raw, &date_python, &shape, &dynamic, &cross));
    Ok(output)
}

fn table_report(title: &str, rows: &[Value]) -> String {
    let mut output = format!("# {title}\n\nObserved records: {}.\n\n| ID | Family | Result |\n| --- | --- | --- |\n", rows.len());
    for row in rows {
        output.push_str(&format!("| {} | {} | {} |\n", text(row, "id"), text(row, "family"), result_text(row)));
    }
    output
}

fn differential_report(title: &str, raw: &[Value], python: &[Value]) -> String {
    format!("# {title}\n\nRaw records: {}. Python-visible records: {}. Python client values are post-conversion observations and are not asserted to be physical return VARTYPEs.\n\n{}", raw.len(), python.len(), table_report("Python-side rows", python))
}

fn requirements_report(mixed_raw: &[Value], mixed_python: &[Value], date_raw: &[Value], date_python: &[Value], shape: &[Value], dynamic: &[Value], cross: &[Value]) -> String {
    let successful_cleanup = cross.iter().filter(|row| row.pointer("/cleanup/owned_process_exit_verified") == Some(&Value::Bool(true))).count();
    format!(
        "# Internal value-model requirements\n\nThis derived report records research requirements only; it defines no public Rust API.\n\n1. Retain raw VARTYPE, HRESULT, EXCEPINFO, `puArgErr`, SAFEARRAY rank, and bounds independently from any Python-client conversion.\n2. A client-visible `None`, numeric value, `datetime`, `Decimal`, or tuple is not evidence of a raw returned VARTYPE.\n3. Preserve the exact mixed-array position and client-side input metadata for all {} Python and {} raw rows.\n4. Preserve `Value` and `Value2`, OA serial, and NumberFormat context for all {} Python and {} raw date rows.\n5. Shape/rank rejection or coercion must retain the target range and source sequence shape ({} rows).\n6. Formula2 spill and blocked-spill results remain capability observations, not assumptions ({} rows).\n7. Each live run must begin with no Excel process and finish with an owned natural exit; {} recorded run(s) meet that condition.\n8. No public enum, conversion policy implementation, or Excel-specific vtable is introduced.\n",
        mixed_python.len(), mixed_raw.len(), date_python.len(), date_raw.len(), shape.len(), dynamic.len(), successful_cleanup
    )
}

fn text(row: &Value, name: &str) -> String {
    row.get(name).and_then(Value::as_str).unwrap_or("--").replace('|', "\\|")
}

fn result_text(row: &Value) -> String {
    row.pointer("/write/status")
        .or_else(|| row.get("success"))
        .or_else(|| row.get("classification"))
        .map(Value::to_string)
        .unwrap_or_else(|| "--".to_owned())
        .replace('|', "\\|")
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn empty_jsonl_is_lf_terminated() {
        let root = std::env::temp_dir().join("excel-com-05i-python-evidence-test");
        let _ = fs::remove_dir_all(&root);
        initialize(&root).expect("initialize");
        assert_eq!(fs::read_to_string(root.join("mixed-array-python-observations.jsonl")).expect("read"), "\n");
        let _ = fs::remove_dir_all(root);
    }

    #[test]
    fn requirements_state_the_python_variant_boundary() {
        let report = requirements_report(&[], &[], &[], &[], &[], &[], &[]);
        assert!(report.contains("not evidence of a raw returned VARTYPE"));
    }

    #[test]
    fn case_definitions_cover_required_families() {
        let families = CASES.iter().map(|(id, _)| *id).collect::<BTreeSet<_>>();
        assert!(families.contains("mixed-array"));
        assert!(families.contains("date"));
        assert!(families.contains("shape-mismatch"));
        assert!(families.contains("dynamic-array"));
    }
}
