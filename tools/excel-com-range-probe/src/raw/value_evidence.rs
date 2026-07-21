//! Deterministic, pointer-free reports for the Prompt 05H value-runtime evidence.

use std::collections::BTreeSet;
use std::fs;
use std::path::Path;

use serde_json::Value;

use super::read_jsonl;

const EVIDENCE_FILES: &[&str] = &[
    "case-definitions.jsonl",
    "scalar-value2-observations.jsonl",
    "scalar-value-observations.jsonl",
    "blank-null-empty-observations.jsonl",
    "date-currency-observations.jsonl",
    "error-observations.jsonl",
    "formula-observations.jsonl",
    "rectangular-read-observations.jsonl",
    "rectangular-write-observations.jsonl",
    "mixed-array-observations.jsonl",
    "precision-observations.jsonl",
    "string-observations.jsonl",
    "dynamic-array-observations.jsonl",
    "stability-observations.jsonl",
    "unresolved.jsonl",
];

pub(super) fn refresh(root: &Path) -> Result<(), String> {
    fs::create_dir_all(root.join("../generated/value-safearray-runtime"))
        .map_err(|error| error.to_string())?;
    for (name, body) in reports(root)? {
        fs::write(root.join("../generated/value-safearray-runtime").join(name), body)
            .map_err(|error| error.to_string())?;
    }
    Ok(())
}

pub(super) fn check(root: &Path) -> Result<(), String> {
    for name in EVIDENCE_FILES {
        let path = root.join(name);
        let text = fs::read_to_string(&path)
            .map_err(|error| format!("missing value-runtime evidence {}: {error}", path.display()))?;
        if text.contains("\r\n") || !text.ends_with('\n') {
            return Err(format!("{} must use LF endings and a final newline", path.display()));
        }
        if text.contains("ptr=") || text.contains("\\\\?\\") || text.contains("\"pid\"") {
            return Err(format!("{} contains prohibited machine-specific data", path.display()));
        }
        let _ = read_jsonl(&path)?;
    }
    let scalar_value2 = read(root, "scalar-value2-observations.jsonl")?;
    let scalar_value = read(root, "scalar-value-observations.jsonl")?;
    let rectangular_read = read(root, "rectangular-read-observations.jsonl")?;
    let rectangular_write = read(root, "rectangular-write-observations.jsonl")?;
    require_count("scalar Value2", &scalar_value2, 28)?;
    require_count("scalar Value", &scalar_value, 28)?;
    require_count("rectangular reads", &rectangular_read, 13)?;
    require_count("rectangular writes", &rectangular_write, 9)?;
    for (name, rows, minimum) in [
        ("blank/null/empty", read(root, "blank-null-empty-observations.jsonl")?, 6),
        ("date/currency", read(root, "date-currency-observations.jsonl")?, 11),
        ("errors", read(root, "error-observations.jsonl")?, 7),
        ("formulas", read(root, "formula-observations.jsonl")?, 8),
        ("mixed array", read(root, "mixed-array-observations.jsonl")?, 1),
        ("precision", read(root, "precision-observations.jsonl")?, 10),
        ("strings", read(root, "string-observations.jsonl")?, 14),
        ("dynamic arrays", read(root, "dynamic-array-observations.jsonl")?, 1),
        ("stability", read(root, "stability-observations.jsonl")?, 40),
    ] {
        require_count(name, &rows, minimum)?;
    }
    for (name, expected) in reports(root)? {
        let path = root.join("../generated/value-safearray-runtime").join(name);
        let actual = fs::read_to_string(&path)
            .map_err(|error| format!("missing generated value-runtime report {}: {error}", path.display()))?;
        if actual != expected {
            return Err(format!("value-runtime report {} is stale; rerun value-matrix-refresh", path.display()));
        }
    }
    Ok(())
}

fn require_count(label: &str, rows: &[Value], minimum: usize) -> Result<(), String> {
    if rows.len() < minimum {
        Err(format!("{label} evidence is incomplete: {}/{} rows", rows.len(), minimum))
    } else {
        Ok(())
    }
}

fn read(root: &Path, name: &str) -> Result<Vec<Value>, String> {
    read_jsonl(&root.join(name))
}

fn reports(root: &Path) -> Result<Vec<(&'static str, String)>, String> {
    let cases = read(root, "case-definitions.jsonl")?;
    let scalar_value2 = read(root, "scalar-value2-observations.jsonl")?;
    let scalar_value = read(root, "scalar-value-observations.jsonl")?;
    let blank = read(root, "blank-null-empty-observations.jsonl")?;
    let dates = read(root, "date-currency-observations.jsonl")?;
    let errors = read(root, "error-observations.jsonl")?;
    let formulas = read(root, "formula-observations.jsonl")?;
    let rectangles_read = read(root, "rectangular-read-observations.jsonl")?;
    let rectangles_write = read(root, "rectangular-write-observations.jsonl")?;
    let mixed = read(root, "mixed-array-observations.jsonl")?;
    let precision = read(root, "precision-observations.jsonl")?;
    let strings = read(root, "string-observations.jsonl")?;
    let dynamic = read(root, "dynamic-array-observations.jsonl")?;
    let stability = read(root, "stability-observations.jsonl")?;
    let unresolved = read(root, "unresolved.jsonl")?;

    Ok(vec![
        ("scalar-value2-matrix.md", report("Scalar `Value2` matrix", &scalar_value2)),
        ("scalar-value-matrix.md", report("Scalar `Value` matrix", &scalar_value)),
        ("value-vs-value2.md", comparison_report(&scalar_value, &scalar_value2)),
        ("blank-null-empty.md", report("Blank, Empty, Null, and cleared cells", &blank)),
        ("date-currency.md", report("Date and currency", &dates)),
        ("excel-errors.md", report("Excel error values", &errors)),
        ("formulas.md", report("Scalar formulas", &formulas)),
        ("formula2-dynamic-arrays.md", report("Formula2 and dynamic arrays", &dynamic)),
        ("safearray-layout.md", safearray_report(&rectangles_read, &rectangles_write)),
        ("rectangular-read-matrix.md", report("Rectangular Value/Value2 reads", &rectangles_read)),
        ("rectangular-write-matrix.md", report("Rectangular SAFEARRAY writes", &rectangles_write)),
        ("mixed-array-coercions.md", report("Mixed SAFEARRAY coercions", &mixed)),
        ("numeric-precision.md", report("Numeric precision", &precision)),
        ("string-edge-cases.md", report("String edge cases", &strings)),
        ("mode-comparison.md", report("Mode comparison", &stability)),
        ("stability.md", report("Stability", &stability)),
        ("value-model-requirements.md", requirements_report(&cases, &scalar_value2, &scalar_value, &rectangles_read, &mixed, &precision, &strings)),
        ("remaining-blockers.md", report("Remaining blockers", &unresolved)),
    ])
}

fn report(title: &str, rows: &[Value]) -> String {
    let mut output = format!("# {title}\n\nObserved records: {}.\n\n| ID | Category | Primary result |\n| --- | --- | --- |\n", rows.len());
    for row in rows {
        output.push_str(&format!(
            "| {} | {} | {} |\n",
            text(row, "id"),
            text(row, "category"),
            primary_result(row)
        ));
    }
    output
}

fn comparison_report(value: &[Value], value2: &[Value]) -> String {
    format!(
        "# `Value` versus `Value2`\n\nThe canonical evidence contains {} `Value` and {} `Value2` scalar rows, keyed by their stable case IDs. Integer, floating-point, BSTR, Empty/Null, error, date, and currency outcomes are retained in the source JSONL rather than coerced in this report.\n",
        value.len(), value2.len()
    )
}

fn safearray_report(reads: &[Value], writes: &[Value]) -> String {
    let layout = reads
        .iter()
        .find(|row| row.get("id").and_then(Value::as_str) == Some("R-09"))
        .and_then(|row| row.pointer("/value2_read/layout"));
    format!(
        "# SAFEARRAY layout\n\nRectangular reads: {}. Rectangular writes: {}. The 3×4 marker case records this runtime layout:\n\n```json\n{}\n```\n\nThe marker traversal is preserved in `rectangular-read-observations.jsonl`; it establishes physical dimension 1 as Excel rows and dimension 2 as Excel columns for this environment.\n",
        reads.len(), writes.len(), compact(layout)
    )
}

fn requirements_report(
    cases: &[Value],
    scalar_value2: &[Value],
    scalar_value: &[Value],
    rectangles: &[Value],
    mixed: &[Value],
    precision: &[Value],
    strings: &[Value],
) -> String {
    let scalar_tags = collect_vartypes(scalar_value2, "/value2_read/vartype");
    let rectangular_tags = collect_vartypes(rectangles, "/value2_read/owner/vartype");
    format!(
        "# Value-model requirements\n\nThis is a derived research requirements report, not a public API design. It is based on {} case definitions, {} `Value2` scalar rows, {} `Value` scalar rows, {} rectangular reads, {} mixed-array rows, {} precision rows, and {} string rows.\n\n1. Scalar physical VARTYPEs observed: {}.\n2. Rectangular physical owner VARTYPEs observed: {}.\n3. Integer inputs must preserve their input tag in evidence; Excel normalizes accepted integer cell values to floating-point reads in the scalar data.\n4. Future policy must distinguish never-written/cleared cells, `VT_EMPTY`, `VT_NULL`, and empty BSTR input.\n5. Excel cell errors are physical `VT_ERROR` values with an exact signed `scode`; they must not be stringified in a raw value layer.\n6. Date and currency require member- and NumberFormat-aware interpretation; the evidence retains `Value` and `Value2` reads separately.\n7. Rectangular reads are observed as rank-2 `SAFEARRAY(VARIANT)` with lower bounds, upper bounds, and per-element VARTYPEs retained.\n8. Rows map to physical SAFEARRAY dimension 1 and columns to dimension 2 in the proven marker case; storage policy must remain explicit.\n9. Values at and above IEEE-754 exact-integer boundaries, non-finite values, and negative zero require explicit loss/normalization policy.\n10. BSTR UTF-16 length, embedded NUL transformation, Unicode, formula-like strings, and Excel cell-length limits require explicit policy.\n11. Unsupported or rejected writes must retain HRESULT, EXCEPINFO, `puArgErr`, and post-failure read-back.\n12. No finalized Rust enum or public API is defined here.\n",
        cases.len(), scalar_value2.len(), scalar_value.len(), rectangles.len(), mixed.len(), precision.len(), strings.len(), set_text(&scalar_tags), set_text(&rectangular_tags)
    )
}

fn collect_vartypes(rows: &[Value], pointer: &str) -> BTreeSet<u64> {
    rows.iter().filter_map(|row| row.pointer(pointer)).filter_map(Value::as_u64).collect()
}

fn set_text(set: &BTreeSet<u64>) -> String {
    if set.is_empty() { "none".to_owned() } else { set.iter().map(u64::to_string).collect::<Vec<_>>().join(", ") }
}

fn primary_result(row: &Value) -> String {
    row.pointer("/write/hresult_hex")
        .or_else(|| row.pointer("/value2_read_call/hresult_hex"))
        .or_else(|| row.get("classification"))
        .and_then(Value::as_str)
        .unwrap_or("--")
        .to_owned()
}

fn text(row: &Value, name: &str) -> String {
    row.get(name).and_then(Value::as_str).unwrap_or("--").replace('|', "\\|")
}

fn compact(value: Option<&Value>) -> String {
    value.map(Value::to_string).unwrap_or_else(|| "null".to_owned())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::raw::jsonl;

    #[test]
    fn report_is_deterministic() {
        let rows = vec![serde_json::json!({"id":"B","category":"x"}), serde_json::json!({"id":"A","category":"y"})];
        assert_eq!(report("Test", &rows), report("Test", &rows));
    }

    #[test]
    fn jsonl_is_lf_terminated() {
        assert_eq!(jsonl(&[]).expect("jsonl"), "\n");
    }
}
