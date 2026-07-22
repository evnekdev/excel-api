//! Deterministic validation of the Prompt 06 evidence tree.

use std::fs;
use std::path::Path;

pub(crate) fn check_evidence(root: &Path) -> Result<(), String> {
    let source = root.join("automation-value-layer");
    let generated = root.join("generated/automation-value-layer");
    for file in [
        "SOURCE_MANIFEST.toml",
        "design-decisions.jsonl",
        "conversion-policies.jsonl",
        "codec-cases.jsonl",
        "live-compatibility-observations.jsonl",
        "unresolved.jsonl",
    ] {
        let text = fs::read_to_string(source.join(file)).map_err(|error| error.to_string())?;
        if text.contains('\r') || !text.ends_with('\n') {
            return Err(format!("automation-value-layer/{file} must use LF and a final newline"));
        }
        if file.ends_with(".jsonl") {
            for line in text.lines() {
                serde_json::from_str::<serde_json::Value>(line)
                    .map_err(|error| format!("automation-value-layer/{file}: {error}"))?;
            }
        }
    }
    for file in [
        "value-model.md",
        "excel-errors.md",
        "dates.md",
        "currency.md",
        "arrays.md",
        "conversion-policies.md",
        "conversion-errors.md",
        "codec-test-matrix.md",
        "live-compatibility.md",
        "remaining-blockers.md",
    ] {
        let text = fs::read_to_string(generated.join(file)).map_err(|error| error.to_string())?;
        if text.contains('\r') || !text.ends_with('\n') || !text.starts_with("# ") {
            return Err(format!("generated/automation-value-layer/{file} is not a valid deterministic report"));
        }
    }
    Ok(())
}
