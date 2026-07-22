use std::fs;
use std::path::Path;

use serde_json::Value;

use crate::typelib;

/// Extracts once into a disposable directory and compares structural records without changing committed metadata.
pub fn check_current_against_committed(root: &Path) -> Result<(), String> {
    let temporary = std::env::temp_dir().join(format!(
        "excel-object-model-inventory-diff-{}",
        std::process::id()
    ));
    if temporary.exists() {
        fs::remove_dir_all(&temporary).map_err(|error| error.to_string())?;
    }
    fs::create_dir_all(&temporary).map_err(|error| error.to_string())?;
    let extraction = typelib::extract(&temporary);
    let comparison = extraction.and_then(|_| compare(root, &temporary));
    let cleanup = fs::remove_dir_all(&temporary);
    comparison?;
    cleanup.map_err(|error| error.to_string())?;
    Ok(())
}

fn compare(committed_root: &Path, extracted_root: &Path) -> Result<(), String> {
    let committed = snapshot(&committed_root.join("metadata/excel-object-model"))?;
    let extracted = snapshot(&extracted_root.join("metadata/excel-object-model"))?;
    (committed == extracted).then_some(()).ok_or_else(|| {
        "registered typelib differs from committed structural metadata; run extract and inspect the diff"
            .to_owned()
    })
}

fn snapshot(root: &Path) -> Result<Value, String> {
    let mut objects = records(&root.join("objects"))?;
    let mut enums = records(&root.join("enums"))?;
    objects.sort_by_key(|value| value["id"].as_str().unwrap_or_default().to_owned());
    enums.sort_by_key(|value| value["id"].as_str().unwrap_or_default().to_owned());
    let relationships: Value = serde_json::from_str(
        &fs::read_to_string(root.join("relationships.json")).map_err(|error| error.to_string())?,
    )
    .map_err(|error| error.to_string())?;
    Ok(serde_json::json!({"objects": objects, "enums": enums, "relationships": relationships}))
}

fn records(directory: &Path) -> Result<Vec<Value>, String> {
    fs::read_dir(directory)
        .map_err(|error| error.to_string())?
        .filter_map(Result::ok)
        .map(|entry| entry.path())
        .filter(|path| path.extension().and_then(|value| value.to_str()) == Some("json"))
        .map(|path| {
            serde_json::from_str(&fs::read_to_string(path).map_err(|error| error.to_string())?)
                .map_err(|error| error.to_string())
        })
        .collect()
}
