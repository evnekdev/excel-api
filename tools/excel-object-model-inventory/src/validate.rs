use crate::{markdown, model};
use serde_json::Value;
use std::collections::BTreeSet;
use std::fs;
use std::path::Path;

pub fn check(root: &Path) -> Result<(), String> {
    let metadata = root.join("metadata/excel-object-model");
    let objects = records(&metadata.join("objects"))?;
    let enums = records(&metadata.join("enums"))?;
    let mut ids = BTreeSet::new();
    let mut members = BTreeSet::new();
    for object in &objects {
        validate_schema_version(object, "object")?;
        validate_status(object)?;
        let surface_class = object["surface_class"]
            .as_str()
            .ok_or("object lacks surface class")?;
        if !model::SURFACE_CLASSES.contains(&surface_class) {
            return Err(format!("invalid surface class {surface_class}"));
        }
        let id = object["id"].as_str().ok_or("object lacks id")?;
        if !ids.insert(id.to_owned()) {
            return Err(format!("duplicate object id {id}"));
        }
        for member in object["members"]
            .as_array()
            .ok_or("object members must be an array")?
        {
            validate_schema_version(member, "member")?;
            validate_status(member)?;
            let id = member["id"].as_str().ok_or("member lacks id")?;
            if !members.insert(id.to_owned()) {
                return Err(format!("duplicate member id {id}"));
            }
            if member["dispid"].as_i64().is_none() {
                return Err(format!("member {id} has invalid DISPID"));
            }
            let code_mapping = model::implemented_member_ids().contains(id);
            let metadata_implemented =
                member["implementation_status"].as_str() == Some("implemented");
            if code_mapping != metadata_implemented {
                return Err(format!(
                    "implementation mapping disagrees with metadata for {id}"
                ));
            }
        }
    }
    for enumeration in &enums {
        validate_schema_version(enumeration, "enum")?;
        let id = enumeration["id"].as_str().ok_or("enum lacks id")?;
        if !ids.insert(id.to_owned()) {
            return Err(format!("duplicate object or enum id {id}"));
        }
    }
    for id in model::implemented_member_ids() {
        if !members.contains(id) {
            return Err(format!(
                "implemented crate member {id} is absent from metadata"
            ));
        }
    }
    let mut object_ids = BTreeSet::new();
    for object in &objects {
        object_ids.insert(object["id"].as_str().ok_or("object lacks id")?.to_owned());
    }
    let relationships_path = metadata.join("relationships.json");
    let relationships_text =
        fs::read_to_string(&relationships_path).map_err(|error| error.to_string())?;
    validate_text(&relationships_text, &relationships_path)?;
    let relationships: Vec<Value> =
        serde_json::from_str(&relationships_text).map_err(|error| error.to_string())?;
    for relationship in relationships {
        validate_schema_version(&relationship, "relationship")?;
        let source = relationship["source"]
            .as_str()
            .ok_or("relationship source missing")?;
        if !object_ids.contains(source) {
            return Err(format!("relationship source {source} does not exist"));
        }
        let target = relationship["target"]
            .as_str()
            .ok_or("relationship target missing")?;
        if !object_ids.contains(target) {
            return Err(format!("relationship target {target} does not exist"));
        }
    }
    for name in ["aliases.json", "manifest.toml", "overrides.toml"] {
        let path = metadata.join(name);
        let text = fs::read_to_string(&path).map_err(|error| error.to_string())?;
        validate_text(&text, &path)?;
    }
    let baseline_path = metadata.join("baseline.json");
    let baseline_text = fs::read_to_string(&baseline_path).map_err(|error| error.to_string())?;
    validate_text(&baseline_text, &baseline_path)?;
    let baseline: Value =
        serde_json::from_str(&baseline_text).map_err(|error| error.to_string())?;
    validate_schema_version(&baseline, "baseline")?;
    let member_count: usize = objects
        .iter()
        .map(|object| object["members"].as_array().map_or(0, Vec::len))
        .sum();
    if baseline["objects"].as_u64() != Some(objects.len() as u64)
        || baseline["members"].as_u64() != Some(member_count as u64)
        || baseline["enums"].as_u64() != Some(enums.len() as u64)
    {
        return Err("baseline counts disagree with committed metadata".to_owned());
    }
    let expected = markdown::planned_outputs(root)?;
    for (path, generated) in expected {
        let actual = fs::read_to_string(&path)
            .map_err(|error| format!("missing generated {}: {error}", path.display()))?;
        if actual != generated {
            return Err(format!("generated Markdown is stale: {}", path.display()));
        }
        validate_generated_markers(&actual, &path)?;
        validate_text(&actual, &path)?;
    }
    Ok(())
}
fn validate_schema_version(value: &Value, kind: &str) -> Result<(), String> {
    if value["schema_version"].as_u64() != Some(u64::from(model::SCHEMA_VERSION)) {
        return Err(format!("{kind} has an unsupported schema version"));
    }
    Ok(())
}
fn validate_status(value: &Value) -> Result<(), String> {
    for (field, allowed) in [
        ("implementation_status", model::IMPLEMENTATION_STATUSES),
        ("documentation_status", model::DOCUMENTATION_STATUSES),
        ("test_status", model::TEST_STATUSES),
    ] {
        if let Some(status) = value.get(field).and_then(Value::as_str)
            && !allowed.contains(&status)
        {
            return Err(format!("invalid {field} {status}"));
        }
    }
    if let Some(status) = value.get("source_confidence").and_then(Value::as_str)
        && !model::CONFIDENCE_STATUSES.contains(&status)
    {
        return Err(format!("invalid source confidence {status}"));
    }
    Ok(())
}
fn records(directory: &Path) -> Result<Vec<Value>, String> {
    let mut values = Vec::new();
    for entry in fs::read_dir(directory).map_err(|error| error.to_string())? {
        let path = entry.map_err(|error| error.to_string())?.path();
        if path.extension().and_then(|value| value.to_str()) == Some("json") {
            let text = fs::read_to_string(&path).map_err(|error| error.to_string())?;
            validate_text(&text, &path)?;
            values.push(serde_json::from_str(&text).map_err(|error| error.to_string())?);
        }
    }
    Ok(values)
}
fn validate_text(text: &str, path: &Path) -> Result<(), String> {
    if text.contains('\r') || !text.ends_with('\n') {
        return Err(format!(
            "{} must have LF endings and a final newline",
            path.display()
        ));
    }
    if text.contains("C:\\\\")
        || text.contains("\\\\?\\")
        || text.contains("\"pid\"")
        || text.contains("ptr=")
    {
        return Err(format!(
            "{} contains prohibited machine-specific data",
            path.display()
        ));
    }
    Ok(())
}
fn validate_generated_markers(text: &str, path: &Path) -> Result<(), String> {
    const BEGIN: &str = "<!-- BEGIN GENERATED MEMBERS -->";
    const END: &str = "<!-- END GENERATED MEMBERS -->";
    let begins = text.matches(BEGIN).count();
    let ends = text.matches(END).count();
    if begins != ends || begins > 1 {
        return Err(format!(
            "{} has unbalanced generated markers",
            path.display()
        ));
    }
    Ok(())
}
