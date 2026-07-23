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
        let roadmap_class = object["roadmap_class"]
            .as_str()
            .ok_or("object lacks roadmap class")?;
        if !model::ROADMAP_CLASSES.contains(&roadmap_class) {
            return Err(format!("invalid roadmap class {roadmap_class}"));
        }
        let id = object["id"].as_str().ok_or("object lacks id")?;
        if !ids.insert(id.to_owned()) {
            return Err(format!("duplicate object id {id}"));
        }
        if let Some(collection) = object["collection"].as_object() {
            let element_type = collection
                .get("element_type")
                .and_then(Value::as_str)
                .ok_or("collection lacks element_type")?;
            if element_type.is_empty() {
                return Err("collection has empty element_type".to_owned());
            }
            for field in ["count_member_id", "item_member_id", "enumerator_member_id"] {
                if let Some(value) = collection.get(field).and_then(Value::as_str)
                    && !object["members"]
                        .as_array()
                        .into_iter()
                        .flatten()
                        .any(|member| member["id"].as_str() == Some(value))
                {
                    return Err(format!(
                        "collection {field} does not reference an object member"
                    ));
                }
            }
            for kind in collection["index_kinds"]
                .as_array()
                .ok_or("collection lacks index_kinds")?
            {
                let kind = kind.as_str().ok_or("collection index kind is not text")?;
                if !model::COLLECTION_INDEX_KINDS.contains(&kind) {
                    return Err(format!("invalid collection index kind {kind}"));
                }
            }
            let status = collection["iterator_status"]
                .as_str()
                .ok_or("collection lacks iterator_status")?;
            if !model::COLLECTION_ITERATOR_STATUSES.contains(&status) {
                return Err(format!("invalid collection iterator status {status}"));
            }
            let requires_policy = matches!(
                object["name"].as_str(),
                Some(
                    "FormatConditions"
                        | "ColorScaleCriteria"
                        | "IconCriteria"
                        | "Styles"
                        | "Comments"
                        | "CommentsThreaded"
                        | "Hyperlinks"
                )
            );
            match collection
                .get("heterogeneous_policy")
                .and_then(Value::as_str)
            {
                Some(policy)
                    if ["homogeneous", "typed-subtype-by-type-property"].contains(&policy) => {}
                Some(policy) => {
                    return Err(format!("invalid collection heterogeneous policy {policy}"));
                }
                None if requires_policy => {
                    return Err("collection lacks heterogeneous_policy".to_owned());
                }
                None => {}
            }
        } else if !object["collection"].is_null() {
            return Err("collection must be an object or null".to_owned());
        }
        if let Some(capabilities) = object["reference_capabilities"].as_object() {
            for field in ["input_styles", "output_styles"] {
                for style in capabilities[field]
                    .as_array()
                    .ok_or("reference capabilities styles must be arrays")?
                {
                    let style = style.as_str().ok_or("reference style is not text")?;
                    if !model::REFERENCE_STYLES.contains(&style) {
                        return Err(format!("invalid reference style {style}"));
                    }
                }
            }
            for field in ["relative_address", "external_address", "formula_conversion"] {
                if !capabilities[field].is_boolean() {
                    return Err(format!("reference capability {field} must be boolean"));
                }
            }
        } else if !object["reference_capabilities"].is_null() {
            return Err("reference_capabilities must be an object or null".to_owned());
        }
        if let Some(categories) = object["evaluation_result_categories"].as_array() {
            for category in categories {
                let category = category.as_str().ok_or("evaluation category is not text")?;
                if !model::EVALUATION_RESULT_CATEGORIES.contains(&category) {
                    return Err(format!("invalid evaluation result category {category}"));
                }
            }
        } else if !object["evaluation_result_categories"].is_null() {
            return Err("evaluation_result_categories must be an array or null".to_owned());
        }
        if let Some(capabilities) = object["formatting_capability"].as_object() {
            for field in [
                "font",
                "fill",
                "borders",
                "number_format",
                "alignment",
                "dimensions",
                "autofit",
            ] {
                if !capabilities[field].is_boolean() {
                    return Err(format!("formatting capability {field} must be boolean"));
                }
            }
        } else if !object["formatting_capability"].is_null() {
            return Err("formatting_capability must be an object or null".to_owned());
        }
        validate_boolean_capability(
            object,
            "formula_capability",
            &[
                "a1",
                "r1c1",
                "formula2",
                "formula2_r1c1",
                "dynamic_array",
                "legacy_array",
                "locale_formula",
                "mixed_values",
            ],
        )?;
        validate_boolean_capability(
            object,
            "calculation_capability",
            &[
                "mode",
                "state",
                "version",
                "before_save",
                "calculate",
                "full",
                "full_rebuild",
                "mark_dirty",
            ],
        )?;
        validate_boolean_capability(
            object,
            "auditing_search_capability",
            &[
                "precedents",
                "dependents",
                "special_cells",
                "find",
                "replace",
                "wrap_safe_iterator",
            ],
        )?;
        validate_boolean_capability(
            object,
            "structured_data_capability",
            &[
                "tables",
                "sort",
                "filter",
                "validation",
                "remove_duplicates",
                "structural_editing",
            ],
        )?;
        validate_boolean_capability(
            object,
            "advanced_presentation_capability",
            &[
                "conditional_formatting",
                "color_scales",
                "data_bars",
                "icon_sets",
                "styles",
                "theme_colors",
                "display_format",
                "legacy_comments",
                "threaded_comment_inspection",
                "hyperlinks",
            ],
        )?;
        validate_boolean_capability(
            object,
            "drawing_capability",
            &[
                "embedded_charts",
                "chart_sheets",
                "series",
                "axes",
                "data_labels",
                "trendlines",
                "error_bars",
                "shapes",
                "text_boxes",
                "pictures",
                "grouping",
                "chart_export",
                "range_image_export",
                "sparklines",
            ],
        )?;
        for member in object["members"]
            .as_array()
            .ok_or("object members must be an array")?
        {
            validate_schema_version(member, "member")?;
            validate_status(member)?;
            let origin = member["member_origin"]
                .as_str()
                .ok_or("member lacks origin")?;
            if !model::MEMBER_ORIGINS.contains(&origin) {
                return Err(format!("member has invalid origin {origin}"));
            }
            let id = member["id"].as_str().ok_or("member lacks id")?;
            if !members.insert(id.to_owned()) {
                return Err(format!("duplicate member id {id}"));
            }
            if member["dispid"].as_i64().is_none() {
                return Err(format!("member {id} has invalid DISPID"));
            }
            if !member["mixed_value_possible"].is_null()
                && !member["mixed_value_possible"].is_boolean()
            {
                return Err(format!("member {id} has invalid mixed_value_possible"));
            }
            for field in [
                "version_sensitive",
                "returns_range_or_nothing",
                "stateful_search",
                "returns_optional_dispatch",
                "one_based_field",
                "modifies_range_in_place",
                "stateful_filter",
                "clipboard_dependent",
            ] {
                if !member[field].is_null() && !member[field].is_boolean() {
                    return Err(format!("member {id} has invalid {field}"));
                }
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
fn validate_boolean_capability(
    object: &Value,
    capability_name: &str,
    allowed_fields: &[&str],
) -> Result<(), String> {
    match object[capability_name].as_object() {
        Some(capabilities) => {
            for (field, value) in capabilities {
                if !allowed_fields.contains(&field.as_str()) {
                    return Err(format!(
                        "{capability_name} has unrecognized capability {field}"
                    ));
                }
                if !value.is_boolean() {
                    return Err(format!("{capability_name} {field} must be boolean"));
                }
            }
            Ok(())
        }
        None if object[capability_name].is_null() => Ok(()),
        None => Err(format!("{capability_name} must be an object or null")),
    }
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
