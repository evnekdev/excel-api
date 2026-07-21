use crate::model::*;
use serde::de::DeserializeOwned;
use std::collections::{BTreeMap, BTreeSet};
use std::fs;
use std::path::Path;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ValidationSummary {
    pub objects: usize,
    pub members: usize,
    pub enums: usize,
    pub relationships: usize,
}

pub type Records = (
    Vec<ObjectRecord>,
    Vec<MemberRecord>,
    Vec<RelationshipRecord>,
    Vec<EnumRecord>,
    SourceCoverage,
);

pub fn validate_root(root: &Path) -> Result<ValidationSummary, String> {
    validate_schema_files(root)?;
    let data = root.join("data");
    let objects: Vec<ObjectRecord> = read_jsonl(&data.join("objects.jsonl"))?;
    let members: Vec<MemberRecord> = read_jsonl(&data.join("members.jsonl"))?;
    let relationships: Vec<RelationshipRecord> = read_jsonl(&data.join("relationships.jsonl"))?;
    let enums: Vec<EnumRecord> = read_jsonl(&data.join("enums.jsonl"))?;
    let coverage: SourceCoverage = read_json(&data.join("source.json"))?;
    ensure_sorted(&objects, |record| &record.id, "objects")?;
    ensure_sorted(&members, |record| &record.id, "members")?;
    ensure_sorted(&relationships, |record| &record.id, "relationships")?;
    ensure_sorted(&enums, |record| &record.id, "enums")?;
    let mut all_ids = BTreeSet::new();
    for record in &objects {
        validate_object(record)?;
        insert_id(&mut all_ids, &record.id, &record.source.path)?;
    }
    for record in &members {
        validate_member(record)?;
        insert_id(&mut all_ids, &record.id, &record.source.path)?;
    }
    for record in &enums {
        validate_enum(record)?;
        insert_id(&mut all_ids, &record.id, &record.source.path)?;
    }
    let object_ids: BTreeSet<&str> = objects.iter().map(|record| record.id.as_str()).collect();
    for record in &members {
        if !object_ids.contains(record.owner.as_str()) {
            return Err(format!(
                "member {} at {} references unknown owner {}",
                record.id, record.source.path, record.owner
            ));
        }
    }
    let mut relationship_ids = BTreeSet::new();
    for record in &relationships {
        validate_relationship(record)?;
        if !relationship_ids.insert(record.id.clone()) {
            return Err(format!(
                "duplicate relationship ID {} at {}",
                record.id, record.source.path
            ));
        }
        if !all_ids.contains(&record.from) {
            return Err(format!(
                "relationship {} at {} has unknown from endpoint {}",
                record.id, record.source.path, record.from
            ));
        }
        if !all_ids.contains(&record.to) {
            return Err(format!(
                "relationship {} at {} has unknown to endpoint {}",
                record.id, record.source.path, record.to
            ));
        }
        let Some(member) = &record.member else {
            continue;
        };
        if !all_ids.contains(member) {
            return Err(format!(
                "relationship {} at {} has unknown member endpoint {}",
                record.id, record.source.path, member
            ));
        }
    }
    validate_coverage(&coverage)?;
    scan_data_for_absolute_paths(&data)?;
    Ok(ValidationSummary {
        objects: objects.len(),
        members: members.len(),
        enums: enums.len(),
        relationships: relationships.len(),
    })
}

pub fn load_records(root: &Path) -> Result<Records, String> {
    let data = root.join("data");
    Ok((
        read_jsonl(&data.join("objects.jsonl"))?,
        read_jsonl(&data.join("members.jsonl"))?,
        read_jsonl(&data.join("relationships.jsonl"))?,
        read_jsonl(&data.join("enums.jsonl"))?,
        read_json(&data.join("source.json"))?,
    ))
}

fn validate_schema_files(root: &Path) -> Result<(), String> {
    for name in [
        "object.schema.json",
        "member.schema.json",
        "relationship.schema.json",
        "enum.schema.json",
        "source.schema.json",
    ] {
        let path = root.join("schema").join(name);
        let value: serde_json::Value = read_json(&path)?;
        if value
            .get("$schema")
            .and_then(serde_json::Value::as_str)
            .is_none()
            || value.get("type").and_then(serde_json::Value::as_str) != Some("object")
        {
            return Err(format!(
                "schema {} must declare a JSON Schema URI and object type",
                path.display()
            ));
        }
    }
    Ok(())
}

fn validate_object(record: &ObjectRecord) -> Result<(), String> {
    validate_schema_version(record.schema_version, &record.id, &record.source.path)?;
    validate_base(
        &record.id,
        &record.kind,
        &record.summary,
        &record.source,
        &record.verification,
        &record.implementation,
        &record.provenance,
    )?;
    if record.kind != "object" || record.name.is_empty() || record.namespace != "Excel" {
        return Err(format!(
            "object {} at {} has invalid object fields",
            record.id, record.source.path
        ));
    }
    Ok(())
}

fn validate_member(record: &MemberRecord) -> Result<(), String> {
    validate_schema_version(record.schema_version, &record.id, &record.source.path)?;
    validate_base(
        &record.id,
        &record.kind,
        &record.summary,
        &record.source,
        &record.verification,
        &record.implementation,
        &record.provenance,
    )?;
    if !matches!(
        record.kind.as_str(),
        "property" | "method" | "event" | "unresolved"
    ) {
        return Err(format!(
            "member {} at {} has unknown kind {}",
            record.id, record.source.path, record.kind
        ));
    }
    if record.owner.is_empty()
        || record.name.is_empty()
        || !matches!(
            record.access.as_str(),
            "read-only" | "read-write" | "write-only" | "unknown" | "not-applicable"
        )
    {
        return Err(format!(
            "member {} at {} has invalid owner or access",
            record.id, record.source.path
        ));
    }
    for parameter in &record.parameters {
        if parameter.name.is_empty()
            || !matches!(
                parameter.optionality.as_str(),
                "required" | "optional" | "unknown"
            )
        {
            return Err(format!(
                "member {} at {} has malformed parameter {}",
                record.id, record.source.path, parameter.name
            ));
        }
    }
    if record.dispatch.dispid.is_some() || !record.dispatch.invoke_kinds.is_empty() {
        return Err(format!(
            "member {} at {} contains unverified type-library dispatch data",
            record.id, record.source.path
        ));
    }
    Ok(())
}

fn validate_enum(record: &EnumRecord) -> Result<(), String> {
    validate_schema_version(record.schema_version, &record.id, &record.source.path)?;
    validate_base(
        &record.id,
        &record.kind,
        &record.summary,
        &record.source,
        &record.verification,
        &record.implementation,
        &record.provenance,
    )?;
    if record.kind != "enum"
        || record.name.is_empty()
        || record.namespace != "Excel"
        || record.members.iter().any(|member| member.name.is_empty())
    {
        return Err(format!(
            "enum {} at {} has invalid enum fields",
            record.id, record.source.path
        ));
    }
    Ok(())
}

fn validate_relationship(record: &RelationshipRecord) -> Result<(), String> {
    validate_schema_version(record.schema_version, &record.id, &record.source.path)?;
    validate_base(
        &record.id,
        &record.kind,
        "relationship",
        &record.source,
        &record.verification,
        &record.implementation,
        &record.provenance,
    )?;
    if !matches!(
        record.kind.as_str(),
        "object-has-member"
            | "returns"
            | "collection-contains-object"
            | "object-belongs-to-collection"
            | "event-belongs-to-object"
            | "member-references-enum"
            | "related-link"
    ) || record.from.is_empty()
        || record.to.is_empty()
    {
        return Err(format!(
            "relationship {} at {} has malformed kind or endpoints",
            record.id, record.source.path
        ));
    }
    Ok(())
}

fn validate_coverage(coverage: &SourceCoverage) -> Result<(), String> {
    if coverage.schema_version != SCHEMA_VERSION
        || coverage.source.commit.len() != 40
        || !coverage
            .source
            .commit
            .bytes()
            .all(|byte| byte.is_ascii_hexdigit())
    {
        return Err(
            "source coverage has an unsupported schema version or missing source commit".to_owned(),
        );
    }
    if coverage.parsed_files + coverage.skipped_files.len() != coverage.selected_files {
        return Err("source coverage file counts do not reconcile".to_owned());
    }
    Ok(())
}

fn validate_base(
    id: &str,
    kind: &str,
    summary: &str,
    source: &SourceRef,
    verification: &Verification,
    implementation: &Implementation,
    provenance: &BTreeMap<String, String>,
) -> Result<(), String> {
    if id.is_empty() || !id.starts_with("Excel.") || kind.is_empty() || summary.is_empty() {
        return Err(format!(
            "record {id} at {} is missing an identifier, kind, or summary",
            source.path
        ));
    }
    if summary.chars().count() > MAX_SUMMARY_CHARS {
        return Err(format!(
            "record {id} at {} exceeds the summary length policy",
            source.path
        ));
    }
    if source.repository.is_empty()
        || source.commit.len() != 40
        || !source.commit.bytes().all(|byte| byte.is_ascii_hexdigit())
        || source.path.is_empty()
        || source.extraction_method.is_empty()
        || source.path.contains('\\')
        || source.path.starts_with('/')
        || source.path.contains(":/")
    {
        return Err(format!(
            "record {id} has missing or non-portable source provenance at {}",
            source.path
        ));
    }
    if !verification.documentation || verification.typelib || verification.runtime {
        return Err(format!(
            "record {id} at {} has an impossible initial verification combination",
            source.path
        ));
    }
    if !matches!(
        implementation.status.as_str(),
        "unplanned" | "planned" | "implemented" | "tested" | "stable" | "deferred" | "unsupported"
    ) {
        return Err(format!(
            "record {id} at {} has invalid implementation status {}",
            source.path, implementation.status
        ));
    }
    for required in ["source", "verification", "implementation.status"] {
        if !provenance.contains_key(required) {
            return Err(format!(
                "record {id} at {} is missing provenance for {required}",
                source.path
            ));
        }
    }
    Ok(())
}

fn validate_schema_version(schema_version: u32, id: &str, path: &str) -> Result<(), String> {
    if schema_version != SCHEMA_VERSION {
        return Err(format!(
            "record {id} at {path} has unsupported schema version {schema_version}"
        ));
    }
    Ok(())
}

fn insert_id(ids: &mut BTreeSet<String>, id: &str, path: &str) -> Result<(), String> {
    if !ids.insert(id.to_owned()) {
        return Err(format!("duplicate canonical ID {id} at {path}"));
    }
    Ok(())
}

fn ensure_sorted<T, F>(records: &[T], key: F, name: &str) -> Result<(), String>
where
    F: Fn(&T) -> &String,
{
    let mut previous: Option<&String> = None;
    for record in records {
        let current = key(record);
        if previous.is_some_and(|last| last >= current) {
            return Err(format!(
                "{name}.jsonl is not in strict stable ID order at {current}"
            ));
        }
        previous = Some(current);
    }
    Ok(())
}

fn read_jsonl<T: DeserializeOwned>(path: &Path) -> Result<Vec<T>, String> {
    let input = fs::read_to_string(path)
        .map_err(|error| format!("cannot read {}: {error}", path.display()))?;
    if !input.is_empty() && !input.ends_with('\n') {
        return Err(format!("{} must end with a newline", path.display()));
    }
    if input.contains("\r\n") {
        return Err(format!("{} must use LF line endings", path.display()));
    }
    input
        .lines()
        .enumerate()
        .map(|(line, value)| {
            serde_json::from_str(value).map_err(|error| {
                format!("invalid JSON at {}:{}: {error}", path.display(), line + 1)
            })
        })
        .collect()
}

fn read_json<T: DeserializeOwned>(path: &Path) -> Result<T, String> {
    let input = fs::read_to_string(path)
        .map_err(|error| format!("cannot read {}: {error}", path.display()))?;
    if !input.ends_with('\n') {
        return Err(format!("{} must end with a newline", path.display()));
    }
    serde_json::from_str(&input)
        .map_err(|error| format!("invalid JSON in {}: {error}", path.display()))
}

fn scan_data_for_absolute_paths(data: &Path) -> Result<(), String> {
    for name in [
        "objects.jsonl",
        "members.jsonl",
        "relationships.jsonl",
        "enums.jsonl",
        "source.json",
    ] {
        let path = data.join(name);
        let input = fs::read_to_string(&path)
            .map_err(|error| format!("cannot read {}: {error}", path.display()))?;
        if input.contains("C:\\\\")
            || input.contains("/Users/")
            || input.contains("/home/")
            || input.contains("\\\\?\\")
        {
            return Err(format!(
                "{} contains an absolute user or machine path",
                path.display()
            ));
        }
    }
    Ok(())
}
