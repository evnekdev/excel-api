use crate::markdown::{self, Page};
use crate::model::*;
use crate::source;
use serde::Serialize;
use std::collections::{BTreeMap, BTreeSet};
use std::fs;
use std::path::Path;

struct ParsedPage {
    path: String,
    page: Page,
    kind: String,
    id: String,
    name: String,
    summary: String,
    summary_origin: String,
}

pub fn ingest(
    source_root: &Path,
    manifest_path: &Path,
    output: &Path,
) -> Result<IngestResult, String> {
    let manifest = source::load_manifest(manifest_path)?;
    source::verify_checkout_commit(source_root, &manifest.source.commit)?;
    let files = source::selected_files(source_root, &manifest.selection)?;
    let mut coverage = SourceCoverage {
        schema_version: SCHEMA_VERSION,
        source: manifest.source.clone(),
        selected_files: files.len(),
        parsed_files: 0,
        skipped_files: Vec::new(),
        warnings: Vec::new(),
        orphan_records: Vec::new(),
        duplicate_ids: Vec::new(),
        unresolved_owner_references: Vec::new(),
        unresolved_return_object_references: Vec::new(),
        missing_summaries: Vec::new(),
        missing_source_paths: Vec::new(),
        unresolved: Vec::new(),
    };
    let mut parsed = Vec::new();
    for (path, file) in files {
        let input =
            fs::read_to_string(&file).map_err(|error| format!("cannot read {path}: {error}"))?;
        let page = match markdown::parse_page(&input) {
            Ok(page) => page,
            Err(detail) => {
                coverage
                    .skipped_files
                    .push(issue("missing-front-matter", None, &path, &detail));
                continue;
            }
        };
        let Some(title) = page.title.as_deref() else {
            coverage
                .skipped_files
                .push(issue("missing-title", None, &path, "No title was found."));
            continue;
        };
        let Some(kind) = markdown::classify_title(title) else {
            coverage.skipped_files.push(issue(
                "unrecognized-page-kind",
                None,
                &path,
                "The title does not identify an Excel object-model entity.",
            ));
            continue;
        };
        let Some(api_name) = page.api_names.first() else {
            coverage.skipped_files.push(issue(
                "missing-api-name",
                None,
                &path,
                "No api_name front-matter value was found.",
            ));
            continue;
        };
        if page.api_names.len() > 1 {
            coverage.unresolved.push(issue("aliased-api-names", None, &path, "The page supplies more than one api_name; the first remains canonical and the aliases are retained."));
        }
        let id = canonical_id(api_name);
        if !id.starts_with("Excel.") || id.split('.').any(|segment| segment.is_empty()) {
            coverage.skipped_files.push(issue(
                "invalid-api-name",
                Some(id),
                &path,
                "The api_name cannot form a canonical Excel identifier.",
            ));
            continue;
        }
        let title_id = canonical_id(&format!("Excel.{}", title_entity_name(title, kind)));
        if !id.eq_ignore_ascii_case(&title_id) {
            coverage.skipped_files.push(issue("api-name-title-conflict", Some(id), &path, &format!("The api_name conflicts with the page title identifier {title_id}; the page was not normalized.")));
            continue;
        }
        let name = id.rsplit('.').next().unwrap_or_default().to_owned();
        let (summary, summary_origin) = markdown::concise_summary(&page, &id);
        if summary.starts_with("Documentation page for ") {
            coverage.missing_summaries.push(issue(
                "missing-summary",
                Some(id.clone()),
                &path,
                "No concise source summary was available; a project-authored placeholder was used.",
            ));
        }
        parsed.push(ParsedPage {
            path,
            page,
            kind: kind.to_owned(),
            id,
            name,
            summary,
            summary_origin,
        });
        coverage.parsed_files += 1;
    }
    parsed.sort_by(|left, right| left.id.cmp(&right.id).then(left.path.cmp(&right.path)));
    let mut grouped: BTreeMap<String, Vec<usize>> = BTreeMap::new();
    for (index, page) in parsed.iter().enumerate() {
        grouped.entry(page.id.clone()).or_default().push(index);
    }
    for (id, indexes) in grouped {
        if indexes.len() < 2 {
            continue;
        }
        let kinds: BTreeSet<String> = indexes
            .iter()
            .map(|index| parsed[*index].kind.clone())
            .collect();
        if kinds.len() == indexes.len() {
            for index in indexes {
                let page = &mut parsed[index];
                coverage.unresolved.push(issue("colliding-member-kinds", Some(page.id.clone()), &page.path, "Documentation assigns the same API name to different member kinds; the stable ID carries a deterministic kind suffix."));
                page.id = format!("{}#{}", page.id, page.kind);
            }
        } else {
            for index in indexes {
                let page = &parsed[index];
                coverage.duplicate_ids.push(issue("duplicate-stable-id", Some(id.clone()), &page.path, "More than one selected page with the same kind normalizes to this stable identifier."));
            }
        }
    }
    if !coverage.duplicate_ids.is_empty() {
        return Err(format!(
            "ingestion found {} duplicate stable ID(s); inspect source coverage before choosing a collision policy",
            coverage.duplicate_ids.len()
        ));
    }

    let mut objects = Vec::new();
    let mut enums = Vec::new();
    let mut members = Vec::new();
    for page in &parsed {
        match page.kind.as_str() {
            "object" => objects.push(object_record(page, &manifest)),
            "enum" => enums.push(enum_record(page, &manifest)),
            _ => members.push(member_record(page, &manifest, &mut coverage)),
        }
    }
    objects.sort_by(|left, right| left.id.cmp(&right.id));
    enums.sort_by(|left, right| left.id.cmp(&right.id));
    members.sort_by(|left, right| left.id.cmp(&right.id));

    let object_ids: BTreeSet<String> = objects.iter().map(|record| record.id.clone()).collect();
    let enum_ids: BTreeSet<String> = enums.iter().map(|record| record.id.clone()).collect();
    for member in &mut members {
        let Some(canonical_owner) = resolve_object_id(&member.owner, &object_ids) else {
            continue;
        };
        if canonical_owner != member.owner {
            member.provenance.insert(
                "owner".to_owned(),
                "source-frontmatter-case-normalized".to_owned(),
            );
            member.owner = canonical_owner;
        }
    }
    let mut relationships = Vec::new();
    for member in &members {
        if !object_ids.contains(&member.owner) {
            coverage.unresolved_owner_references.push(issue(
                "unknown-owner",
                Some(member.id.clone()),
                &member.source.path,
                &format!("Owner {} has no normalized object record.", member.owner),
            ));
            coverage.unresolved.push(issue(
                "unknown-owner",
                Some(member.id.clone()),
                &member.source.path,
                &format!("Owner {} has no normalized object record.", member.owner),
            ));
            continue;
        }
        let relation_kind = if member.kind == "event" {
            "event-belongs-to-object"
        } else {
            "object-has-member"
        };
        relationships.push(relationship(
            relation_kind,
            &member.owner,
            Some(&member.id),
            &member.id,
            &member.source,
            "source-frontmatter",
        ));
        if let Some(return_type) = &member.documented_return_type {
            if let Some(target) = resolve_object_name(return_type, &object_ids) {
                relationships.push(relationship(
                    "returns",
                    &member.owner,
                    Some(&member.id),
                    &target,
                    &member.source,
                    "source-return-value",
                ));
            } else if looks_like_excel_object(return_type) {
                coverage.unresolved_return_object_references.push(issue(
                    "unknown-return-object",
                    Some(member.id.clone()),
                    &member.source.path,
                    &format!(
                        "Documented return type {return_type} has no normalized object record."
                    ),
                ));
                coverage.unresolved.push(issue(
                    "unknown-return-object",
                    Some(member.id.clone()),
                    &member.source.path,
                    &format!(
                        "Documented return type {return_type} has no normalized object record."
                    ),
                ));
            }
        }
        for parameter in &member.parameters {
            let Some(documented_type) = &parameter.documented_type else {
                continue;
            };
            let Some(enum_id) = resolve_enum_name(documented_type, &enum_ids) else {
                continue;
            };
            relationships.push(relationship(
                "member-references-enum",
                &member.id,
                Some(&member.id),
                &enum_id,
                &member.source,
                "source-parameter-table",
            ));
        }
    }
    for object in &objects {
        if !object.classification.is_collection {
            continue;
        }
        if let Some(item) = &object.classification.collection_item_type {
            if let Some(target) = resolve_object_name(item, &object_ids) {
                relationships.push(relationship(
                    "collection-contains-object",
                    &object.id,
                    None,
                    &target,
                    &object.source,
                    "source-remarks",
                ));
                relationships.push(relationship(
                    "object-belongs-to-collection",
                    &target,
                    None,
                    &object.id,
                    &object.source,
                    "source-remarks",
                ));
            } else {
                coverage.unresolved.push(issue(
                    "possible-collection-item",
                    Some(object.id.clone()),
                    &object.source.path,
                    &format!("Collection item type {item} has no normalized object record."),
                ));
            }
        } else {
            coverage.unresolved.push(issue(
                "possible-collection",
                Some(object.id.clone()),
                &object.source.path,
                "The documentation identifies a collection but no item type was extracted.",
            ));
        }
    }
    relationships.sort_by(|left, right| left.id.cmp(&right.id));
    relationships.dedup_by(|left, right| left.id == right.id);
    sort_issues(&mut coverage);

    fs::create_dir_all(output)
        .map_err(|error| format!("cannot create {}: {error}", output.display()))?;
    write_jsonl(&output.join("objects.jsonl"), &objects)?;
    write_jsonl(&output.join("members.jsonl"), &members)?;
    write_jsonl(&output.join("relationships.jsonl"), &relationships)?;
    write_jsonl(&output.join("enums.jsonl"), &enums)?;
    write_json(&output.join("source.json"), &coverage)?;
    Ok(IngestResult {
        objects: objects.len(),
        members: members.len(),
        enums: enums.len(),
        relationships: relationships.len(),
        coverage,
    })
}

fn object_record(page: &ParsedPage, manifest: &Manifest) -> ObjectRecord {
    let title = page.page.title.clone().unwrap_or_default();
    let is_collection = title.to_ascii_lowercase().contains(" collection (excel)");
    let item = is_collection
        .then(|| markdown::object_collection_item(&page.page.body))
        .flatten()
        .map(|name| canonical_segment(&name));
    let mut provenance = provenance(page, "source-title");
    provenance.insert(
        "classification.is_collection".to_owned(),
        "source-title".to_owned(),
    );
    provenance.insert(
        "classification.collection_item_type".to_owned(),
        if item.is_some() {
            "source-remarks"
        } else {
            "not-extracted"
        }
        .to_owned(),
    );
    ObjectRecord {
        schema_version: SCHEMA_VERSION,
        id: page.id.clone(),
        kind: "object".to_owned(),
        name: page.name.clone(),
        namespace: "Excel".to_owned(),
        summary: page.summary.clone(),
        source: source_ref(&manifest.source, &page.path),
        documentation: Documentation {
            title,
            api_names: page.page.api_names.clone(),
        },
        classification: Classification {
            is_collection,
            collection_item_type: item,
        },
        verification: Verification::default(),
        implementation: Implementation::default(),
        examples: markdown::examples(&page.page.body),
        provenance,
    }
}

fn member_record(
    page: &ParsedPage,
    manifest: &Manifest,
    coverage: &mut SourceCoverage,
) -> MemberRecord {
    let (parameters, parameter_warnings) = markdown::parameters(&page.page.body);
    for warning in parameter_warnings {
        coverage.warnings.push(issue(
            "parameter-table",
            Some(page.id.clone()),
            &page.path,
            &warning,
        ));
    }
    if parameters
        .iter()
        .any(|parameter| parameter.optionality == "unknown")
    {
        coverage.unresolved.push(issue(
            "optionality-not-reliably-extractable",
            Some(page.id.clone()),
            &page.path,
            "At least one parameter lacks a Required/Optional classification.",
        ));
    }
    let access = if page.kind == "property" {
        match markdown::property_access(&page.page.body) {
            Some(access) => access,
            None => {
                coverage.unresolved.push(issue(
                    "ambiguous-property-access",
                    Some(page.id.clone()),
                    &page.path,
                    "The source page does not state a recognizable read/write access marker.",
                ));
                "unknown".to_owned()
            }
        }
    } else {
        "not-applicable".to_owned()
    };
    let documented_return_type = markdown::return_type(&page.page.body);
    if documented_return_type.is_none() && page.kind != "event" {
        coverage.unresolved.push(issue(
            "unknown-return-type",
            Some(page.id.clone()),
            &page.path,
            "No return-value type was extracted.",
        ));
    }
    let owner = page
        .id
        .rsplit_once('.')
        .map(|(owner, _)| owner.to_owned())
        .unwrap_or_default();
    let mut provenance = provenance(page, "source-title");
    provenance.insert("owner".to_owned(), "source-frontmatter".to_owned());
    provenance.insert("parameters".to_owned(), "source-parameter-table".to_owned());
    provenance.insert(
        "access".to_owned(),
        if access == "unknown" {
            "not-extracted"
        } else {
            "source-syntax"
        }
        .to_owned(),
    );
    provenance.insert(
        "documented_return_type".to_owned(),
        if documented_return_type.is_some() {
            "source-return-value"
        } else {
            "not-extracted"
        }
        .to_owned(),
    );
    MemberRecord {
        schema_version: SCHEMA_VERSION,
        id: page.id.clone(),
        kind: page.kind.clone(),
        owner,
        name: page.name.clone(),
        summary: page.summary.clone(),
        access,
        parameters,
        documented_return_type,
        source: source_ref(&manifest.source, &page.path),
        documentation: Documentation {
            title: page.page.title.clone().unwrap_or_default(),
            api_names: page.page.api_names.clone(),
        },
        dispatch: Dispatch {
            dispid: None,
            invoke_kinds: Vec::new(),
        },
        verification: Verification::default(),
        implementation: Implementation::default(),
        examples: markdown::examples(&page.page.body),
        provenance,
    }
}

fn enum_record(page: &ParsedPage, manifest: &Manifest) -> EnumRecord {
    let members = markdown::enum_values(&page.page.body)
        .into_iter()
        .map(|(name, documented_value, summary)| EnumValue {
            name,
            documented_value,
            summary,
        })
        .collect();
    let mut provenance = provenance(page, "source-title");
    provenance.insert("members".to_owned(), "source-table".to_owned());
    EnumRecord {
        schema_version: SCHEMA_VERSION,
        id: page.id.clone(),
        kind: "enum".to_owned(),
        name: page.name.clone(),
        namespace: "Excel".to_owned(),
        summary: page.summary.clone(),
        members,
        source: source_ref(&manifest.source, &page.path),
        documentation: Documentation {
            title: page.page.title.clone().unwrap_or_default(),
            api_names: page.page.api_names.clone(),
        },
        verification: Verification::default(),
        implementation: Implementation::default(),
        examples: markdown::examples(&page.page.body),
        provenance,
    }
}

fn relationship(
    kind: &str,
    from: &str,
    member: Option<&str>,
    to: &str,
    source: &SourceRef,
    origin: &str,
) -> RelationshipRecord {
    let member_component = member.unwrap_or("none");
    RelationshipRecord {
        schema_version: SCHEMA_VERSION,
        id: format!("{from}--{kind}--{member_component}--{to}"),
        kind: kind.to_owned(),
        from: from.to_owned(),
        member: member.map(str::to_owned),
        to: to.to_owned(),
        source: source.clone(),
        verification: Verification::default(),
        implementation: Implementation::default(),
        provenance: BTreeMap::from([
            ("relationship".to_owned(), origin.to_owned()),
            ("source".to_owned(), "source-manifest".to_owned()),
            (
                "verification".to_owned(),
                "project-policy-documentation".to_owned(),
            ),
            (
                "implementation.status".to_owned(),
                "project-policy".to_owned(),
            ),
        ]),
    }
}

fn source_ref(source: &ManifestSource, path: &str) -> SourceRef {
    SourceRef {
        repository: source.repository.clone(),
        commit: source.commit.clone(),
        path: path.to_owned(),
        extraction_method: "markdown-frontmatter-and-structure".to_owned(),
    }
}

fn provenance(page: &ParsedPage, classification_origin: &str) -> BTreeMap<String, String> {
    BTreeMap::from([
        ("name".to_owned(), "source-frontmatter".to_owned()),
        ("summary".to_owned(), page.summary_origin.clone()),
        ("source".to_owned(), "source-manifest".to_owned()),
        ("documentation".to_owned(), "source-frontmatter".to_owned()),
        (
            "verification".to_owned(),
            "project-policy-documentation".to_owned(),
        ),
        (
            "implementation.status".to_owned(),
            "project-policy".to_owned(),
        ),
        (
            "classification".to_owned(),
            classification_origin.to_owned(),
        ),
    ])
}

fn canonical_id(api_name: &str) -> String {
    api_name
        .trim()
        .split('.')
        .filter(|segment| !segment.is_empty())
        .map(canonical_segment)
        .collect::<Vec<_>>()
        .join(".")
}

fn title_entity_name(title: &str, kind: &str) -> String {
    let suffix = match kind {
        "object" if title.to_ascii_lowercase().contains(" collection (excel)") => {
            " collection (Excel)"
        }
        "object" => " object (Excel)",
        "property" => " property (Excel)",
        "method" => " method (Excel)",
        "event" => " event (Excel)",
        "enum" => " enumeration (Excel)",
        _ => "",
    };
    title
        .strip_suffix(suffix)
        .unwrap_or(title)
        .trim()
        .to_owned()
}

fn canonical_segment(segment: &str) -> String {
    let mut characters = segment.trim().chars();
    match characters.next() {
        None => String::new(),
        Some(first) => first.to_uppercase().chain(characters).collect(),
    }
}

fn resolve_object_name(name: &str, object_ids: &BTreeSet<String>) -> Option<String> {
    let normalized = canonical_segment(name.trim().trim_matches('.'));
    object_ids
        .iter()
        .find(|id| {
            id.rsplit('.')
                .next()
                .is_some_and(|candidate| candidate.eq_ignore_ascii_case(&normalized))
        })
        .cloned()
}

fn resolve_object_id(id: &str, object_ids: &BTreeSet<String>) -> Option<String> {
    object_ids
        .iter()
        .find(|candidate| candidate.eq_ignore_ascii_case(id))
        .cloned()
}

fn resolve_enum_name(name: &str, enum_ids: &BTreeSet<String>) -> Option<String> {
    let normalized = canonical_segment(name.trim().trim_matches('.'));
    enum_ids
        .iter()
        .find(|id| {
            id.rsplit('.')
                .next()
                .is_some_and(|candidate| candidate.eq_ignore_ascii_case(&normalized))
        })
        .cloned()
}

fn looks_like_excel_object(value: &str) -> bool {
    let value = value.trim();
    value.chars().next().is_some_and(char::is_uppercase)
        && !matches!(
            value,
            "Variant" | "String" | "Boolean" | "Long" | "Double" | "Integer" | "Nothing"
        )
}

fn issue(category: &str, id: Option<String>, path: &str, detail: &str) -> Issue {
    Issue {
        category: category.to_owned(),
        id,
        path: path.to_owned(),
        detail: detail.to_owned(),
    }
}

fn sort_issues(coverage: &mut SourceCoverage) {
    for list in [
        &mut coverage.skipped_files,
        &mut coverage.warnings,
        &mut coverage.orphan_records,
        &mut coverage.duplicate_ids,
        &mut coverage.unresolved_owner_references,
        &mut coverage.unresolved_return_object_references,
        &mut coverage.missing_summaries,
        &mut coverage.missing_source_paths,
        &mut coverage.unresolved,
    ] {
        list.sort_by(|left, right| {
            left.category
                .cmp(&right.category)
                .then(left.id.cmp(&right.id))
                .then(left.path.cmp(&right.path))
                .then(left.detail.cmp(&right.detail))
        });
        list.dedup();
    }
}

fn write_jsonl<T: Serialize>(path: &Path, records: &[T]) -> Result<(), String> {
    let mut output = String::new();
    for record in records {
        output.push_str(
            &serde_json::to_string(record)
                .map_err(|error| format!("cannot encode {}: {error}", path.display()))?,
        );
        output.push('\n');
    }
    fs::write(path, output).map_err(|error| format!("cannot write {}: {error}", path.display()))
}

fn write_json<T: Serialize>(path: &Path, value: &T) -> Result<(), String> {
    let mut output = serde_json::to_string_pretty(value)
        .map_err(|error| format!("cannot encode {}: {error}", path.display()))?;
    output.push('\n');
    fs::write(path, output).map_err(|error| format!("cannot write {}: {error}", path.display()))
}
