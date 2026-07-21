use serde::{Deserialize, Serialize};
use std::collections::BTreeMap;

pub const SCHEMA_VERSION: u32 = 1;
pub const MAX_SUMMARY_CHARS: usize = 240;

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
#[serde(deny_unknown_fields)]
pub struct SourceRef {
    pub repository: String,
    pub commit: String,
    pub path: String,
    pub extraction_method: String,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
#[serde(deny_unknown_fields)]
pub struct Documentation {
    pub title: String,
    pub api_names: Vec<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
#[serde(deny_unknown_fields)]
pub struct Verification {
    pub documentation: bool,
    pub typelib: bool,
    pub runtime: bool,
}

impl Default for Verification {
    fn default() -> Self {
        Self {
            documentation: true,
            typelib: false,
            runtime: false,
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
#[serde(deny_unknown_fields)]
pub struct Implementation {
    pub status: String,
}

impl Default for Implementation {
    fn default() -> Self {
        Self {
            status: "unplanned".to_owned(),
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
#[serde(deny_unknown_fields)]
pub struct Examples {
    pub count: usize,
    pub anchors: Vec<String>,
    pub copied: bool,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
#[serde(deny_unknown_fields)]
pub struct Classification {
    pub is_collection: bool,
    pub collection_item_type: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
#[serde(deny_unknown_fields)]
pub struct ObjectRecord {
    pub schema_version: u32,
    pub id: String,
    pub kind: String,
    pub name: String,
    pub namespace: String,
    pub summary: String,
    pub source: SourceRef,
    pub documentation: Documentation,
    pub classification: Classification,
    pub verification: Verification,
    pub implementation: Implementation,
    pub examples: Examples,
    pub provenance: BTreeMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
#[serde(deny_unknown_fields)]
pub struct Parameter {
    pub name: String,
    pub optionality: String,
    pub documented_type: Option<String>,
    pub summary: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
#[serde(deny_unknown_fields)]
pub struct Dispatch {
    pub dispid: Option<i32>,
    pub invoke_kinds: Vec<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
#[serde(deny_unknown_fields)]
pub struct MemberRecord {
    pub schema_version: u32,
    pub id: String,
    pub kind: String,
    pub owner: String,
    pub name: String,
    pub summary: String,
    pub access: String,
    pub parameters: Vec<Parameter>,
    pub documented_return_type: Option<String>,
    pub source: SourceRef,
    pub documentation: Documentation,
    pub dispatch: Dispatch,
    pub verification: Verification,
    pub implementation: Implementation,
    pub examples: Examples,
    pub provenance: BTreeMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
#[serde(deny_unknown_fields)]
pub struct EnumValue {
    pub name: String,
    pub documented_value: Option<String>,
    pub summary: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
#[serde(deny_unknown_fields)]
pub struct EnumRecord {
    pub schema_version: u32,
    pub id: String,
    pub kind: String,
    pub name: String,
    pub namespace: String,
    pub summary: String,
    pub members: Vec<EnumValue>,
    pub source: SourceRef,
    pub documentation: Documentation,
    pub verification: Verification,
    pub implementation: Implementation,
    pub examples: Examples,
    pub provenance: BTreeMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
#[serde(deny_unknown_fields)]
pub struct RelationshipRecord {
    pub schema_version: u32,
    pub id: String,
    pub kind: String,
    pub from: String,
    pub member: Option<String>,
    pub to: String,
    pub source: SourceRef,
    pub verification: Verification,
    pub implementation: Implementation,
    pub provenance: BTreeMap<String, String>,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
#[serde(deny_unknown_fields)]
pub struct ManifestSource {
    pub name: String,
    pub repository: String,
    pub commit: String,
    pub retrieved: String,
    pub documentation_license: String,
    pub sample_code_license: String,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
#[serde(deny_unknown_fields)]
pub struct ManifestSelection {
    pub include: Vec<String>,
    pub exclude: Vec<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
#[serde(deny_unknown_fields)]
pub struct ManifestTransform {
    pub tool_version: u32,
    pub summary_policy: String,
    pub copy_examples: bool,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
#[serde(deny_unknown_fields)]
pub struct Manifest {
    pub source: ManifestSource,
    pub selection: ManifestSelection,
    pub transform: ManifestTransform,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
#[serde(deny_unknown_fields)]
pub struct Issue {
    pub category: String,
    pub id: Option<String>,
    pub path: String,
    pub detail: String,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
#[serde(deny_unknown_fields)]
pub struct SourceCoverage {
    pub schema_version: u32,
    pub source: ManifestSource,
    pub selected_files: usize,
    pub parsed_files: usize,
    pub skipped_files: Vec<Issue>,
    pub warnings: Vec<Issue>,
    pub orphan_records: Vec<Issue>,
    pub duplicate_ids: Vec<Issue>,
    pub unresolved_owner_references: Vec<Issue>,
    pub unresolved_return_object_references: Vec<Issue>,
    pub missing_summaries: Vec<Issue>,
    pub missing_source_paths: Vec<Issue>,
    pub unresolved: Vec<Issue>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct IngestResult {
    pub objects: usize,
    pub members: usize,
    pub enums: usize,
    pub relationships: usize,
    pub coverage: SourceCoverage,
}
