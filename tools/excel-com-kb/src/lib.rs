#![forbid(unsafe_code)]

pub mod generate;
pub mod markdown;
pub mod model;
pub mod normalize;
pub mod source;
pub mod validate;

use std::path::Path;

pub fn ingest(
    source: &Path,
    manifest: &Path,
    output: &Path,
) -> Result<model::IngestResult, String> {
    normalize::ingest(source, manifest, output)
}

pub fn validate(root: &Path) -> Result<validate::ValidationSummary, String> {
    validate::validate_root(root)
}

pub fn generate(root: &Path) -> Result<generate::GeneratedSummary, String> {
    generate::generate_reports(root)
}

pub fn check(root: &Path) -> Result<(), String> {
    validate(root)?;
    generate::check_reports(root)
}
