use excel_com_kb::ingest;
use std::path::PathBuf;

fn fixture(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join(name)
}

#[test]
fn ingests_objects_members_events_collections_and_enums() {
    let output = tempfile::tempdir().expect("temporary output");
    let result = ingest(
        &fixture("vba-docs"),
        &fixture("manifest.toml"),
        output.path(),
    )
    .expect("fixture ingest");
    assert_eq!(result.objects, 4);
    assert_eq!(result.members, 5);
    assert_eq!(result.enums, 1);
    assert!(
        result
            .coverage
            .skipped_files
            .iter()
            .any(|issue| issue.category == "missing-front-matter")
    );
    assert!(
        result
            .coverage
            .warnings
            .iter()
            .any(|issue| issue.category == "parameter-table")
    );
    assert!(
        result
            .coverage
            .unresolved_return_object_references
            .iter()
            .any(|issue| issue.id.as_deref() == Some("Excel.Workbook.UnknownResult"))
    );
    let members = std::fs::read_to_string(output.path().join("members.jsonl")).expect("members");
    assert!(members.contains("\"id\":\"Excel.Workbooks.Open\""));
    assert!(members.contains("\"optionality\":\"optional\""));
    assert!(members.contains("\"id\":\"Excel.Application.WorkbookOpen\""));
    assert!(members.contains("\"summary\":\"source-short-description\""));
    let objects = std::fs::read_to_string(output.path().join("objects.jsonl")).expect("objects");
    assert!(objects.contains("\"collection_item_type\":\"Workbook\""));
    assert!(objects.contains("Unicode Ω value"));
    let enums = std::fs::read_to_string(output.path().join("enums.jsonl")).expect("enums");
    assert!(enums.contains("\"id\":\"Excel.XlMode\""));
}

#[test]
fn rejects_duplicate_stable_ids() {
    let output = tempfile::tempdir().expect("temporary output");
    let error = ingest(
        &fixture("duplicate-vba-docs"),
        &fixture("manifest.toml"),
        output.path(),
    )
    .expect_err("duplicate should fail");
    assert!(error.contains("duplicate stable ID"));
}
