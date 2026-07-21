use excel_com_kb::{analyze, check, generate, ingest, validate};
use std::fs;
use std::path::{Path, PathBuf};

fn fixture(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join(name)
}

fn copy_directory(from: &Path, to: &Path) {
    fs::create_dir_all(to).expect("directory");
    for entry in fs::read_dir(from).expect("read directory") {
        let entry = entry.expect("entry");
        let target = to.join(entry.file_name());
        if entry.file_type().expect("type").is_dir() {
            copy_directory(&entry.path(), &target);
        } else {
            fs::copy(entry.path(), target).expect("copy");
        }
    }
}

#[test]
fn ingestion_and_generation_are_byte_deterministic_and_portable() {
    let first = tempfile::tempdir().expect("first");
    let second = tempfile::tempdir().expect("second");
    ingest(
        &fixture("vba-docs"),
        &fixture("manifest.toml"),
        first.path(),
    )
    .expect("first ingest");
    ingest(
        &fixture("vba-docs"),
        &fixture("manifest.toml"),
        second.path(),
    )
    .expect("second ingest");
    for name in [
        "objects.jsonl",
        "members.jsonl",
        "relationships.jsonl",
        "enums.jsonl",
        "source.json",
    ] {
        let first_bytes = fs::read(first.path().join(name)).expect("first data");
        let second_bytes = fs::read(second.path().join(name)).expect("second data");
        assert_eq!(first_bytes, second_bytes, "{name}");
        assert!(first_bytes.ends_with(b"\n"));
        assert!(!first_bytes.windows(2).any(|bytes| bytes == b"\r\n"));
        assert!(!String::from_utf8_lossy(&first_bytes).contains("C:\\"));
    }
    let root = tempfile::tempdir().expect("root");
    copy_directory(first.path(), &root.path().join("data"));
    copy_directory(
        &PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../knowledge/excel-object-model/schema"),
        &root.path().join("schema"),
    );
    validate(root.path()).expect("valid fixture root");
    generate(root.path()).expect("first reports");
    let first_reports = fs::read(root.path().join("generated/object-index.md")).expect("report");
    generate(root.path()).expect("second reports");
    assert_eq!(
        first_reports,
        fs::read(root.path().join("generated/object-index.md")).expect("second report")
    );
    analyze(root.path()).expect("first analysis");
    let first_analysis = fs::read(
        root.path()
            .join("generated/analysis/architectural-spine.md"),
    )
    .expect("analysis report");
    analyze(root.path()).expect("second analysis");
    assert_eq!(
        first_analysis,
        fs::read(
            root.path()
                .join("generated/analysis/architectural-spine.md")
        )
        .expect("second analysis report")
    );
    check(root.path()).expect("checked analysis reports");
}
