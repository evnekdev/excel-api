use excel_com_kb::{ingest, validate};
use std::fs;
use std::io::Write;
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
fn validates_checked_in_schema_contract_and_rejects_duplicate_records() {
    let root = tempfile::tempdir().expect("root");
    let data = root.path().join("data");
    ingest(&fixture("vba-docs"), &fixture("manifest.toml"), &data).expect("ingest");
    copy_directory(
        &PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../knowledge/excel-object-model/schema"),
        &root.path().join("schema"),
    );
    validate(root.path()).expect("valid root");
    let objects = data.join("objects.jsonl");
    let original = fs::read_to_string(&objects).expect("objects");
    let with_unknown_field = original.replacen("}\n", ",\"unexpected\":true}\n", 1);
    fs::write(&objects, with_unknown_field).expect("write unknown field");
    let error = validate(root.path()).expect_err("unknown fields must fail");
    assert!(error.contains("unknown field"));
    fs::write(&objects, &original).expect("restore objects");
    let members = data.join("members.jsonl");
    let original_members = fs::read_to_string(&members).expect("members");
    fs::write(
        &members,
        original_members.replacen(
            "\"owner\":\"Excel.Application\"",
            "\"owner\":\"Excel.Missing\"",
            1,
        ),
    )
    .expect("write invalid owner");
    let error = validate(root.path()).expect_err("unknown owners must fail");
    assert!(error.contains("unknown owner"));
    fs::write(&members, &original_members).expect("restore members");
    let duplicate = original.lines().next().expect("first object").to_owned();
    fs::OpenOptions::new()
        .append(true)
        .open(&objects)
        .expect("append")
        .write_all(format!("{duplicate}\n").as_bytes())
        .expect("write duplicate");
    let error = validate(root.path()).expect_err("duplicate records must fail");
    assert!(error.contains("strict stable ID order") || error.contains("duplicate canonical ID"));
}
