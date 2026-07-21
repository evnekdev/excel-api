use crate::model::{Manifest, ManifestSelection};
use std::fs;
use std::path::{Path, PathBuf};
use std::process::Command;

pub fn load_manifest(path: &Path) -> Result<Manifest, String> {
    let input = fs::read_to_string(path)
        .map_err(|error| format!("cannot read manifest {}: {error}", path.display()))?;
    let manifest: Manifest = toml::from_str(&input)
        .map_err(|error| format!("invalid manifest {}: {error}", path.display()))?;
    if manifest.transform.tool_version != 1 {
        return Err(format!(
            "manifest {} has unsupported tool_version {}",
            path.display(),
            manifest.transform.tool_version
        ));
    }
    if manifest.source.commit.len() != 40
        || !manifest
            .source
            .commit
            .bytes()
            .all(|byte| byte.is_ascii_hexdigit())
    {
        return Err(format!(
            "manifest {} must contain a full 40-character Git commit",
            path.display()
        ));
    }
    Ok(manifest)
}

pub fn selected_files(
    root: &Path,
    selection: &ManifestSelection,
) -> Result<Vec<(String, PathBuf)>, String> {
    let mut found = Vec::new();
    walk(root, root, &mut found)?;
    found.retain(|(relative, _)| {
        selection
            .include
            .iter()
            .any(|pattern| path_matches(relative, pattern))
            && !selection
                .exclude
                .iter()
                .any(|pattern| path_matches(relative, pattern))
    });
    found.sort_by(|left, right| left.0.cmp(&right.0));
    Ok(found)
}

pub fn verify_checkout_commit(root: &Path, expected: &str) -> Result<(), String> {
    if !root.join(".git").exists() {
        return Ok(());
    }
    let output = Command::new("git")
        .arg("-C")
        .arg(root)
        .args(["rev-parse", "HEAD"])
        .output()
        .map_err(|error| format!("cannot run Git to verify {}: {error}", root.display()))?;
    if !output.status.success() {
        return Err(format!(
            "cannot determine source commit for {}",
            root.display()
        ));
    }
    let actual = String::from_utf8_lossy(&output.stdout)
        .trim()
        .to_ascii_lowercase();
    if actual != expected.to_ascii_lowercase() {
        return Err(format!(
            "source checkout commit {actual} does not match manifest commit {expected}"
        ));
    }
    Ok(())
}

fn walk(root: &Path, current: &Path, result: &mut Vec<(String, PathBuf)>) -> Result<(), String> {
    let entries = fs::read_dir(current)
        .map_err(|error| format!("cannot read {}: {error}", current.display()))?;
    for entry in entries {
        let entry =
            entry.map_err(|error| format!("cannot enumerate {}: {error}", current.display()))?;
        let file_type = entry
            .file_type()
            .map_err(|error| format!("cannot inspect {}: {error}", entry.path().display()))?;
        if file_type.is_dir() {
            if entry.file_name() != ".git" {
                walk(root, &entry.path(), result)?;
            }
        } else if file_type.is_file() {
            let relative = entry
                .path()
                .strip_prefix(root)
                .expect("walk root prefix")
                .to_string_lossy()
                .replace('\\', "/");
            result.push((relative, entry.path()));
        }
    }
    Ok(())
}

fn path_matches(path: &str, pattern: &str) -> bool {
    let path_parts: Vec<&str> = path.split('/').collect();
    let pattern_parts: Vec<&str> = pattern.split('/').collect();
    matches_parts(&path_parts, &pattern_parts)
}

fn matches_parts(path: &[&str], pattern: &[&str]) -> bool {
    match pattern.split_first() {
        None => path.is_empty(),
        Some((&"**", rest)) => {
            matches_parts(path, rest) || (!path.is_empty() && matches_parts(&path[1..], pattern))
        }
        Some((segment, rest)) => path.split_first().is_some_and(|(part, path_rest)| {
            segment_matches(part, segment) && matches_parts(path_rest, rest)
        }),
    }
}

fn segment_matches(value: &str, pattern: &str) -> bool {
    let (mut value_index, mut pattern_index, mut star) = (0usize, 0usize, None);
    let value_bytes = value.as_bytes();
    let pattern_bytes = pattern.as_bytes();
    let mut retry = 0usize;
    while value_index < value_bytes.len() {
        if pattern_index < pattern_bytes.len()
            && (pattern_bytes[pattern_index] == value_bytes[value_index])
        {
            value_index += 1;
            pattern_index += 1;
        } else if pattern_index < pattern_bytes.len() && pattern_bytes[pattern_index] == b'*' {
            star = Some(pattern_index);
            pattern_index += 1;
            retry = value_index;
        } else if let Some(star_index) = star {
            pattern_index = star_index + 1;
            retry += 1;
            value_index = retry;
        } else {
            return false;
        }
    }
    while pattern_index < pattern_bytes.len() && pattern_bytes[pattern_index] == b'*' {
        pattern_index += 1;
    }
    pattern_index == pattern_bytes.len()
}
