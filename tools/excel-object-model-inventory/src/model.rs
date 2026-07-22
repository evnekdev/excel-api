use std::collections::BTreeSet;

pub const SCHEMA_VERSION: u32 = 1;
pub const IMPLEMENTATION_STATUSES: &[&str] = &[
    "not-planned",
    "not-started",
    "metadata-only",
    "stub",
    "partial",
    "implemented",
    "blocked",
    "unsupported",
];
pub const DOCUMENTATION_STATUSES: &[&str] =
    &["not-started", "generated", "review-needed", "reviewed"];
pub const TEST_STATUSES: &[&str] = &[
    "not-tested",
    "unit-tested",
    "integration-tested",
    "live-tested",
    "blocked",
];
pub const CONFIDENCE_STATUSES: &[&str] = &[
    "typelib-only",
    "docs-only",
    "typelib-and-docs",
    "runtime-confirmed",
    "conflicting",
    "unknown",
];
/// Controlled classification of how broadly an object belongs to the initial wrapper surface.
pub const SURFACE_CLASSES: &[&str] = &[
    "implemented-wrapper",
    "priority-inventory",
    "deferred-inventory",
    "event-surface",
];

pub fn slug(value: &str) -> String {
    let mut result = String::new();
    let mut hyphen = false;
    for character in value.trim_start_matches('_').chars() {
        if character.is_ascii_alphanumeric() {
            result.push(character.to_ascii_lowercase());
            hyphen = false;
        } else if !hyphen {
            result.push('-');
            hyphen = true;
        }
    }
    result.trim_matches('-').to_owned()
}

pub fn object_id(name: &str) -> String {
    format!("excel.{}", slug(canonical_name(name)))
}
pub fn canonical_name(name: &str) -> &str {
    match name {
        "_Application" => "Application",
        "_Workbook" => "Workbook",
        "_Worksheet" => "Worksheet",
        other => other.trim_start_matches('_'),
    }
}
pub fn documentation_url(name: &str) -> Option<&'static str> {
    match canonical_name(name) {
        "Application" => {
            Some("https://learn.microsoft.com/en-us/office/vba/api/excel.application(object)")
        }
        "Workbooks" => Some("https://learn.microsoft.com/en-us/office/vba/api/excel.workbooks"),
        "Workbook" => Some("https://learn.microsoft.com/en-us/office/vba/api/excel.workbook"),
        "Worksheets" => Some("https://learn.microsoft.com/en-us/office/vba/api/excel.worksheets"),
        "Worksheet" => Some("https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet"),
        "Range" => Some("https://learn.microsoft.com/en-us/office/vba/api/excel.range(object)"),
        _ => None,
    }
}
pub fn priority_object(name: &str) -> bool {
    matches!(
        canonical_name(name),
        "Application" | "Workbooks" | "Workbook" | "Worksheets" | "Worksheet" | "Range"
    )
}
pub fn wrapper_object(name: &str) -> bool {
    matches!(
        canonical_name(name),
        "Application" | "Workbooks" | "Workbook" | "Worksheets" | "Worksheet" | "Range"
    )
}
pub fn surface_class(name: &str, event_interface: bool) -> &'static str {
    if event_interface {
        "event-surface"
    } else if wrapper_object(name) {
        "implemented-wrapper"
    } else if priority_object(name) {
        "priority-inventory"
    } else {
        "deferred-inventory"
    }
}
pub fn implemented_member_ids() -> BTreeSet<&'static str> {
    excel_com::IMPLEMENTED_MEMBER_IDS.iter().copied().collect()
}
pub fn member_id(object: &str, member: &str) -> String {
    format!("{}.{}", object_id(object), slug(member))
}
