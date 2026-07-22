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
/// Controlled classification of an object's type-library surface.
pub const SURFACE_CLASSES: &[&str] = &[
    "primary-object-model",
    "secondary-public",
    "hidden",
    "restricted",
    "event-interface",
    "coclass",
    "legacy",
    "internal",
    "unknown",
];
/// Controlled implementation-roadmap classification independent of type-library surface.
pub const ROADMAP_CLASSES: &[&str] = &[
    "implemented-wrapper",
    "priority-inventory",
    "deferred-inventory",
];
/// Controlled provenance for a member listed by a type library.
pub const MEMBER_ORIGINS: &[&str] = &[
    "declared",
    "inherited-iunknown",
    "inherited-idispatch",
    "inherited-base-interface",
];
pub const COLLECTION_INDEX_KINDS: &[&str] = &[
    "one-based-integer",
    "string-key",
    "enum-key",
    "variant-key",
    "no-index",
    "unknown",
];
pub const COLLECTION_ITERATOR_STATUSES: &[&str] = &[
    "not-started",
    "metadata-only",
    "implemented",
    "blocked",
    "unsupported",
];
/// Controlled styles used by reference-input and reference-output metadata.
pub const REFERENCE_STYLES: &[&str] = &["a1", "r1c1"];
/// Controlled result categories for explicitly typed Excel evaluation APIs.
pub const EVALUATION_RESULT_CATEGORIES: &[&str] = &[
    "automation-value",
    "range-object",
    "other-object",
    "unknown",
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
        "Areas" => Some("https://learn.microsoft.com/en-us/office/vba/api/excel.areas"),
        "Names" => Some("https://learn.microsoft.com/en-us/office/vba/api/excel.names"),
        "Name" => Some("https://learn.microsoft.com/en-us/office/vba/api/excel.name(object)"),
        "Font" => Some("https://learn.microsoft.com/en-us/office/vba/api/excel.font"),
        "Interior" => Some("https://learn.microsoft.com/en-us/office/vba/api/excel.interior"),
        "Borders" => Some("https://learn.microsoft.com/en-us/office/vba/api/excel.borders"),
        "Border" => Some("https://learn.microsoft.com/en-us/office/vba/api/excel.border"),
        _ => None,
    }
}
pub fn priority_object(name: &str) -> bool {
    matches!(
        canonical_name(name),
        "Application"
            | "Workbooks"
            | "Workbook"
            | "Worksheets"
            | "Worksheet"
            | "Range"
            | "Areas"
            | "Names"
            | "Name"
            | "Font"
            | "Interior"
            | "Borders"
            | "Border"
    )
}
pub fn wrapper_object(name: &str) -> bool {
    matches!(
        canonical_name(name),
        "Application"
            | "Workbooks"
            | "Workbook"
            | "Worksheets"
            | "Worksheet"
            | "Range"
            | "Areas"
            | "Names"
            | "Name"
            | "Font"
            | "Interior"
            | "Borders"
            | "Border"
    )
}
pub fn surface_class(
    name: &str,
    kind: &str,
    event_interface: bool,
    type_flags: u16,
) -> &'static str {
    if event_interface {
        "event-interface"
    } else if kind == "coclass" {
        "coclass"
    } else if type_flags & 0x0200 != 0 {
        "restricted"
    } else if type_flags & 0x0010 != 0 {
        "hidden"
    } else if wrapper_object(name) {
        "primary-object-model"
    } else if kind == "alias" {
        "internal"
    } else if name.starts_with('_') {
        "legacy"
    } else if matches!(kind, "dispatch-interface" | "interface") {
        "secondary-public"
    } else {
        "unknown"
    }
}
pub fn roadmap_class(name: &str, kind: &str, event_interface: bool) -> &'static str {
    if event_interface || kind == "coclass" || kind == "alias" {
        "deferred-inventory"
    } else if wrapper_object(name) && kind == "dispatch-interface" {
        "implemented-wrapper"
    } else if priority_object(name) {
        "priority-inventory"
    } else {
        "deferred-inventory"
    }
}
pub fn member_origin(name: &str) -> &'static str {
    match name {
        "QueryInterface" | "AddRef" | "Release" => "inherited-iunknown",
        "GetTypeInfoCount" | "GetTypeInfo" | "GetIDsOfNames" | "Invoke" => "inherited-idispatch",
        _ => "declared",
    }
}
pub fn implemented_member_ids() -> BTreeSet<&'static str> {
    excel_com::IMPLEMENTED_MEMBER_IDS.iter().copied().collect()
}
pub fn member_id(object: &str, member: &str) -> String {
    format!("{}.{}", object_id(object), slug(member))
}
