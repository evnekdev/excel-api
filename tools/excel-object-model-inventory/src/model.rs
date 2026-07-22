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
        "ListObjects" => Some("https://learn.microsoft.com/en-us/office/vba/api/excel.listobjects"),
        "ListObject" => Some("https://learn.microsoft.com/en-us/office/vba/api/excel.listobject"),
        "ListColumns" => Some("https://learn.microsoft.com/en-us/office/vba/api/excel.listcolumns"),
        "ListColumn" => Some("https://learn.microsoft.com/en-us/office/vba/api/excel.listcolumn"),
        "ListRows" => Some("https://learn.microsoft.com/en-us/office/vba/api/excel.listrows"),
        "ListRow" => Some("https://learn.microsoft.com/en-us/office/vba/api/excel.listrow"),
        "AutoFilter" => Some("https://learn.microsoft.com/en-us/office/vba/api/excel.autofilter"),
        "Filters" => Some("https://learn.microsoft.com/en-us/office/vba/api/excel.filters"),
        "Filter" => Some("https://learn.microsoft.com/en-us/office/vba/api/excel.filter"),
        "Sort" => Some("https://learn.microsoft.com/en-us/office/vba/api/excel.sort"),
        "SortFields" => Some("https://learn.microsoft.com/en-us/office/vba/api/excel.sortfields"),
        "SortField" => Some("https://learn.microsoft.com/en-us/office/vba/api/excel.sortfield"),
        "Validation" => Some("https://learn.microsoft.com/en-us/office/vba/api/excel.validation"),
        "Sheets" => Some("https://learn.microsoft.com/en-us/office/vba/api/excel.sheets"),
        "Windows" => Some("https://learn.microsoft.com/en-us/office/vba/api/excel.windows"),
        "Window" => Some("https://learn.microsoft.com/en-us/office/vba/api/excel.window"),
        "PageSetup" => Some("https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup"),
        "Tab" => Some("https://learn.microsoft.com/en-us/office/vba/api/excel.tab"),
        "Outline" => Some("https://learn.microsoft.com/en-us/office/vba/api/excel.outline"),
        "HPageBreaks" => Some("https://learn.microsoft.com/en-us/office/vba/api/excel.hpagebreaks"),
        "HPageBreak" => Some("https://learn.microsoft.com/en-us/office/vba/api/excel.hpagebreak"),
        "VPageBreaks" => Some("https://learn.microsoft.com/en-us/office/vba/api/excel.vpagebreaks"),
        "VPageBreak" => Some("https://learn.microsoft.com/en-us/office/vba/api/excel.vpagebreak"),
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
            | "ListObjects"
            | "ListObject"
            | "ListColumns"
            | "ListColumn"
            | "ListRows"
            | "ListRow"
            | "AutoFilter"
            | "Filters"
            | "Filter"
            | "Sort"
            | "SortFields"
            | "SortField"
            | "Validation"
            | "Sheets"
            | "Windows"
            | "Window"
            | "PageSetup"
            | "Tab"
            | "Outline"
            | "HPageBreaks"
            | "HPageBreak"
            | "VPageBreaks"
            | "VPageBreak"
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
            | "ListObjects"
            | "ListObject"
            | "ListColumns"
            | "ListColumn"
            | "ListRows"
            | "ListRow"
            | "AutoFilter"
            | "Filters"
            | "Filter"
            | "Sort"
            | "SortFields"
            | "SortField"
            | "Validation"
            | "Sheets"
            | "Windows"
            | "Window"
            | "PageSetup"
            | "Tab"
            | "Outline"
            | "HPageBreaks"
            | "HPageBreak"
            | "VPageBreaks"
            | "VPageBreak"
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

/// Returns whether an implemented member is presently blocked before live
/// acceptance coverage by the recorded Prompt 15 `Workbooks.Add` baseline.
pub fn runtime_blocked_member(id: &str) -> bool {
    matches!(
        id,
        "excel.application.activecell"
            | "excel.application.activesheet"
            | "excel.application.activewindow"
            | "excel.application.activeworkbook"
            | "excel.application.automationsecurity"
            | "excel.application.goto"
            | "excel.application.run"
            | "excel.application.selection"
            | "excel.application.sheets"
            | "excel.application.windows"
            | "excel.application.worksheets"
            | "excel.workbooks.application"
            | "excel.workbook.activate"
            | "excel.workbook.activesheet"
            | "excel.workbook.exportasfixedformat-3175"
            | "excel.workbook.hasvbproject"
            | "excel.workbook.protectstructure"
            | "excel.workbook.protectwindows"
            | "excel.workbook.printout-2361"
            | "excel.workbook.printpreview"
            | "excel.workbook.protect-2029"
            | "excel.workbook.sheets"
            | "excel.workbook.unprotect"
            | "excel.workbook.windows"
            | "excel.worksheet.activate"
            | "excel.worksheet.copy"
            | "excel.worksheet.delete"
            | "excel.worksheet.exportasfixedformat-3175"
            | "excel.worksheet.hpagebreaks"
            | "excel.worksheet.move"
            | "excel.worksheet.outline"
            | "excel.worksheet.pagesetup"
            | "excel.worksheet.printout-2361"
            | "excel.worksheet.printpreview"
            | "excel.worksheet.protect-2029"
            | "excel.worksheet.protectcontents"
            | "excel.worksheet.protectdrawingobjects"
            | "excel.worksheet.protectionmode"
            | "excel.worksheet.protectscenarios"
            | "excel.worksheet.resetallpagebreaks"
            | "excel.worksheet.select"
            | "excel.worksheet.tab"
            | "excel.worksheet.type"
            | "excel.worksheet.unprotect"
            | "excel.worksheet.vpagebreaks"
    ) || id.starts_with("excel.sheets.")
        || id.starts_with("excel.windows.")
        || id.starts_with("excel.window.")
        || id.starts_with("excel.pagesetup.")
        || id.starts_with("excel.tab.")
        || id.starts_with("excel.outline.")
        || id.starts_with("excel.hpagebreak")
        || id.starts_with("excel.vpagebreak")
        || matches!(
            id,
            "excel.range.activate"
                | "excel.range.exportasfixedformat-3175"
                | "excel.range.group"
                | "excel.range.locked"
                | "excel.range.formulahidden"
                | "excel.range.indentlevel"
                | "excel.range.merge"
                | "excel.range.mergearea"
                | "excel.range.mergecells"
                | "excel.range.orientation"
                | "excel.range.printout-2361"
                | "excel.range.printpreview"
                | "excel.range.readingorder"
                | "excel.range.select"
                | "excel.range.showdetail"
                | "excel.range.shrinktofit"
                | "excel.range.ungroup"
                | "excel.range.unmerge"
        )
}

/// Returns whether a new Prompt 15 object family has no passing live test yet.
pub fn runtime_blocked_object(name: &str) -> bool {
    matches!(
        canonical_name(name),
        "Sheets"
            | "Windows"
            | "Window"
            | "PageSetup"
            | "Tab"
            | "Outline"
            | "HPageBreaks"
            | "HPageBreak"
            | "VPageBreaks"
            | "VPageBreak"
    )
}
