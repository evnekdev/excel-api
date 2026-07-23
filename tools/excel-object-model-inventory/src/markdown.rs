use std::collections::BTreeMap;
use std::fs;
use std::path::{Path, PathBuf};

use serde_json::Value;

use crate::model;

pub struct Summary {
    pub pages: usize,
    pub enum_pages: usize,
    pub indexes: usize,
}

pub fn generate(root: &Path) -> Result<Summary, String> {
    let outputs = planned_outputs(root)?;
    for (path, content) in &outputs {
        write(path, content)?;
    }
    Ok(Summary {
        pages: priority_records(&read_records(
            &root.join("metadata/excel-object-model/objects"),
        )?)
        .len(),
        enum_pages: outputs
            .keys()
            .filter(|path| {
                path.to_string_lossy().contains("/enums/")
                    || path.to_string_lossy().contains("\\enums\\")
            })
            .count(),
        indexes: outputs
            .keys()
            .filter(|path| path.to_string_lossy().contains("indexes"))
            .count(),
    })
}

pub fn planned_outputs(root: &Path) -> Result<BTreeMap<PathBuf, String>, String> {
    let objects = read_records(&root.join("metadata/excel-object-model/objects"))?;
    let enums = read_records(&root.join("metadata/excel-object-model/enums"))?;
    let relationships =
        read_relationships(&root.join("metadata/excel-object-model/relationships.json"))?;
    let docs = root.join("docs/excel-object-model");
    let mut output = BTreeMap::new();
    output.insert(docs.join("README.md"), "# Excel Object Model inventory\n\nThis maintained inventory is generated from the locally registered Excel type library plus explicit policy metadata. It is an implementation guide for the experimental `excel-com` crate, not a claim of complete wrapper coverage.\n\nEvery object has independent `surface_class` (what the typelib exposes) and `roadmap_class` (the wrapper plan) fields. Standard IUnknown and IDispatch entries are retained structurally but excluded from human Excel-member coverage. The experimental crate implements a bounded `Application -> Workbooks -> Workbook -> Worksheets -> Worksheet -> Range` slice, with typed workbooks, worksheets, sheets, windows, page setup, tab, outline, page-break, table, conditional-format, Style, Note, and hyperlink collections plus core Range navigation, formulas, calculation, formatting, presentation, print/fixed-format output, filtering, sorting, validation, duplicate removal, and structural edits. Advanced-presentation capability metadata is recorded on Range; typed collection metadata and its [dashboard](indexes/collections.md) describe only typed collections implemented by the crate. See [STATUS](STATUS.md) for coverage and the indexes directory for objects, members, events, enums, and deferred surface area. Historical runtime research remains in `docs/research/excel-com/`.\n".to_owned());
    for object in priority_records(&objects) {
        let file = docs.join("objects").join(format!(
            "{}.md",
            model::slug(object["name"].as_str().unwrap())
        ));
        let existing = fs::read_to_string(&file).ok();
        output.insert(
            file,
            object_page(object, &relationships, existing.as_deref()),
        );
    }
    output.insert(docs.join("indexes/objects.md"), object_index(&objects));
    output.insert(
        docs.join("indexes/properties.md"),
        member_index_root(&objects, true),
    );
    output.insert(
        docs.join("indexes/methods.md"),
        member_index_root(&objects, false),
    );
    for object in priority_records(&objects) {
        let name = object["name"].as_str().unwrap();
        output.insert(
            docs.join("indexes/properties")
                .join(format!("{}.md", model::slug(name))),
            member_index_for_objects(&[object], true, &format!("Properties: {name}")),
        );
        output.insert(
            docs.join("indexes/methods")
                .join(format!("{}.md", model::slug(name))),
            member_index_for_objects(&[object], false, &format!("Methods: {name}")),
        );
    }
    for initial in 'a'..='z' {
        let selected: Vec<&Value> = objects
            .iter()
            .filter(|object| index_initial(object) == initial)
            .collect();
        output.insert(
            docs.join("indexes/properties")
                .join(format!("all-{initial}.md")),
            member_index_for_objects(&selected, true, &format!("Properties: {initial}")),
        );
        output.insert(
            docs.join("indexes/methods")
                .join(format!("all-{initial}.md")),
            member_index_for_objects(&selected, false, &format!("Methods: {initial}")),
        );
    }
    output.insert(docs.join("indexes/events.md"), event_index(&objects));
    output.insert(
        docs.join("indexes/collections.md"),
        collection_index(&objects),
    );
    output.insert(docs.join("indexes/enums.md"), enum_index(&enums));
    output.insert(docs.join("indexes/unsupported.md"), "# Unsupported and deferred inventory\n\nThe initial crate intentionally defers broad and generic Range/Worksheet coverage, charts, events, macros, connection points, generic collections, cross-apartment marshaling, and stable API commitments. Their metadata remains structurally inventoried.\n".to_owned());
    output.insert(docs.join("STATUS.md"), status(&objects));
    for enumeration in &enums {
        let file = docs.join("enums").join(format!(
            "{}.md",
            model::slug(enumeration["name"].as_str().unwrap())
        ));
        output.insert(file, enum_page(enumeration));
    }
    Ok(output)
}

fn object_page(object: &Value, relationships: &[Value], existing: Option<&str>) -> String {
    let default_manual = format!(
        "# {}\n\n## Summary\n\n{}\n\n## Sources\n\n- registered Excel type library\n- official Microsoft documentation URL recorded in metadata\n\n",
        object["name"].as_str().unwrap(),
        summary(object["name"].as_str().unwrap())
    );
    let manual = existing.unwrap_or(&default_manual).replace(
        "Independent summary pending review.",
        summary(object["name"].as_str().unwrap()),
    );
    let generated = format!(
        "## Identity\n\n| Field | Value |\n|---|---|\n| Interface | `{}` |\n| GUID | `{}` |\n| Object kind | {} |\n| Surface class | {} |\n| Roadmap class | {} |\n| Type flags | {} |\n| Crate type | `excel_com::{}` |\n| Implementation | {} |\n| Documentation | {} |\n| Tests | {} |\n\n## Capabilities\n\n{}\n\n## Relationships\n\n{}\n\n## Properties\n\n{}\n\n## Methods\n\n{}\n\n## Events\n\n{}\n\n## Unsupported or deferred behaviour\n\nSee the global unsupported index for unimplemented object-model areas.\n",
        object["source_interface"].as_str().unwrap_or("--"),
        object["guid"].as_str().unwrap_or("--"),
        object["kind"].as_str().unwrap_or("--"),
        object["surface_class"].as_str().unwrap_or("--"),
        object["roadmap_class"].as_str().unwrap_or("--"),
        object["typelib_type_flags"],
        object["name"].as_str().unwrap_or("Object"),
        title(&object["implemented_status"]),
        title(&object["documentation_status"]),
        title(&object["test_status"]),
        capability_table(object),
        relationship_table(object, relationships),
        member_table(object, true),
        member_table(object, false),
        event_table(object)
    );
    replace_region(&manual, &generated)
}

fn capability_table(object: &Value) -> String {
    let groups = [
        ("Formula", "formula_capability"),
        ("Calculation", "calculation_capability"),
        ("Auditing and search", "auditing_search_capability"),
        ("Formatting", "formatting_capability"),
        ("Structured data", "structured_data_capability"),
        ("Advanced presentation", "advanced_presentation_capability"),
    ];
    let mut text = String::new();
    for (label, field) in groups {
        let Some(capabilities) = object[field].as_object() else {
            continue;
        };
        text.push_str(&format!(
            "### {label}\n\n| Capability | Available |\n|---|---|\n"
        ));
        for (capability, available) in capabilities {
            text.push_str(&format!("| `{capability}` | {} |\n", available));
        }
        text.push('\n');
    }
    if text.is_empty() {
        "No capability metadata is recorded for this surface.\n".to_owned()
    } else {
        text
    }
}
fn summary(name: &str) -> &'static str {
    match name {
        "Application" => {
            "The root Automation object for a locally activated Excel instance. The initial crate exposes only a deliberately small lifecycle and workbook-navigation slice."
        }
        "Workbooks" => {
            "The typed collection through which an Application exposes open workbooks, including Count, Item, and fallible _NewEnum iteration."
        }
        "Workbook" => {
            "An Excel workbook object. The initial crate supports basic identity, saved-state, and explicit close-without-saving operations."
        }
        "Worksheets" => {
            "The typed workbook worksheet collection. The bounded slice supports Count, Item, constrained Add options, and fallible _NewEnum iteration."
        }
        "Worksheet" => {
            "A worksheet object within a workbook. The bounded crate slice exposes identity, visibility, Range, and UsedRange navigation."
        }
        "Range" => {
            "The cell and rectangular-value object. The bounded crate slice supports values plus Cells, Item, Offset, Resize, Rows, Columns, Areas, EntireRow, and EntireColumn navigation."
        }
        "Areas" => "The typed collection of contiguous ranges produced by a multi-area Range.",
        "Names" => "The typed workbook- or worksheet-scoped collection of Excel defined names.",
        "Name" => {
            "One Excel defined name, which may resolve to a Range, scalar, formula, or invalid reference."
        }
        "Font" => "The apartment-bound Excel formatting object for a Range's font properties.",
        "Interior" => "The apartment-bound Excel formatting object for a Range's fill properties.",
        "Borders" => {
            "The enum-keyed, apartment-bound collection of Excel Border objects for a Range."
        }
        "Border" => "One apartment-bound Excel border selected from a Borders collection.",
        "ListObjects" => {
            "The typed worksheet collection of Excel tables, with one-based and name lookup plus fallible enumeration."
        }
        "ListObject" => {
            "An apartment-bound Excel table wrapper with bounded table, row, column, filter, and sort operations."
        }
        "ListColumns" => "The typed collection of columns belonging to one Excel table.",
        "ListColumn" => {
            "One apartment-bound Excel table column, including calculated-column and totals operations."
        }
        "ListRows" => "The typed collection of data rows belonging to one Excel table.",
        "ListRow" => "One apartment-bound data row within an Excel table.",
        "AutoFilter" => "Excel's stateful AutoFilter object for a table or worksheet range.",
        "Filters" => "The typed collection of AutoFilter field state objects.",
        "Filter" => "Read-only criteria state for one AutoFilter field.",
        "Sort" => "Excel's persistent sort configuration for a range or table.",
        "SortFields" => "The typed collection of persistent Excel sort fields.",
        "SortField" => "One configured Excel persistent sort field.",
        "Validation" => "Excel data-validation state associated with a Range.",
        _ => "This type-library object is structurally inventoried for future wrapper planning.",
    }
}
fn relationship_table(object: &Value, relationships: &[Value]) -> String {
    let id = object["id"].as_str().unwrap_or_default();
    let mut lines = vec![
        "| Relationship | Target | Status |".to_owned(),
        "|---|---|---|".to_owned(),
    ];
    for relationship in relationships
        .iter()
        .filter(|relationship| relationship["source"].as_str() == Some(id))
    {
        let member_id = relationship["member_id"].as_str().unwrap_or("--");
        let member = object["members"]
            .as_array()
            .into_iter()
            .flatten()
            .find(|member| member["id"].as_str() == Some(member_id));
        let name = member
            .and_then(|member| member["name"].as_str())
            .unwrap_or(member_id);
        let status = member
            .map(|member| title(&member["implementation_status"]))
            .unwrap_or_else(|| "Metadata Only".to_owned());
        lines.push(format!(
            "| `{name}` | `{}` | {status} |",
            relationship["target"].as_str().unwrap_or("--")
        ));
    }
    if lines.len() == 2 {
        lines.push("| -- | -- | -- |".to_owned());
    }
    lines.join("\n")
}

fn member_table(object: &Value, property: bool) -> String {
    let mut lines = vec![
        if property {
            "| Property | Access | Type | Origin | DISPID | Implementation | Docs | Tests | Notes |"
        } else {
            "| Method | Return | Arguments | Origin | DISPID | Implementation | Docs | Tests | Notes |"
        }
        .to_owned(),
        if property {
            "|---|---|---|---|---:|---|---|---|---|"
        } else {
            "|---|---|---:|---|---:|---|---|---|---|"
        }
        .to_owned(),
    ];
    for member in object["members"]
        .as_array()
        .into_iter()
        .flatten()
        .filter(|member| is_property(member) == property)
    {
        let name = member["name"].as_str().unwrap_or("--");
        let args = member["arguments"].as_array().map_or(0, Vec::len);
        let kind = member["invoke_kinds"]
            .as_array()
            .map(|values| {
                values
                    .iter()
                    .filter_map(Value::as_str)
                    .collect::<Vec<_>>()
                    .join("/")
            })
            .unwrap_or_default();
        if property {
            lines.push(format!(
                "| {name} | {kind} | {} | {} | {} | {} | {} | {} | |",
                member["return_type"].as_str().unwrap_or("--"),
                member["member_origin"].as_str().unwrap_or("--"),
                member["dispid"],
                title(&member["implementation_status"]),
                title(&member["documentation_status"]),
                title(&member["test_status"])
            ));
        } else {
            lines.push(format!(
                "| {name} | {} | {args} | {} | {} | {} | {} | {} | |",
                member["return_type"].as_str().unwrap_or("--"),
                member["member_origin"].as_str().unwrap_or("--"),
                member["dispid"],
                title(&member["implementation_status"]),
                title(&member["documentation_status"]),
                title(&member["test_status"])
            ));
        }
    }
    if lines.len() == 2 {
        lines.push("| -- | -- | -- | -- | -- | -- | -- | -- | -- |".to_owned());
    }
    lines.join("\n")
}
fn event_table(_object: &Value) -> String {
    "| Event | Arguments | DISPID | Implementation | Docs | Tests |\n|---|---:|---:|---|---|---|\n| -- | -- | -- | Not started | Generated | Not tested |".to_owned()
}
fn is_property(member: &Value) -> bool {
    member["invoke_kinds"]
        .as_array()
        .into_iter()
        .flatten()
        .any(|kind| kind.as_str().is_some_and(|kind| kind.contains("PROPERTY")))
}
fn replace_region(existing: &str, generated: &str) -> String {
    const BEGIN: &str = "<!-- BEGIN GENERATED MEMBERS -->";
    const END: &str = "<!-- END GENERATED MEMBERS -->";
    let replacement = format!("{BEGIN}\n{generated}{END}\n");
    if let (Some(begin), Some(end)) = (existing.find(BEGIN), existing.find(END)) {
        let after = end + END.len();
        format!(
            "{}{}{}",
            &existing[..begin],
            replacement,
            &existing[after..]
        )
        .trim_end()
        .to_owned()
            + "\n"
    } else {
        format!("{}\n{replacement}", existing.trim_end())
    }
}
fn object_index(objects: &[Value]) -> String {
    let mut lines = vec![
        "# Object index".to_owned(),
        "".to_owned(),
        "| Object | Kind | Implementation | Documentation | Tests |".to_owned(),
        "|---|---|---|---|---|".to_owned(),
    ];
    for object in objects {
        let name = object["name"].as_str().unwrap();
        let link = if model::priority_object(name)
            && object["kind"].as_str() == Some("dispatch-interface")
        {
            format!("[{}](../objects/{}.md)", name, model::slug(name))
        } else {
            name.to_owned()
        };
        lines.push(format!(
            "| {link} | {} | {} | {} | {} |",
            object["kind"].as_str().unwrap_or("--"),
            title(&object["implemented_status"]),
            title(&object["documentation_status"]),
            title(&object["test_status"])
        ));
    }
    lines.join("\n") + "\n"
}
fn member_index_root(objects: &[Value], property: bool) -> String {
    let kind = if property { "properties" } else { "methods" };
    let title = if property { "Property" } else { "Method" };
    let mut lines = vec![
        format!("# {title} index"),
        "".to_owned(),
        "The complete inventory is sharded alphabetically by owning object; the core wrapper objects also have focused shards.".to_owned(),
        "".to_owned(),
        "## Core wrapper objects".to_owned(),
        "".to_owned(),
    ];
    for object in priority_records(objects) {
        let name = object["name"].as_str().unwrap();
        lines.push(format!("- [{name}]({kind}/{}.md)", model::slug(name)));
    }
    lines.extend([
        "".to_owned(),
        "## Complete alphabetical inventory".to_owned(),
        "".to_owned(),
    ]);
    let links = ('a'..='z')
        .map(|initial| format!("[{initial}]({kind}/all-{initial}.md)"))
        .collect::<Vec<_>>()
        .join(" · ");
    lines.push(links);
    lines.join("\n") + "\n"
}

fn member_index_for_objects(objects: &[&Value], property: bool, heading: &str) -> String {
    let mut lines = vec![
        format!("# {heading}"),
        "".to_owned(),
        "| Object | Member | DISPID | Status |".to_owned(),
        "|---|---|---:|---|".to_owned(),
    ];
    for object in objects {
        for member in object["members"]
            .as_array()
            .into_iter()
            .flatten()
            .filter(|member| {
                member["kind"].as_str() != Some("event") && is_property(member) == property
            })
        {
            lines.push(format!(
                "| {} | {} | {} | {} |",
                object["name"].as_str().unwrap_or("--"),
                member["name"].as_str().unwrap_or("--"),
                member["dispid"],
                title(&member["implementation_status"])
            ));
        }
    }
    lines.join("\n") + "\n"
}

fn index_initial(object: &Value) -> char {
    object["name"]
        .as_str()
        .and_then(|name| model::slug(name).chars().next())
        .filter(char::is_ascii_lowercase)
        .unwrap_or('z')
}
fn event_index(objects: &[Value]) -> String {
    let mut lines = vec![
        "# Event index".to_owned(),
        "".to_owned(),
        "| Event interface | Event | DISPID | Implementation |".to_owned(),
        "|---|---|---:|---|".to_owned(),
    ];
    for object in objects.iter().filter(|object| {
        object["name"]
            .as_str()
            .is_some_and(|name| name.contains("Events"))
    }) {
        for member in object["members"].as_array().into_iter().flatten() {
            lines.push(format!(
                "| {} | {} | {} | Not started |",
                object["name"].as_str().unwrap_or("--"),
                member["name"].as_str().unwrap_or("--"),
                member["dispid"]
            ));
        }
    }
    lines.join("\n") + "\n"
}
fn collection_index(objects: &[Value]) -> String {
    let mut lines = vec![
        "# Collection inventory".to_owned(),
        "".to_owned(),
        "Collections are detected structurally from Count and Item. Iterator status is independent from the broader wrapper status.".to_owned(),
        "".to_owned(),
        "## Priority collection status".to_owned(),
        "".to_owned(),
        "| Collection | Iterator status |".to_owned(),
        "|---|---|".to_owned(),
    ];
    for name in [
        "Workbooks",
        "Worksheets",
        "Areas",
        "Names",
        "Borders",
        "Charts",
        "Shapes",
        "ListObjects",
        "FormatConditions",
        "ColorScaleCriteria",
        "IconCriteria",
        "Styles",
        "Comments",
        "CommentsThreaded",
        "Hyperlinks",
    ] {
        let status = objects
            .iter()
            .find(|object| object["name"].as_str() == Some(name))
            .and_then(|object| object["collection"]["iterator_status"].as_str())
            .unwrap_or("not-started");
        lines.push(format!("| {name} | {status} |"));
    }
    lines.extend([
        "".to_owned(),
        "## All structurally identified collections".to_owned(),
        "".to_owned(),
        "| Collection | Element | Count member | Item member | Enumerator | Index kinds | Iterator |".to_owned(),
        "|---|---|---|---|---|---|---|".to_owned(),
    ]);
    for object in objects
        .iter()
        .filter(|object| object["collection"].is_object())
    {
        let collection = &object["collection"];
        lines.push(format!(
            "| {} | {} | {} | {} | {} | {} | {} |",
            object["name"].as_str().unwrap_or("--"),
            collection["element_type"].as_str().unwrap_or("Unknown"),
            collection["count_member_id"].as_str().unwrap_or("--"),
            collection["item_member_id"].as_str().unwrap_or("--"),
            collection["enumerator_member_id"].as_str().unwrap_or("--"),
            collection["index_kinds"]
                .as_array()
                .into_iter()
                .flatten()
                .filter_map(Value::as_str)
                .collect::<Vec<_>>()
                .join(", "),
            collection["iterator_status"].as_str().unwrap_or("--"),
        ));
    }
    lines.extend([
        "".to_owned(),
        "## Prompt 16 typed collection policy".to_owned(),
        "".to_owned(),
        "| Collection | Element | Index kinds | Iterator | Heterogeneous policy |".to_owned(),
        "|---|---|---|---|---|".to_owned(),
    ]);
    for object in objects
        .iter()
        .filter(|object| object["collection"]["heterogeneous_policy"].is_string())
    {
        let collection = &object["collection"];
        lines.push(format!(
            "| {} | {} | {} | {} | {} |",
            object["name"].as_str().unwrap_or("--"),
            collection["element_type"].as_str().unwrap_or("Unknown"),
            collection["index_kinds"]
                .as_array()
                .into_iter()
                .flatten()
                .filter_map(Value::as_str)
                .collect::<Vec<_>>()
                .join(", "),
            collection["iterator_status"].as_str().unwrap_or("--"),
            collection["heterogeneous_policy"].as_str().unwrap_or("--"),
        ));
    }
    lines.join("\n") + "\n"
}
fn enum_index(enums: &[Value]) -> String {
    let mut lines = vec![
        "# Enum index".to_owned(),
        "".to_owned(),
        "| Enum | Variants | Documentation | Crate status |".to_owned(),
        "|---|---:|---|---|".to_owned(),
    ];
    for enumeration in enums {
        let name = enumeration["name"].as_str().unwrap();
        lines.push(format!(
            "| [{}](../enums/{}.md) | {} | Generated | Not started |",
            name,
            model::slug(name),
            enumeration["variants"].as_array().map_or(0, Vec::len)
        ));
    }
    lines.join("\n") + "\n"
}
fn enum_page(enumeration: &Value) -> String {
    let mut lines = vec![
        format!("# {}", enumeration["name"].as_str().unwrap()),
        "".to_owned(),
        "| Variant | Numeric value | Documentation | Crate status |".to_owned(),
        "|---|---:|---|---|".to_owned(),
    ];
    for variant in enumeration["variants"].as_array().into_iter().flatten() {
        lines.push(format!(
            "| {} | {} | Generated | Not started |",
            variant["name"].as_str().unwrap_or("--"),
            variant["numeric_value"]
        ));
    }
    lines.join("\n") + "\n"
}
fn status(objects: &[Value]) -> String {
    let mut object_counts = BTreeMap::new();
    let mut member_counts = BTreeMap::new();
    let mut test_counts = BTreeMap::new();
    let mut surface_counts = BTreeMap::new();
    let mut roadmap_counts = BTreeMap::new();
    let mut raw_member_count = 0usize;
    let mut declared_member_count = 0usize;
    let mut inherited_member_count = 0usize;
    let mut implemented_declared_count = 0usize;
    for object in objects {
        *object_counts
            .entry(object["implemented_status"].as_str().unwrap_or("unknown"))
            .or_insert(0usize) += 1;
        *surface_counts
            .entry(object["surface_class"].as_str().unwrap_or("unknown"))
            .or_insert(0usize) += 1;
        *roadmap_counts
            .entry(object["roadmap_class"].as_str().unwrap_or("unknown"))
            .or_insert(0usize) += 1;
        *test_counts
            .entry(object["test_status"].as_str().unwrap_or("unknown"))
            .or_insert(0usize) += 1;
        for member in object["members"].as_array().into_iter().flatten() {
            raw_member_count += 1;
            if member["member_origin"].as_str() != Some("declared") {
                inherited_member_count += 1;
                continue;
            }
            declared_member_count += 1;
            if member["implementation_status"].as_str() == Some("implemented") {
                implemented_declared_count += 1;
            }
            let kind = if member["kind"].as_str() == Some("event") {
                "Events"
            } else if is_property(member) {
                "Properties"
            } else {
                "Methods"
            };
            *member_counts
                .entry((
                    kind,
                    member["implementation_status"]
                        .as_str()
                        .unwrap_or("unknown"),
                ))
                .or_insert(0usize) += 1;
            *test_counts
                .entry(member["test_status"].as_str().unwrap_or("unknown"))
                .or_insert(0usize) += 1;
        }
    }
    let mut text =
        "# Excel object-model status\n\n## Object coverage\n\n| Status | Count |\n|---|---:|\n"
            .to_owned();
    for status in [
        "implemented",
        "partial",
        "metadata-only",
        "not-started",
        "blocked",
        "unsupported",
    ] {
        text.push_str(&format!(
            "| {} | {} |\n",
            title_string(status),
            object_counts.get(status).unwrap_or(&0)
        ));
    }
    text.push_str(
        "\n## Type-library surface classes\n\n| Surface class | Object count |\n|---|---:|\n",
    );
    for class in model::SURFACE_CLASSES {
        text.push_str(&format!(
            "| {class} | {} |\n",
            surface_counts.get(class).unwrap_or(&0)
        ));
    }
    text.push_str("\n## Roadmap classes\n\n| Roadmap class | Object count |\n|---|---:|\n");
    for class in model::ROADMAP_CLASSES {
        text.push_str(&format!(
            "| {class} | {} |\n",
            roadmap_counts.get(class).unwrap_or(&0)
        ));
    }
    text.push_str("\n## Member coverage\n\n| Coverage basis | Count |\n|---|---:|\n");
    text.push_str(&format!(
        "| Raw type-library members | {raw_member_count} |\n| Declared Excel members | {declared_member_count} |\n| Inherited COM members | {inherited_member_count} |\n| Implemented declared Excel members | {implemented_declared_count} |\n"
    ));
    text.push_str("\nOnly declared Excel members are included in the human coverage table below.\n\n| Member type | Declared total | Implemented | Partial | Metadata only | Not started |\n|---|---:|---:|---:|---:|---:|\n");
    for kind in ["Properties", "Methods", "Events"] {
        let total: usize = member_counts
            .iter()
            .filter(|((member_kind, _), _)| *member_kind == kind)
            .map(|(_, count)| *count)
            .sum();
        text.push_str(&format!(
            "| {kind} | {total} | {} | {} | {} | {} |\n",
            member_counts.get(&(kind, "implemented")).unwrap_or(&0),
            member_counts.get(&(kind, "partial")).unwrap_or(&0),
            member_counts.get(&(kind, "metadata-only")).unwrap_or(&0),
            member_counts.get(&(kind, "not-started")).unwrap_or(&0)
        ));
    }
    text.push_str("\n## Test coverage\n\n| Test status | Count |\n|---|---:|\n");
    for status in [
        "live-tested",
        "integration-tested",
        "unit-tested",
        "not-tested",
        "blocked",
    ] {
        text.push_str(&format!(
            "| {} | {} |\n",
            title_string(status),
            test_counts.get(status).unwrap_or(&0)
        ));
    }
    text.push_str("\n## Priority objects\n\n| Object | Status |\n|---|---|\n");
    for object in priority_records(objects) {
        let name = object["name"].as_str().unwrap();
        text.push_str(&format!(
            "| [{}](objects/{}.md) | {} |\n",
            name,
            model::slug(name),
            title(&object["implemented_status"])
        ));
    }
    text
}
fn priority_records(objects: &[Value]) -> Vec<&Value> {
    [
        "Application",
        "Workbooks",
        "Workbook",
        "Worksheets",
        "Worksheet",
        "Range",
        "Areas",
        "Names",
        "Name",
        "Font",
        "Interior",
        "Borders",
        "Border",
        "ListObjects",
        "ListObject",
        "ListColumns",
        "ListColumn",
        "ListRows",
        "ListRow",
        "AutoFilter",
        "Filters",
        "Filter",
        "Sort",
        "SortFields",
        "SortField",
        "Validation",
        "Sheets",
        "Windows",
        "Window",
        "PageSetup",
        "Tab",
        "Outline",
        "HPageBreaks",
        "HPageBreak",
        "VPageBreaks",
        "VPageBreak",
    ]
    .into_iter()
    .filter_map(|name| {
        objects.iter().find(|object| {
            object["name"].as_str() == Some(name)
                && object["kind"].as_str() == Some("dispatch-interface")
        })
    })
    .collect()
}
fn title(value: &Value) -> String {
    value
        .as_str()
        .map(title_string)
        .unwrap_or_else(|| "--".to_owned())
}
fn title_string(value: &str) -> String {
    value
        .split('-')
        .map(|word| {
            let mut chars = word.chars();
            chars
                .next()
                .map(|first| first.to_uppercase().collect::<String>() + chars.as_str())
                .unwrap_or_default()
        })
        .collect::<Vec<_>>()
        .join(" ")
}
fn read_records(directory: &Path) -> Result<Vec<Value>, String> {
    let mut result: Vec<Value> = Vec::new();
    for entry in fs::read_dir(directory).map_err(|error| error.to_string())? {
        let path = entry.map_err(|error| error.to_string())?.path();
        if path.extension().and_then(|value| value.to_str()) == Some("json") {
            result.push(
                serde_json::from_str(&fs::read_to_string(path).map_err(|error| error.to_string())?)
                    .map_err(|error| error.to_string())?,
            );
        }
    }
    result.sort_by_key(|value| value["id"].as_str().unwrap_or_default().to_owned());
    Ok(result)
}
fn read_relationships(path: &Path) -> Result<Vec<Value>, String> {
    let mut relationships: Vec<Value> =
        serde_json::from_str(&fs::read_to_string(path).map_err(|error| error.to_string())?)
            .map_err(|error| error.to_string())?;
    relationships.sort_by_key(|value| {
        (
            value["source"].as_str().unwrap_or_default().to_owned(),
            value["member_id"].as_str().unwrap_or_default().to_owned(),
        )
    });
    Ok(relationships)
}
fn write(path: &Path, content: &str) -> Result<(), String> {
    if let Some(parent) = path.parent() {
        fs::create_dir_all(parent).map_err(|error| error.to_string())?;
    }
    fs::write(path, format!("{}\n", content.trim_end())).map_err(|error| error.to_string())
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn generated_region_preserves_manual_sections() {
        let input = "# Example\n\nManual note.\n\n<!-- BEGIN GENERATED MEMBERS -->\nold\n<!-- END GENERATED MEMBERS -->\n\nTail.\n";
        let output = replace_region(input, "new");
        assert!(output.contains("Manual note."));
        assert!(output.contains("Tail."));
        assert!(output.contains("new"));
        assert!(!output.contains("old"));
    }
}
