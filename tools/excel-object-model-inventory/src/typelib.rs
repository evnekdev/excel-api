use std::collections::{BTreeMap, BTreeSet};
use std::fs;
use std::path::Path;

use serde_json::{Value, json};
use windows::Win32::System::Com::{
    COINIT_APARTMENTTHREADED, CoInitializeEx, CoUninitialize, FUNCDESC, ITypeInfo, TKIND_ALIAS,
    TKIND_COCLASS, TKIND_DISPATCH, TKIND_ENUM, TKIND_INTERFACE, TYPEATTR, TYPEDESC, VARDESC,
};
use windows::Win32::System::Ole::{LoadRegTypeLib, PARAMFLAG_FOPT};
use windows::Win32::System::Variant::{
    VARIANT, VT_BOOL, VT_BSTR, VT_CY, VT_DATE, VT_DISPATCH, VT_ERROR, VT_I1, VT_I2, VT_I4, VT_I8,
    VT_INT, VT_PTR, VT_R4, VT_R8, VT_UI1, VT_UI2, VT_UI4, VT_UI8, VT_UINT, VT_USERDEFINED,
    VT_VARIANT,
};
use windows::core::{BSTR, GUID};

use crate::model::{self, SCHEMA_VERSION};

const EXCEL_TYPELIB: GUID = GUID::from_u128(0x00020813_0000_0000_c000_000000000046);

pub struct Summary {
    pub type_infos: usize,
    pub objects: usize,
    pub members: usize,
    pub enums: usize,
}
struct Apartment;
impl Apartment {
    fn sta() -> Result<Self, String> {
        // SAFETY: null reserved pointer and apartment flag are documented valid inputs.
        let status = unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED) };
        if status.is_ok() {
            Ok(Self)
        } else {
            Err(format!("CoInitializeEx: {status}"))
        }
    }
}
impl Drop for Apartment {
    fn drop(&mut self) {
        // SAFETY: this guard exists only after successful CoInitializeEx.
        unsafe { CoUninitialize() }
    }
}
struct Attr<'a> {
    info: &'a ITypeInfo,
    raw: *mut TYPEATTR,
}
impl<'a> Attr<'a> {
    fn get(info: &'a ITypeInfo) -> Result<Self, String> {
        // SAFETY: ITypeInfo is valid for this lexical borrow and releases the descriptor in Drop.
        let raw = unsafe { info.GetTypeAttr() }.map_err(|error| error.to_string())?;
        (!raw.is_null())
            .then_some(Self { info, raw })
            .ok_or("ITypeInfo returned null TYPEATTR".to_owned())
    }
}
impl Drop for Attr<'_> {
    fn drop(&mut self) {
        // SAFETY: `raw` was acquired from this exact ITypeInfo and is released exactly once.
        unsafe { self.info.ReleaseTypeAttr(self.raw) }
    }
}
struct Function<'a> {
    info: &'a ITypeInfo,
    raw: *mut FUNCDESC,
}
impl<'a> Function<'a> {
    fn get(info: &'a ITypeInfo, index: u32) -> Result<Self, String> {
        // SAFETY: `index` is bounded by TYPEATTR.cFuncs and the guard releases the descriptor.
        let raw = unsafe { info.GetFuncDesc(index) }.map_err(|error| error.to_string())?;
        (!raw.is_null())
            .then_some(Self { info, raw })
            .ok_or("ITypeInfo returned null FUNCDESC".to_owned())
    }
}
impl Drop for Function<'_> {
    fn drop(&mut self) {
        // SAFETY: `raw` was acquired from this exact ITypeInfo and is released exactly once.
        unsafe { self.info.ReleaseFuncDesc(self.raw) }
    }
}
struct Variable<'a> {
    info: &'a ITypeInfo,
    raw: *mut VARDESC,
}
impl<'a> Variable<'a> {
    fn get(info: &'a ITypeInfo, index: u32) -> Result<Self, String> {
        // SAFETY: `index` is bounded by TYPEATTR.cVars and the guard releases the descriptor.
        let raw = unsafe { info.GetVarDesc(index) }.map_err(|error| error.to_string())?;
        (!raw.is_null())
            .then_some(Self { info, raw })
            .ok_or("ITypeInfo returned null VARDESC".to_owned())
    }
}
impl Drop for Variable<'_> {
    fn drop(&mut self) {
        // SAFETY: `raw` was acquired from this exact ITypeInfo and is released exactly once.
        unsafe { self.info.ReleaseVarDesc(self.raw) }
    }
}

pub fn extract(root: &Path) -> Result<Summary, String> {
    let _apartment = Apartment::sta()?;
    // SAFETY: Excel's documented typelib GUID/version and a zero LCID are read-only inputs.
    let library = unsafe { LoadRegTypeLib(&EXCEL_TYPELIB, 1, 9, 0) }
        .map_err(|error| format!("registered Excel 1.9 typelib unavailable: {error}"))?;
    // SAFETY: the library is valid and its TLIBATTR is released immediately below.
    let attribute = unsafe { library.GetLibAttr() }.map_err(|error| error.to_string())?;
    // SAFETY: `attribute` remains valid until ReleaseTLibAttr below.
    let attribute_value = unsafe { *attribute };
    // SAFETY: the pointer was acquired from this exact type library.
    unsafe { library.ReleaseTLibAttr(attribute) };
    // SAFETY: querying the count is read-only on a valid type library.
    let count = unsafe { library.GetTypeInfoCount() };
    let mut objects = Vec::new();
    let mut enums = Vec::new();
    let mut relationships = BTreeSet::new();
    let implemented = model::implemented_member_ids();
    for index in 0..count {
        // SAFETY: `index` is bounded by GetTypeInfoCount above.
        let info = unsafe { library.GetTypeInfo(index) }.map_err(|error| error.to_string())?;
        let attr = Attr::get(&info)?;
        // SAFETY: Attr retains this TYPEATTR until the end of the loop body.
        let value = unsafe { *attr.raw };
        let name = documentation(&info, -1).0;
        let kind = typekind(value.typekind);
        if matches!(
            value.typekind,
            TKIND_DISPATCH | TKIND_INTERFACE | TKIND_COCLASS | TKIND_ALIAS
        ) {
            let record =
                object_record(&info, &name, kind, &value, &implemented, &mut relationships)?;
            objects.push(record);
        }
        if value.typekind == TKIND_ENUM {
            enums.push(enum_record(&info, &name, &value)?);
        }
    }
    objects.sort_by_key(|value| value["id"].as_str().unwrap_or_default().to_owned());
    enums.sort_by_key(|value| value["id"].as_str().unwrap_or_default().to_owned());
    let metadata = root.join("metadata/excel-object-model");
    fs::create_dir_all(metadata.join("objects")).map_err(|error| error.to_string())?;
    fs::create_dir_all(metadata.join("enums")).map_err(|error| error.to_string())?;
    for directory in [metadata.join("objects"), metadata.join("enums")] {
        for entry in fs::read_dir(&directory).map_err(|error| error.to_string())? {
            let path = entry.map_err(|error| error.to_string())?.path();
            if path.extension().and_then(|extension| extension.to_str()) == Some("json") {
                fs::remove_file(path).map_err(|error| error.to_string())?;
            }
        }
    }
    for object in &objects {
        write_json(
            &metadata.join("objects").join(format!(
                "{}.json",
                object["id"]
                    .as_str()
                    .unwrap()
                    .trim_start_matches("excel.")
                    .replace('.', "-")
            )),
            object,
        )?;
    }
    for enumeration in &enums {
        write_json(
            &metadata.join("enums").join(format!(
                "{}.json",
                enumeration["id"]
                    .as_str()
                    .unwrap()
                    .trim_start_matches("excel.enum.")
                    .replace('.', "-")
            )),
            enumeration,
        )?;
    }
    write_json(
        &metadata.join("relationships.json"),
        &Value::Array(
            relationships
                .into_iter()
                .map(|value| serde_json::from_str(&value).unwrap())
                .collect(),
        ),
    )?;
    write_text(&metadata.join("aliases.json"), &serde_json::to_string_pretty(&json!({"schema_version": SCHEMA_VERSION, "aliases": {"_Application":"Application", "_Workbook":"Workbook", "_Worksheet":"Worksheet"}})).map_err(|error| error.to_string())?)?;
    write_text(
        &metadata.join("overrides.toml"),
        "# Manual overrides require reason, source, and date.\n# No overrides are needed for the initial extracted inventory.\n",
    )?;
    write_text(
        &metadata.join("manifest.toml"),
        &format!(
            "schema_version = 1\ninventory_tool_version = \"0.1.0\"\nsource = \"registered-typelib\"\ntypelib_guid = \"{}\"\ntypelib_major = {}\ntypelib_minor = {}\nlcid = {}\nexcel_version = \"16.0\"\noffice_bitness = \"64-bit\"\nextraction_date = \"2026-07-22\"\n",
            guid(&attribute_value.guid),
            attribute_value.wMajorVerNum,
            attribute_value.wMinorVerNum,
            attribute_value.lcid
        ),
    )?;
    let summary = Summary {
        type_infos: count as usize,
        objects: objects.len(),
        members: objects
            .iter()
            .map(|object| object["members"].as_array().map_or(0, Vec::len))
            .sum(),
        enums: enums.len(),
    };
    write_json(
        &metadata.join("baseline.json"),
        &json!({
            "schema_version": SCHEMA_VERSION,
            "type_infos": summary.type_infos,
            "objects": summary.objects,
            "members": summary.members,
            "enums": summary.enums,
            "source": "registered-typelib structural baseline",
        }),
    )?;
    Ok(summary)
}

fn object_record(
    info: &ITypeInfo,
    raw_name: &str,
    kind: &str,
    attr: &TYPEATTR,
    implemented: &BTreeSet<&str>,
    relationships: &mut BTreeSet<String>,
) -> Result<Value, String> {
    let event_interface = raw_name.contains("Events") || raw_name == "DocEvents";
    let mut members: BTreeMap<(String, i32), Value> = BTreeMap::new();
    for index in 0..u32::from(attr.cFuncs) {
        let function = Function::get(info, index)?;
        // SAFETY: Function retains this FUNCDESC until the end of the iteration.
        let value = unsafe { *function.raw };
        let (name, _) = documentation(info, value.memid);
        let id = model::member_id(raw_name, &name);
        let entry = members.entry((name.clone(), value.memid)).or_insert_with(|| json!({"schema_version": SCHEMA_VERSION, "id": id, "name": name, "kind": member_kind(value.invkind), "member_origin": model::member_origin(&name), "dispid": value.memid, "invoke_kinds": [], "arguments": [], "optional_argument_count": 0, "return_type": normalized_type(info, &value.elemdescFunc.tdesc), "raw_return_type": raw_type(info, &value.elemdescFunc.tdesc), "notes": []}));
        let kind = member_kind(value.invkind).to_owned();
        if !entry["invoke_kinds"]
            .as_array()
            .unwrap()
            .iter()
            .any(|candidate| candidate == &Value::String(kind.clone()))
        {
            entry["invoke_kinds"]
                .as_array_mut()
                .unwrap()
                .push(Value::String(kind));
        }
        let params = if value.cParams == 0 {
            &[]
        } else {
            if value.lprgelemdescParam.is_null() {
                return Err(format!(
                    "{} has {} parameters but a null ELEMDESC pointer",
                    name, value.cParams
                ));
            }
            // SAFETY: cParams is nonzero and the null case is rejected immediately above.
            unsafe { std::slice::from_raw_parts(value.lprgelemdescParam, value.cParams as usize) }
        };
        let arguments = entry["arguments"].as_array_mut().unwrap();
        if arguments.is_empty() {
            for (position, parameter) in params.iter().enumerate() {
                // SAFETY: FUNCDESC parameter elements are ELEMDESC parameter descriptors.
                let flags = unsafe { parameter.Anonymous.paramdesc.wParamFlags.0 };
                arguments.push(json!({"position": position, "raw_com_type": raw_type(info, &parameter.tdesc), "normalized_type": normalized_type(info, &parameter.tdesc), "direction": direction(flags), "optional": flags & PARAMFLAG_FOPT.0 != 0, "default": null}));
            }
        }
        entry["optional_argument_count"] = Value::from(
            arguments
                .iter()
                .filter(|argument| argument["optional"].as_bool() == Some(true))
                .count(),
        );
    }
    let mut members: Vec<Value> = members.into_values().collect();
    members.sort_by_key(|member| {
        (
            member["dispid"].as_i64().unwrap_or_default(),
            member["name"].as_str().unwrap_or_default().to_owned(),
        )
    });
    let mut member_ids = BTreeSet::new();
    for member in &mut members {
        let base = member["id"].as_str().unwrap().to_owned();
        if !member_ids.insert(base.clone()) {
            member["id"] = Value::String(format!(
                "{}-{}",
                base,
                member["dispid"].as_i64().unwrap_or_default()
            ));
            member_ids.insert(member["id"].as_str().unwrap().to_owned());
        }
    }
    for member in &mut members {
        let id = member["id"].as_str().unwrap().to_owned();
        let target = member["return_type"].as_str().map(str::to_owned);
        let selected = !event_interface && implemented.contains(id.as_str());
        if event_interface {
            member["kind"] = Value::String("event".to_owned());
        }
        member["typelib_invoke_kinds"] = member["invoke_kinds"].clone();
        member["runtime_confirmed_invoke_kinds"] = Value::Array(if selected {
            crate_policy(&id)
                .iter()
                .map(|value| Value::String((*value).to_owned()))
                .collect()
        } else {
            Vec::new()
        });
        member["microsoft_sample_invoke_kinds"] = Value::Array(if id == "excel.workbooks.add" {
            vec![Value::String("PROPERTYGET".to_owned())]
        } else {
            Vec::new()
        });
        member["chosen_crate_invoke_kinds"] = member["runtime_confirmed_invoke_kinds"].clone();
        member["source"] = json!({"typelib":true,"microsoft_docs":model::documentation_url(raw_name).is_some(),"runtime_verified":selected});
        member["implementation_status"] = Value::String(
            if event_interface {
                "not-started"
            } else if selected {
                "implemented"
            } else {
                "metadata-only"
            }
            .to_owned(),
        );
        member["documentation_status"] = Value::String(
            if model::priority_object(raw_name) {
                "reviewed"
            } else {
                "generated"
            }
            .to_owned(),
        );
        member["test_status"] = Value::String(
            if selected {
                "live-tested"
            } else {
                "not-tested"
            }
            .to_owned(),
        );
        member["source_confidence"] = Value::String(
            if selected {
                "runtime-confirmed"
            } else {
                "typelib-only"
            }
            .to_owned(),
        );
        if mixed_value_possible(&id) {
            member["mixed_value_possible"] = Value::Bool(true);
        }
        if version_sensitive(&id) {
            member["version_sensitive"] = Value::Bool(true);
        }
        if returns_range_or_nothing(&id) {
            member["returns_range_or_nothing"] = Value::Bool(true);
        }
        if stateful_search(&id) {
            member["stateful_search"] = Value::Bool(true);
        }
        if returns_optional_dispatch(&id) {
            member["returns_optional_dispatch"] = Value::Bool(true);
        }
        if one_based_field(&id) {
            member["one_based_field"] = Value::Bool(true);
        }
        if modifies_range_in_place(&id) {
            member["modifies_range_in_place"] = Value::Bool(true);
        }
        if stateful_filter(&id) {
            member["stateful_filter"] = Value::Bool(true);
        }
        if clipboard_dependent(&id) {
            member["clipboard_dependent"] = Value::Bool(true);
        }
        let target = target
            .filter(|target| {
                matches!(
                    target.as_str(),
                    "Application"
                        | "Workbooks"
                        | "Workbook"
                        | "Worksheets"
                        | "Worksheet"
                        | "Range"
                        | "Areas"
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
                )
            })
            .or_else(|| relationship_fallback(&id).map(str::to_owned));
        if let Some(target) = target {
            relationships.insert(serde_json::to_string(&json!({"schema_version":SCHEMA_VERSION,"source":model::object_id(raw_name),"member_id":id,"target":format!("excel.{}", model::slug(&target)),"source_confidence":"typelib-only"})).unwrap());
        }
    }
    let object_id = match kind {
        "coclass" => format!("{}.coclass", model::object_id(raw_name)),
        "alias" => format!("{}.alias", model::object_id(raw_name)),
        _ => model::object_id(raw_name),
    };
    let is_collection = members
        .iter()
        .any(|member| member["name"].as_str() == Some("Count"))
        && members
            .iter()
            .any(|member| member["name"].as_str() == Some("Item"));
    let collection_element = match model::canonical_name(raw_name) {
        "Workbooks" => Some("Workbook"),
        "Worksheets" => Some("Worksheet"),
        "Areas" => Some("Range"),
        "Names" => Some("Name"),
        "Borders" => Some("Border"),
        "ListObjects" => Some("ListObject"),
        "ListColumns" => Some("ListColumn"),
        "ListRows" => Some("ListRow"),
        "Filters" => Some("Filter"),
        "SortFields" => Some("SortField"),
        _ => None,
    };
    let member_id = |name: &str| {
        members
            .iter()
            .find(|member| member["name"].as_str() == Some(name))
            .and_then(|member| member["id"].as_str())
            .map(str::to_owned)
    };
    let index_kinds = match model::canonical_name(raw_name) {
        "Workbooks" | "Worksheets" | "Names" | "ListObjects" | "ListColumns" => {
            vec!["one-based-integer", "string-key"]
        }
        "Areas" | "ListRows" | "Filters" | "SortFields" => vec!["one-based-integer"],
        "Borders" => vec!["enum-key"],
        _ if is_collection => vec!["variant-key"],
        _ => vec!["no-index"],
    };
    let iterator_status = match model::canonical_name(raw_name) {
        "Workbooks" | "Worksheets" | "Areas" | "Names" | "Borders" | "ListObjects"
        | "ListColumns" | "ListRows" | "Filters" => "implemented",
        _ if is_collection => "metadata-only",
        _ => "not-started",
    };
    let collection = is_collection.then(|| {
        json!({
            "element_type": collection_element.unwrap_or("Unknown"),
            "count_member_id": member_id("Count"),
            "item_member_id": member_id("Item"),
            "enumerator_member_id": member_id("_NewEnum"),
            "index_kinds": index_kinds,
            "iterator_status": iterator_status,
        })
    });
    if let Some(element) = collection_element {
        if let Some(item) = members
            .iter()
            .find(|member| member["name"].as_str() == Some("Item"))
        {
            relationships.insert(serde_json::to_string(&json!({"schema_version":SCHEMA_VERSION,"source":model::object_id(raw_name),"member_id":item["id"].as_str().unwrap(),"target":format!("excel.{}", model::slug(element)),"source_confidence":"typelib-only"})).unwrap());
        }
    }
    let wrapper_object =
        !event_interface && model::wrapper_object(raw_name) && kind == "dispatch-interface";
    let priority_documentation = model::priority_object(raw_name) && kind == "dispatch-interface";
    let alias_target = if kind == "alias" {
        Some(normalized_type(info, &attr.tdescAlias))
    } else {
        None
    };
    let mut record = json!({"schema_version": SCHEMA_VERSION, "id": object_id, "name": model::canonical_name(raw_name), "kind": if event_interface { "event-interface" } else { kind }, "surface_class": model::surface_class(raw_name, kind, event_interface, attr.wTypeFlags), "roadmap_class": model::roadmap_class(raw_name, kind, event_interface), "typelib_type_flags": attr.wTypeFlags, "guid": guid(&attr.guid), "source_interface": raw_name, "typelib_version": {"major":1,"minor":9}, "documentation_url": model::documentation_url(raw_name), "implemented_status": if event_interface { "not-started" } else if wrapper_object { "partial" } else { "metadata-only" }, "documentation_status": if priority_documentation { "reviewed" } else { "generated" }, "test_status": if wrapper_object { "live-tested" } else { "not-tested" }, "implemented_interface_count": attr.cImplTypes, "alias_target": alias_target, "collection": collection, "members": members});
    let capabilities = reference_capabilities(raw_name);
    if !capabilities.is_null() {
        record["reference_capabilities"] = capabilities;
    }
    let categories = evaluation_result_categories(raw_name);
    if !categories.is_null() {
        record["evaluation_result_categories"] = categories;
    }
    let formatting = formatting_capability(raw_name);
    if !formatting.is_null() {
        record["formatting_capability"] = formatting;
    }
    let formulas = formula_capability(raw_name);
    if !formulas.is_null() {
        record["formula_capability"] = formulas;
    }
    let calculation = calculation_capability(raw_name);
    if !calculation.is_null() {
        record["calculation_capability"] = calculation;
    }
    let auditing_search = auditing_search_capability(raw_name);
    if !auditing_search.is_null() {
        record["auditing_search_capability"] = auditing_search;
    }
    let structured_data = structured_data_capability(raw_name);
    if !structured_data.is_null() {
        record["structured_data_capability"] = structured_data;
    }
    Ok(record)
}

fn enum_record(info: &ITypeInfo, name: &str, attr: &TYPEATTR) -> Result<Value, String> {
    let mut variants = Vec::new();
    for index in 0..u32::from(attr.cVars) {
        let variable = Variable::get(info, index)?;
        // SAFETY: Variable retains this VARDESC until the end of the iteration.
        let value = unsafe { *variable.raw };
        let (variant, _) = documentation(info, value.memid);
        // SAFETY: enum VARDESC stores its declared constant in lpvarValue when present.
        let numeric_value = unsafe { value.Anonymous.lpvarValue.as_ref() }
            .map(enum_numeric)
            .unwrap_or(Value::Null);
        variants.push(json!({"name":variant,"numeric_value":numeric_value,"dispid":value.memid,"documentation_status":"generated","implementation_status":"not-started"}));
    }
    variants.sort_by_key(|value| value["name"].as_str().unwrap_or_default().to_owned());
    Ok(
        json!({"schema_version":SCHEMA_VERSION,"id":format!("excel.enum.{}",model::slug(name)),"name":name,"guid":guid(&attr.guid),"variants":variants,"documentation_status":"generated","implementation_status":"not-started"}),
    )
}

fn enum_numeric(value: &VARIANT) -> Value {
    // SAFETY: callers pass a VARIANT pointer owned by the typelib descriptor for inspection only.
    let inner = unsafe {
        &*(&value.Anonymous.Anonymous as *const std::mem::ManuallyDrop<_>
            as *const windows::Win32::System::Variant::VARIANT_0_0)
    };
    let contents = &inner.Anonymous;
    // SAFETY: the active VARTYPE selects the matching scalar union field.
    unsafe {
        match inner.vt {
            VT_I1 => json!(contents.cVal),
            VT_I2 => json!(contents.iVal),
            VT_I4 | VT_INT => json!(contents.lVal),
            VT_I8 => json!(contents.llVal),
            VT_UI1 => json!(contents.bVal),
            VT_UI2 => json!(contents.uiVal),
            VT_UI4 | VT_UINT => json!(contents.ulVal),
            VT_UI8 => json!(contents.ullVal),
            _ => Value::Null,
        }
    }
}

fn crate_policy(id: &str) -> &'static [&'static str] {
    match id {
        "excel.application.visible"
        | "excel.application.displayalerts"
        | "excel.application.referencestyle"
        | "excel.application.calculation"
        | "excel.application.calculatebeforesave"
        | "excel.workbook.saved"
        | "excel.worksheet.name"
        | "excel.worksheet.visible"
        | "excel.range.value"
        | "excel.range.value2"
        | "excel.range.formula"
        | "excel.range.formula2"
        | "excel.range.formular1c1"
        | "excel.range.formula2r1c1"
        | "excel.range.formulalocal"
        | "excel.range.formular1c1local"
        | "excel.range.formulaarray"
        | "excel.range.numberformat"
        | "excel.range.horizontalalignment"
        | "excel.range.verticalalignment"
        | "excel.range.wraptext"
        | "excel.range.rowheight"
        | "excel.range.columnwidth"
        | "excel.font.name"
        | "excel.font.size"
        | "excel.font.bold"
        | "excel.font.italic"
        | "excel.font.underline"
        | "excel.font.strikethrough"
        | "excel.font.color"
        | "excel.font.colorindex"
        | "excel.interior.color"
        | "excel.interior.colorindex"
        | "excel.interior.pattern"
        | "excel.interior.patterncolor"
        | "excel.interior.patterncolorindex"
        | "excel.borders.color"
        | "excel.borders.colorindex"
        | "excel.borders.linestyle"
        | "excel.borders.weight"
        | "excel.border.color"
        | "excel.border.colorindex"
        | "excel.border.linestyle"
        | "excel.border.weight" => &["PROPERTYGET", "PROPERTYPUT"],
        "excel.application.convertformula"
        | "excel.application.evaluate-1"
        | "excel.worksheet.evaluate-1"
        | "excel.application.calculate"
        | "excel.application.calculatefull"
        | "excel.application.calculatefullrebuild"
        | "excel.worksheet.calculate"
        | "excel.range.calculate"
        | "excel.range.dirty"
        | "excel.range.directprecedents"
        | "excel.range.directdependents"
        | "excel.range.precedents"
        | "excel.range.dependents"
        | "excel.range.specialcells"
        | "excel.range.find"
        | "excel.range.findnext"
        | "excel.range.findprevious"
        | "excel.range.replace-3305"
        | "excel.names.item"
        | "excel.names.add"
        | "excel.name.delete" => &["METHOD"],
        "excel.application.quit"
        | "excel.workbook.close"
        | "excel.workbooks.open-1923"
        | "excel.workbook.save"
        | "excel.workbook.saveas-3174"
        | "excel.workbook.savecopyas" => &["METHOD"],
        "excel.workbooks.add" => &["PROPERTYGET"],
        "excel.worksheets.add" | "excel.range.clearcontents" | "excel.range.autofit" => &["METHOD"],
        "excel.application.version"
        | "excel.application.workbooks"
        | "excel.application.calculationstate"
        | "excel.application.calculationversion"
        | "excel.workbook.names"
        | "excel.workbooks.count"
        | "excel.workbook.name"
        | "excel.workbook.fullname"
        | "excel.workbook.path"
        | "excel.workbook.fileformat"
        | "excel.workbook.readonly"
        | "excel.workbook.worksheets"
        | "excel.worksheets.count"
        | "excel.worksheets.item"
        | "excel.worksheet.index"
        | "excel.worksheet.range"
        | "excel.worksheet.usedrange"
        | "excel.worksheet.application"
        | "excel.worksheet.cells"
        | "excel.worksheet.names"
        | "excel.range.address"
        | "excel.range.row"
        | "excel.range.column"
        | "excel.range.count"
        | "excel.range.rows"
        | "excel.range.columns"
        | "excel.range.hasformula"
        | "excel.range.hasarray"
        | "excel.range.currentarray"
        | "excel.range.hasspill"
        | "excel.range.spillingtorange"
        | "excel.range.spillparent"
        | "excel.range.font"
        | "excel.range.interior"
        | "excel.range.borders"
        | "excel.borders.count"
        | "excel.borders.item"
        | "excel.borders.newenum"
        | "excel.names.count"
        | "excel.names.newenum"
        | "excel.name.name"
        | "excel.name.refersto"
        | "excel.name.referstor1c1"
        | "excel.name.referstorange"
        | "excel.name.visible" => &["PROPERTYGET"],
        "excel.listobject.name"
        | "excel.listobject.displayname"
        | "excel.listobject.showheaders"
        | "excel.listobject.showtotals"
        | "excel.listobject.showautofilter"
        | "excel.listobject.tablestyle"
        | "excel.listcolumn.name"
        | "excel.listcolumn.totalscalculation"
        | "excel.sort.header"
        | "excel.sort.matchcase"
        | "excel.sort.orientation"
        | "excel.validation.ignoreblank"
        | "excel.validation.incelldropdown"
        | "excel.validation.showinput"
        | "excel.validation.showerror"
        | "excel.validation.inputtitle"
        | "excel.validation.inputmessage"
        | "excel.validation.errortitle"
        | "excel.validation.errormessage"
        | "excel.range.hidden" => &["PROPERTYGET", "PROPERTYPUT"],
        "excel.listobjects.add"
        | "excel.listobject.resize"
        | "excel.listobject.delete"
        | "excel.listobject.unlist"
        | "excel.listcolumns.add"
        | "excel.listcolumn.delete"
        | "excel.listrows.add"
        | "excel.listrow.delete"
        | "excel.worksheet.showalldata"
        | "excel.range.autofilter-3289"
        | "excel.range.sort"
        | "excel.sort.setrange"
        | "excel.sort.apply"
        | "excel.sortfields.clear"
        | "excel.sortfields.add"
        | "excel.validation.add"
        | "excel.validation.delete"
        | "excel.range.removeduplicates"
        | "excel.range.insert"
        | "excel.range.delete"
        | "excel.range.clear"
        | "excel.range.clearformats"
        | "excel.range.clearcomments"
        | "excel.range.clearhyperlinks"
        | "excel.range.copy"
        | "excel.range.cut"
        | "excel.range.pastespecial-1928" => &["METHOD"],
        "excel.worksheet.listobjects"
        | "excel.worksheet.autofiltermode"
        | "excel.worksheet.filtermode"
        | "excel.worksheet.autofilter-3289"
        | "excel.range.worksheet"
        | "excel.listobjects.count"
        | "excel.listobjects.item"
        | "excel.listobjects.newenum"
        | "excel.listobject.range"
        | "excel.listobject.headerrowrange"
        | "excel.listobject.databodyrange"
        | "excel.listobject.totalsrowrange"
        | "excel.listobject.insertrowrange"
        | "excel.listobject.listcolumns"
        | "excel.listobject.listrows"
        | "excel.listobject.autofilter-3289"
        | "excel.listobject.sort-3288"
        | "excel.listcolumns.count"
        | "excel.listcolumns.item"
        | "excel.listcolumns.newenum"
        | "excel.listcolumn.index"
        | "excel.listcolumn.range"
        | "excel.listcolumn.databodyrange"
        | "excel.listcolumn.total"
        | "excel.listrows.count"
        | "excel.listrows.item"
        | "excel.listrows.newenum"
        | "excel.listrow.index"
        | "excel.listrow.range"
        | "excel.autofilter.range"
        | "excel.autofilter.filters"
        | "excel.filters.count"
        | "excel.filters.item"
        | "excel.filters.newenum"
        | "excel.filter.on"
        | "excel.filter.operator-2641"
        | "excel.filter.criteria1"
        | "excel.filter.criteria2"
        | "excel.sort.sortfields"
        | "excel.validation.value"
        | "excel.validation.type"
        | "excel.validation.alertstyle"
        | "excel.validation.operator"
        | "excel.validation.formula1"
        | "excel.validation.formula2"
        | "excel.range.validation"
        | "excel.range.currentregion" => &["PROPERTYGET"],
        _ => &[],
    }
}

fn reference_capabilities(name: &str) -> Value {
    match model::canonical_name(name) {
        "Worksheet" => {
            json!({"input_styles":["a1", "r1c1"], "output_styles":[], "relative_address":false, "external_address":false, "formula_conversion":false})
        }
        "Range" => {
            json!({"input_styles":[], "output_styles":["a1", "r1c1"], "relative_address":true, "external_address":true, "formula_conversion":false})
        }
        "Application" => {
            json!({"input_styles":["a1", "r1c1"], "output_styles":["a1", "r1c1"], "relative_address":true, "external_address":false, "formula_conversion":true})
        }
        _ => Value::Null,
    }
}

fn evaluation_result_categories(name: &str) -> Value {
    match model::canonical_name(name) {
        "Application" | "Worksheet" => json!([
            "automation-value",
            "range-object",
            "other-object",
            "unknown"
        ]),
        _ => Value::Null,
    }
}

fn formatting_capability(name: &str) -> Value {
    match model::canonical_name(name) {
        "Range" => json!({
            "font": true,
            "fill": true,
            "borders": true,
            "number_format": true,
            "alignment": true,
            "dimensions": true,
            "autofit": true,
        }),
        _ => Value::Null,
    }
}

fn formula_capability(name: &str) -> Value {
    match model::canonical_name(name) {
        "Range" => json!({
            "a1": true,
            "r1c1": true,
            "formula2": true,
            "formula2_r1c1": true,
            "dynamic_array": true,
            "legacy_array": true,
            "locale_formula": true,
            "mixed_values": true,
        }),
        _ => Value::Null,
    }
}

fn calculation_capability(name: &str) -> Value {
    match model::canonical_name(name) {
        "Application" => json!({
            "mode": true,
            "state": true,
            "version": true,
            "before_save": true,
            "calculate": true,
            "full": true,
            "full_rebuild": true,
        }),
        "Worksheet" => json!({"calculate": true}),
        "Range" => json!({"calculate": true, "mark_dirty": true}),
        _ => Value::Null,
    }
}

fn auditing_search_capability(name: &str) -> Value {
    match model::canonical_name(name) {
        "Range" => json!({
            "precedents": true,
            "dependents": true,
            "special_cells": true,
            "find": true,
            "replace": true,
            "wrap_safe_iterator": true,
        }),
        _ => Value::Null,
    }
}

fn structured_data_capability(name: &str) -> Value {
    match model::canonical_name(name) {
        "Worksheet" | "Range" => json!({
            "tables": true,
            "sort": true,
            "filter": true,
            "validation": true,
            "remove_duplicates": true,
            "structural_editing": true,
        }),
        "ListObjects" | "ListObject" | "ListColumns" | "ListColumn" | "ListRows" | "ListRow" => {
            json!({
                "tables": true,
                "sort": true,
                "filter": true,
                "validation": false,
                "remove_duplicates": false,
                "structural_editing": true,
            })
        }
        "AutoFilter" | "Filters" | "Filter" => json!({
            "tables": false,
            "sort": false,
            "filter": true,
            "validation": false,
            "remove_duplicates": false,
            "structural_editing": false,
        }),
        "Sort" | "SortFields" | "SortField" => json!({
            "tables": false,
            "sort": true,
            "filter": false,
            "validation": false,
            "remove_duplicates": false,
            "structural_editing": false,
        }),
        "Validation" => json!({
            "tables": false,
            "sort": false,
            "filter": false,
            "validation": true,
            "remove_duplicates": false,
            "structural_editing": false,
        }),
        _ => Value::Null,
    }
}

fn returns_optional_dispatch(id: &str) -> bool {
    matches!(
        id,
        "excel.worksheet.autofilter-3289"
            | "excel.listobject.headerrowrange"
            | "excel.listobject.databodyrange"
            | "excel.listobject.totalsrowrange"
            | "excel.listobject.insertrowrange"
            | "excel.listcolumn.databodyrange"
            | "excel.listcolumn.total"
    )
}

fn one_based_field(id: &str) -> bool {
    matches!(
        id,
        "excel.listobjects.item"
            | "excel.listcolumns.item"
            | "excel.listrows.item"
            | "excel.filters.item"
            | "excel.range.autofilter-3289"
            | "excel.range.removeduplicates"
    )
}

fn modifies_range_in_place(id: &str) -> bool {
    matches!(
        id,
        "excel.range.sort"
            | "excel.sort.apply"
            | "excel.range.removeduplicates"
            | "excel.range.insert"
            | "excel.range.delete"
            | "excel.range.clear"
            | "excel.range.clearformats"
            | "excel.range.clearcomments"
            | "excel.range.clearhyperlinks"
            | "excel.range.copy"
            | "excel.range.cut"
            | "excel.range.pastespecial-1928"
    )
}

fn stateful_filter(id: &str) -> bool {
    matches!(
        id,
        "excel.worksheet.autofiltermode"
            | "excel.worksheet.filtermode"
            | "excel.worksheet.showalldata"
            | "excel.worksheet.autofilter-3289"
            | "excel.listobject.autofilter-3289"
            | "excel.autofilter.range"
            | "excel.autofilter.filters"
            | "excel.filters.item"
            | "excel.filter.on"
            | "excel.filter.operator-2641"
            | "excel.filter.criteria1"
            | "excel.filter.criteria2"
            | "excel.range.autofilter-3289"
    )
}

fn clipboard_dependent(id: &str) -> bool {
    matches!(
        id,
        "excel.range.copy" | "excel.range.cut" | "excel.range.pastespecial-1928"
    )
}

fn mixed_value_possible(id: &str) -> bool {
    matches!(
        id,
        "excel.range.numberformat"
            | "excel.range.horizontalalignment"
            | "excel.range.verticalalignment"
            | "excel.range.wraptext"
            | "excel.range.rowheight"
            | "excel.range.columnwidth"
            | "excel.range.formula"
            | "excel.range.formula2"
            | "excel.range.formular1c1"
            | "excel.range.formula2r1c1"
            | "excel.range.formulalocal"
            | "excel.range.formular1c1local"
            | "excel.range.hasformula"
            | "excel.font.name"
            | "excel.font.size"
            | "excel.font.bold"
            | "excel.font.italic"
            | "excel.font.underline"
            | "excel.font.strikethrough"
            | "excel.font.color"
            | "excel.font.colorindex"
            | "excel.interior.color"
            | "excel.interior.colorindex"
            | "excel.interior.pattern"
            | "excel.interior.patterncolor"
            | "excel.interior.patterncolorindex"
            | "excel.borders.color"
            | "excel.borders.colorindex"
            | "excel.borders.linestyle"
            | "excel.borders.weight"
            | "excel.border.color"
            | "excel.border.colorindex"
            | "excel.border.linestyle"
            | "excel.border.weight"
    )
}

fn version_sensitive(id: &str) -> bool {
    matches!(
        id,
        "excel.range.formula2"
            | "excel.range.formula2r1c1"
            | "excel.range.hasspill"
            | "excel.range.spillingtorange"
            | "excel.range.spillparent"
    )
}

fn returns_range_or_nothing(id: &str) -> bool {
    matches!(
        id,
        "excel.range.find" | "excel.range.findnext" | "excel.range.findprevious"
    )
}

fn stateful_search(id: &str) -> bool {
    matches!(
        id,
        "excel.range.find"
            | "excel.range.findnext"
            | "excel.range.findprevious"
            | "excel.range.replace-3305"
    )
}

fn relationship_fallback(id: &str) -> Option<&'static str> {
    match id {
        "excel.workbook.worksheets" => Some("Worksheets"),
        "excel.worksheets.item" => Some("Worksheet"),
        _ => None,
    }
}

fn documentation(info: &ITypeInfo, memid: i32) -> (String, String) {
    let mut name = BSTR::new();
    let mut text = BSTR::new();
    let mut context = 0;
    // SAFETY: the ITypeInfo is valid and BSTR output owners remain alive through the call.
    let _ = unsafe {
        info.GetDocumentation(memid, Some(&mut name), Some(&mut text), &mut context, None)
    };
    (name.to_string(), text.to_string())
}
fn typekind(kind: windows::Win32::System::Com::TYPEKIND) -> &'static str {
    match kind {
        TKIND_DISPATCH => "dispatch-interface",
        TKIND_INTERFACE => "interface",
        TKIND_COCLASS => "coclass",
        TKIND_ALIAS => "alias",
        _ => "other",
    }
}
fn member_kind(kind: windows::Win32::System::Com::INVOKEKIND) -> &'static str {
    use windows::Win32::System::Com::*;
    match kind {
        INVOKE_PROPERTYGET => "PROPERTYGET",
        INVOKE_PROPERTYPUT => "PROPERTYPUT",
        INVOKE_PROPERTYPUTREF => "PROPERTYPUTREF",
        INVOKE_FUNC => "METHOD",
        _ => "UNKNOWN",
    }
}
fn raw_type(info: &ITypeInfo, descriptor: &TYPEDESC) -> String {
    let value = descriptor.vt.0;
    if value == VT_PTR.0 {
        return format!(
            "VT_PTR({})",
            pointee(descriptor)
                .map(|nested| raw_type(info, nested))
                .unwrap_or_else(|| "null".to_owned())
        );
    }
    if value == VT_USERDEFINED.0 {
        return format!(
            "VT_USERDEFINED({})",
            referenced_type_name(info, descriptor).unwrap_or_else(|| "Unknown".to_owned())
        );
    }
    format!("VT_{} ({value})", primitive_type_name(value))
}
fn normalized_type(info: &ITypeInfo, descriptor: &TYPEDESC) -> String {
    let value = descriptor.vt.0;
    if value == VT_PTR.0 {
        return pointee(descriptor)
            .map(|nested| normalized_type(info, nested))
            .unwrap_or_else(|| "Unknown".to_owned());
    }
    if value == VT_USERDEFINED.0 {
        return referenced_type_name(info, descriptor).unwrap_or_else(|| "Unknown".to_owned());
    }
    primitive_type_name(value).to_owned()
}
fn referenced_type_name(info: &ITypeInfo, descriptor: &TYPEDESC) -> Option<String> {
    // SAFETY: VT_USERDEFINED selects hreftype in the TYPEDESC union.
    let reference = unsafe { descriptor.Anonymous.hreftype };
    // SAFETY: the href belongs to this ITypeInfo and GetRefTypeInfo returns an owned COM wrapper.
    let target = unsafe { info.GetRefTypeInfo(reference) }.ok()?;
    Some(model::canonical_name(&documentation(&target, -1).0).to_owned())
}
fn pointee(descriptor: &TYPEDESC) -> Option<&TYPEDESC> {
    // SAFETY: VT_PTR selects lptdesc in the TYPEDESC union; null denotes an unusable signature.
    unsafe { descriptor.Anonymous.lptdesc.as_ref() }
}
fn primitive_type_name(value: u16) -> &'static str {
    match value {
        value if value == VT_BSTR.0 => "String",
        value if value == VT_BOOL.0 => "bool",
        value if value == VT_I2.0 => "i16",
        value if value == VT_I4.0 => "i32",
        value if value == VT_I8.0 => "i64",
        value if value == VT_R4.0 => "f32",
        value if value == VT_R8.0 => "f64",
        value if value == VT_DATE.0 => "OaDate",
        value if value == VT_CY.0 => "Currency",
        value if value == VT_ERROR.0 => "ExcelError",
        value if value == VT_VARIANT.0 => "AutomationValue",
        value if value == VT_DISPATCH.0 => "Object",
        _ => "Unknown",
    }
}
fn direction(flags: u16) -> &'static str {
    use windows::Win32::System::Ole::{PARAMFLAG_FIN, PARAMFLAG_FOUT};
    match (flags & PARAMFLAG_FIN.0 != 0, flags & PARAMFLAG_FOUT.0 != 0) {
        (true, true) => "in-out",
        (true, false) => "in",
        (false, true) => "out",
        _ => "unknown",
    }
}
fn guid(guid: &GUID) -> String {
    let raw = format!("{:032x}", guid.to_u128());
    format!(
        "{{{}-{}-{}-{}-{}}}",
        &raw[0..8],
        &raw[8..12],
        &raw[12..16],
        &raw[16..20],
        &raw[20..32]
    )
}
fn write_json(path: &Path, value: &Value) -> Result<(), String> {
    write_text(
        path,
        &serde_json::to_string_pretty(value).map_err(|error| error.to_string())?,
    )
}
fn write_text(path: &Path, content: &str) -> Result<(), String> {
    if let Some(parent) = path.parent() {
        fs::create_dir_all(parent).map_err(|error| error.to_string())?;
    }
    fs::write(path, format!("{}\n", content.trim_end())).map_err(|error| error.to_string())
}
