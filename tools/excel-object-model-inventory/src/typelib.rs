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
        "Workbooks" | "Worksheets" => vec!["one-based-integer", "string-key"],
        "Areas" => vec!["one-based-integer"],
        _ if is_collection => vec!["variant-key"],
        _ => vec!["no-index"],
    };
    let iterator_status = match model::canonical_name(raw_name) {
        "Workbooks" | "Worksheets" | "Areas" => "implemented",
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
    Ok(
        json!({"schema_version": SCHEMA_VERSION, "id": object_id, "name": model::canonical_name(raw_name), "kind": if event_interface { "event-interface" } else { kind }, "surface_class": model::surface_class(raw_name, kind, event_interface, attr.wTypeFlags), "roadmap_class": model::roadmap_class(raw_name, kind, event_interface), "typelib_type_flags": attr.wTypeFlags, "guid": guid(&attr.guid), "source_interface": raw_name, "typelib_version": {"major":1,"minor":9}, "documentation_url": model::documentation_url(raw_name), "implemented_status": if event_interface { "not-started" } else if wrapper_object { "partial" } else { "metadata-only" }, "documentation_status": if priority_documentation { "reviewed" } else { "generated" }, "test_status": if wrapper_object { "live-tested" } else { "not-tested" }, "implemented_interface_count": attr.cImplTypes, "alias_target": alias_target, "collection": collection, "members": members}),
    )
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
        | "excel.workbook.saved"
        | "excel.worksheet.name"
        | "excel.worksheet.visible"
        | "excel.range.value"
        | "excel.range.value2"
        | "excel.range.formula"
        | "excel.range.formula2" => &["PROPERTYGET", "PROPERTYPUT"],
        "excel.application.quit"
        | "excel.workbook.close"
        | "excel.workbooks.open-1923"
        | "excel.workbook.save"
        | "excel.workbook.saveas-3174"
        | "excel.workbook.savecopyas" => &["METHOD"],
        "excel.workbooks.add" => &["PROPERTYGET"],
        "excel.worksheets.add" | "excel.range.clearcontents" => &["METHOD"],
        "excel.application.version"
        | "excel.application.workbooks"
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
        | "excel.range.address"
        | "excel.range.row"
        | "excel.range.column"
        | "excel.range.count"
        | "excel.range.rows"
        | "excel.range.columns" => &["PROPERTYGET"],
        _ => &[],
    }
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
