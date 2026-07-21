//! Read-only inspection of the installed Excel Automation type library.
//!
//! The tool deliberately records Excel-specific ABI evidence in a layer beside
//! the documentation knowledge base. It neither creates Excel nor changes COM
//! registration, and all type-information descriptors are released by guards.

#![cfg(windows)]

use serde_json::{Value, json};
use sha2::{Digest, Sha256};
use std::collections::{BTreeMap, BTreeSet};
use std::fs;
use std::path::{Path, PathBuf};
use windows::Win32::System::Com::{
    CC_STDCALL, COINIT_APARTMENTTHREADED, CoInitializeEx, CoUninitialize, FUNCDESC,
    FUNCFLAG_FDEFAULTBIND, FUNCFLAG_FDEFAULTCOLLELEM, FUNCFLAG_FHIDDEN, FUNCFLAG_FNONBROWSABLE,
    FUNCFLAG_FRESTRICTED, IMPLTYPEFLAG_FDEFAULT, IMPLTYPEFLAG_FSOURCE, INVOKE_FUNC,
    INVOKE_PROPERTYGET, INVOKE_PROPERTYPUT, INVOKE_PROPERTYPUTREF, ITypeInfo, ITypeLib,
    TKIND_ALIAS, TKIND_COCLASS, TKIND_DISPATCH, TKIND_ENUM, TKIND_INTERFACE, TYPEATTR, TYPEDESC,
    TYPEKIND,
};
use windows::Win32::System::Ole::{
    LoadRegTypeLib, LoadTypeLibEx, PARAMDESC, PARAMFLAG_FHASDEFAULT, PARAMFLAG_FIN, PARAMFLAG_FOPT,
    PARAMFLAG_FOUT, PARAMFLAG_FRETVAL, REGKIND_NONE,
};
use windows::Win32::System::Variant::{
    VARENUM, VARIANT, VT_BOOL, VT_BSTR, VT_CARRAY, VT_CY, VT_DATE, VT_DISPATCH, VT_EMPTY, VT_ERROR,
    VT_HRESULT, VT_I1, VT_I2, VT_I4, VT_I8, VT_INT, VT_NULL, VT_PTR, VT_R4, VT_R8, VT_SAFEARRAY,
    VT_UI1, VT_UI2, VT_UI4, VT_UI8, VT_UINT, VT_UNKNOWN, VT_USERDEFINED, VT_VARIANT, VT_VOID,
};
use windows::core::{BSTR, GUID, HSTRING};

const EXCEL_TYPELIB: GUID = GUID::from_u128(0x00020813_0000_0000_c000_000000000046);
const EXCEL_MAJOR: u16 = 1;
const EXCEL_MINOR: u16 = 9;
const TOOL_VERSION: u32 = 1;

#[derive(Debug, Clone)]
pub struct AuditInput {
    pub typelib_path: Option<PathBuf>,
    pub windows_version: String,
    pub excel_file_version: String,
    pub office_bitness: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AuditSummary {
    pub coclasses: usize,
    pub interfaces: usize,
    pub members: usize,
    pub parameters: usize,
    pub enums: usize,
}

#[derive(Debug, Clone)]
struct TypeEntry {
    index: u32,
    name: String,
    kind: TYPEKIND,
    guid: String,
}

#[derive(Debug, Clone)]
struct AuditData {
    library: Value,
    coclasses: Vec<Value>,
    interfaces: Vec<Value>,
    members: Vec<Value>,
    parameters: Vec<Value>,
    enums: Vec<Value>,
    aliases: Vec<Value>,
    unresolved: Vec<Value>,
    documentation: BTreeMap<String, Value>,
    source_manifest: String,
}

pub fn audit(root: &Path, input: &AuditInput) -> Result<AuditSummary, String> {
    let data = inspect(root, input)?;
    let summary = summary(&data);
    write_artifacts(root, &data)?;
    Ok(summary)
}

pub fn check(root: &Path, input: &AuditInput) -> Result<(), String> {
    let first = inspect(root, input)?;
    let second = inspect(root, input)?;
    let first_artifacts = artifacts(&first)?;
    let second_artifacts = artifacts(&second)?;
    if first_artifacts != second_artifacts {
        return Err("type-library inspection is not deterministic for this input".to_owned());
    }
    for (relative, expected) in first_artifacts {
        let path = root.join(relative);
        let actual = fs::read_to_string(&path)
            .map_err(|error| format!("cannot read {}: {error}", path.display()))?;
        if actual != expected {
            return Err(format!(
                "generated type-library artifact {} is stale; run audit",
                path.display()
            ));
        }
        if actual.contains("\r\n") || !actual.ends_with('\n') {
            return Err(format!(
                "type-library artifact {} must use LF endings and a final newline",
                path.display()
            ));
        }
    }
    Ok(())
}

fn inspect(root: &Path, input: &AuditInput) -> Result<AuditData, String> {
    let _apartment = Apartment::initialize()?;
    let (library, file) = load_library(input)?;
    let library_attr = TypeLibAttr::acquire(&library)?;
    let attributes = unsafe { *library_attr.pointer };
    let (library_name, library_doc) = library_documentation(&library);
    let type_entries = type_entries(&library)?;
    let type_lookup: BTreeMap<String, TypeEntry> = type_entries
        .iter()
        .map(|entry| (entry.name.clone(), entry.clone()))
        .collect();

    let mut interfaces = Vec::new();
    let mut members = Vec::new();
    let mut parameters = Vec::new();
    let mut referenced_enums = BTreeSet::new();
    let mut referenced_aliases = BTreeSet::new();
    let mut unresolved = Vec::new();

    for entry in type_entries
        .iter()
        .filter(|entry| targeted_interface(entry))
    {
        let info = unsafe { library.GetTypeInfo(entry.index) }
            .map_err(|error| format!("cannot load type info {}: {error}", entry.name))?;
        let attr = TypeAttr::acquire(&info)?;
        let type_attr = unsafe { *attr.pointer };
        let owner = canonical_owner(&entry.name);
        let implementations = implementation_records(&info, type_attr.cImplTypes)?;
        interfaces.push(json!({
            "schema_version": 1,
            "id": stable_id("interface", &entry.name),
            "canonical_owner": owner,
            "name": entry.name,
            "guid": entry.guid,
            "typekind": typekind_name(entry.kind),
            "type_flags_raw": type_attr.wTypeFlags,
            "is_dual": (type_attr.wTypeFlags & windows::Win32::System::Ole::TYPEFLAG_FDUAL.0 as u16) != 0,
            "is_dispatchable": (type_attr.wTypeFlags & windows::Win32::System::Ole::TYPEFLAG_FDISPATCHABLE.0 as u16) != 0,
            "implementation_types": implementations,
            "documentation_name": entry.name,
            "declared_member_count": type_attr.cFuncs,
            "source": "installed-excel-typelib"
        }));
        for index in 0..u32::from(type_attr.cFuncs) {
            let descriptor = FuncDesc::acquire(&info, index)?;
            let function = unsafe { *descriptor.pointer };
            let names = member_names(&info, function.memid, function.cParams)?;
            let member_name = names
                .first()
                .cloned()
                .unwrap_or_else(|| format!("memid_{}", function.memid));
            if !audited_member(owner, &member_name) {
                continue;
            }
            let signature = function_record(
                &info,
                owner,
                &entry.name,
                &entry.guid,
                &member_name,
                &names,
                &function,
                &type_lookup,
                &mut referenced_enums,
                &mut referenced_aliases,
            )?;
            let member_id = signature
                .get("id")
                .and_then(Value::as_str)
                .ok_or_else(|| "member record omitted stable ID".to_owned())?
                .to_owned();
            for parameter in signature
                .get("parameters")
                .and_then(Value::as_array)
                .into_iter()
                .flatten()
            {
                let mut parameter = parameter.clone();
                parameter["member_id"] = Value::String(member_id.clone());
                parameters.push(parameter);
            }
            let mut member = signature;
            member.as_object_mut().expect("object").remove("parameters");
            members.push(member);
        }
    }

    let coclasses = coclass_records(&library, &type_entries)?;
    let minimum_enums: BTreeSet<String> = [
        "XlCalculation",
        "XlFileFormat",
        "XlReferenceStyle",
        "XlDirection",
        "XlFindLookIn",
        "XlLookAt",
        "XlSearchOrder",
        "XlSearchDirection",
        "XlSortOrder",
        "XlSortOrientation",
        "XlSheetType",
        "XlFixedFormatType",
        "XlFixedFormatQuality",
    ]
    .into_iter()
    .map(str::to_owned)
    .collect();
    referenced_enums.extend(minimum_enums);
    let enums = enum_records(&library, &type_entries, &referenced_enums, &mut unresolved)?;
    let aliases = alias_records(&library, &type_lookup, &referenced_aliases, &mut unresolved)?;
    let documentation = load_documentation_members(root)?;
    add_candidate_unresolved(&members, &documentation, &mut unresolved);

    let library = json!({
        "schema_version": 1,
        "id": "Excel.TypeLib",
        "name": library_name,
        "documentation": library_doc,
        "guid": guid_string(&attributes.guid),
        "major_version": attributes.wMajorVerNum,
        "minor_version": attributes.wMinorVerNum,
        "lcid": attributes.lcid,
        "syskind": syskind_name(attributes.syskind.0),
        "library_flags_raw": attributes.wLibFlags,
        "type_info_count": type_entries.len(),
        "source": "installed-excel-typelib"
    });
    let source_manifest = source_manifest(&library, &file, input);
    sort_by_id(&mut coclasses.clone());
    let mut data = AuditData {
        library,
        coclasses,
        interfaces,
        members,
        parameters,
        enums,
        aliases,
        unresolved,
        documentation,
        source_manifest,
    };
    sort_data(&mut data);
    Ok(data)
}

fn load_library(input: &AuditInput) -> Result<(ITypeLib, FileEvidence), String> {
    match &input.typelib_path {
        Some(path) => {
            let canonical = fs::canonicalize(path).map_err(|error| {
                format!(
                    "cannot resolve explicit type-library path {}: {error}",
                    path.display()
                )
            })?;
            let hstring = HSTRING::from(canonical.to_string_lossy().as_ref());
            let library = unsafe { LoadTypeLibEx(&hstring, REGKIND_NONE) }.map_err(|error| {
                format!(
                    "cannot load Excel type library {}: {error}",
                    canonical.display()
                )
            })?;
            Ok((library, file_evidence(&canonical)?))
        }
        None => {
            let library = unsafe { LoadRegTypeLib(&EXCEL_TYPELIB, EXCEL_MAJOR, EXCEL_MINOR, 0) }
                .map_err(|error| {
                    format!(
                        "Excel {EXCEL_MAJOR}.{EXCEL_MINOR} type library is not registered: {error}"
                    )
                })?;
            Ok((
                library,
                FileEvidence {
                    basename: "registered-type-library".to_owned(),
                    sha256: "not-recorded-without-explicit-path".to_owned(),
                    registration_category:
                        "HKCR\\TypeLib\\{00020813-0000-0000-C000-000000000046}\\1.9".to_owned(),
                },
            ))
        }
    }
}

#[derive(Debug, Clone)]
struct FileEvidence {
    basename: String,
    sha256: String,
    registration_category: String,
}

fn file_evidence(path: &Path) -> Result<FileEvidence, String> {
    let bytes =
        fs::read(path).map_err(|error| format!("cannot hash {}: {error}", path.display()))?;
    let hash = Sha256::digest(bytes);
    let basename = path
        .file_name()
        .and_then(|name| name.to_str())
        .ok_or_else(|| {
            format!(
                "type-library path {} has no portable basename",
                path.display()
            )
        })?
        .to_owned();
    Ok(FileEvidence {
        basename,
        sha256: format!("{hash:x}"),
        registration_category:
            "HKCR\\TypeLib\\{00020813-0000-0000-C000-000000000046}\\1.9\\0\\Win64".to_owned(),
    })
}

fn type_entries(library: &ITypeLib) -> Result<Vec<TypeEntry>, String> {
    let mut entries = Vec::new();
    for index in 0..unsafe { library.GetTypeInfoCount() } {
        let info = unsafe { library.GetTypeInfo(index) }
            .map_err(|error| format!("cannot read type info index {index}: {error}"))?;
        let attr = TypeAttr::acquire(&info)?;
        let type_attr = unsafe { *attr.pointer };
        let (name, _) = type_documentation(&info);
        entries.push(TypeEntry {
            index,
            name,
            kind: type_attr.typekind,
            guid: guid_string(&type_attr.guid),
        });
    }
    entries.sort_by(|left, right| {
        left.name
            .cmp(&right.name)
            .then(left.index.cmp(&right.index))
    });
    Ok(entries)
}

fn coclass_records(library: &ITypeLib, entries: &[TypeEntry]) -> Result<Vec<Value>, String> {
    let mut records = Vec::new();
    for entry in entries
        .iter()
        .filter(|entry| entry.kind == TKIND_COCLASS && entry.name == "Application")
    {
        let info = unsafe { library.GetTypeInfo(entry.index) }
            .map_err(|error| format!("cannot load coclass {}: {error}", entry.name))?;
        let attr = TypeAttr::acquire(&info)?;
        let value = unsafe { *attr.pointer };
        records.push(json!({
            "schema_version": 1,
            "id": "Excel.Application.Coclass",
            "name": entry.name,
            "guid": entry.guid,
            "typekind": typekind_name(entry.kind),
            "type_flags_raw": value.wTypeFlags,
            "implemented_interfaces": implementation_records(&info, value.cImplTypes)?,
            "documentation_name": entry.name,
            "source": "installed-excel-typelib"
        }));
    }
    if records.is_empty() {
        records.push(json!({
            "schema_version": 1,
            "id": "Excel.Application.Coclass",
            "status": "missing-from-installed-typelib",
            "source": "installed-excel-typelib"
        }));
    }
    Ok(records)
}

fn implementation_records(info: &ITypeInfo, count: u16) -> Result<Vec<Value>, String> {
    let mut records = Vec::new();
    for index in 0..u32::from(count) {
        let href = unsafe { info.GetRefTypeOfImplType(index) }
            .map_err(|error| format!("cannot read implemented type {index}: {error}"))?;
        let reference = unsafe { info.GetRefTypeInfo(href) }
            .map_err(|error| format!("cannot load implemented type {index}: {error}"))?;
        let (name, _) = type_documentation(&reference);
        let attr = TypeAttr::acquire(&reference)?;
        let value = unsafe { *attr.pointer };
        let flags = unsafe { info.GetImplTypeFlags(index) }
            .map_err(|error| format!("cannot read implementation flags {index}: {error}"))?;
        records.push(json!({
            "name": name,
            "guid": guid_string(&value.guid),
            "typekind": typekind_name(value.typekind),
            "flags_raw": flags.0,
            "default": (flags.0 & IMPLTYPEFLAG_FDEFAULT.0) != 0,
            "source": (flags.0 & IMPLTYPEFLAG_FSOURCE.0) != 0,
        }));
    }
    Ok(records)
}

#[allow(clippy::too_many_arguments)]
fn function_record(
    info: &ITypeInfo,
    owner: &str,
    interface_name: &str,
    interface_guid: &str,
    member_name: &str,
    names: &[String],
    function: &FUNCDESC,
    types: &BTreeMap<String, TypeEntry>,
    referenced_enums: &mut BTreeSet<String>,
    referenced_aliases: &mut BTreeSet<String>,
) -> Result<Value, String> {
    let return_type = normalize_type(info, &function.elemdescFunc.tdesc, types)?;
    collect_user_types(&return_type, referenced_enums, referenced_aliases);
    let mut params = Vec::new();
    for index in
        0..usize::try_from(function.cParams).map_err(|_| "negative parameter count".to_owned())?
    {
        let element = unsafe { *function.lprgelemdescParam.add(index) };
        let parameter_type = normalize_type(info, &element.tdesc, types)?;
        collect_user_types(&parameter_type, referenced_enums, referenced_aliases);
        let parameter_name = names
            .get(index + 1)
            .cloned()
            .unwrap_or_else(|| format!("arg{}", index + 1));
        let paramdesc = unsafe { element.Anonymous.paramdesc };
        params.push(json!({
            "schema_version": 1,
            "id": stable_id("parameter", &format!("{owner}.{member_name}.{}.{}", invkind_name(function.invkind), index)),
            "ordinal": index,
            "name": parameter_name,
            "type": parameter_type,
            "param_flags_raw": paramdesc.wParamFlags.0,
            "param_flags": parameter_flags(paramdesc.wParamFlags.0),
            "optional": (paramdesc.wParamFlags.0 & PARAMFLAG_FOPT.0) != 0,
            "required": (paramdesc.wParamFlags.0 & PARAMFLAG_FOPT.0) == 0,
            "default": parameter_default(&paramdesc),
        }));
    }
    let canonical_id =
        canonical_member_id(owner, member_name).unwrap_or_else(|| format!("{owner}.{member_name}"));
    let invkind = invkind_name(function.invkind);
    Ok(json!({
        "schema_version": 1,
        "id": stable_id("member", &format!("{canonical_id}.{invkind}")),
        "canonical_documentation_id": canonical_id,
        "owner": owner,
        "typelib_interface": interface_name,
        "typelib_interface_guid": interface_guid,
        "name": member_name,
        "dispid": function.memid,
        "invoke_kind": invkind,
        "invoke_kind_raw": function.invkind.0,
        "calling_convention": calling_convention(function.callconv.0),
        "calling_convention_raw": function.callconv.0,
        "return_type": return_type,
        "parameter_count": function.cParams,
        "optional_parameter_count": function.cParamsOpt,
        "function_flags_raw": function.wFuncFlags.0,
        "function_flags": function_flags(function.wFuncFlags.0),
        "default_member": function.memid == 0 || member_name == "_Default" || (function.wFuncFlags.0 & FUNCFLAG_FDEFAULTBIND.0) != 0 || (function.wFuncFlags.0 & FUNCFLAG_FDEFAULTCOLLELEM.0) != 0,
        "hidden": (function.wFuncFlags.0 & FUNCFLAG_FHIDDEN.0) != 0,
        "restricted": (function.wFuncFlags.0 & FUNCFLAG_FRESTRICTED.0) != 0,
        "nonbrowsable": (function.wFuncFlags.0 & FUNCFLAG_FNONBROWSABLE.0) != 0,
        "inherited": false,
        "declared": true,
        "documentation": type_documentation(info).1,
        "parameters": params,
        "source": "installed-excel-typelib"
    }))
}

fn member_names(info: &ITypeInfo, memid: i32, count: i16) -> Result<Vec<String>, String> {
    let requested = usize::try_from(count.max(0)).unwrap_or(0) + 1;
    let mut bstrs = (0..requested).map(|_| BSTR::new()).collect::<Vec<_>>();
    let mut actual = 0;
    unsafe { info.GetNames(memid, &mut bstrs, &mut actual) }
        .map_err(|error| format!("cannot obtain names for DISPID {memid}: {error}"))?;
    Ok(bstrs
        .into_iter()
        .take(usize::try_from(actual).unwrap_or(0))
        .map(|value| value.to_string())
        .collect())
}

fn enum_records(
    library: &ITypeLib,
    entries: &[TypeEntry],
    wanted: &BTreeSet<String>,
    unresolved: &mut Vec<Value>,
) -> Result<Vec<Value>, String> {
    let mut records = Vec::new();
    for name in wanted {
        let Some(entry) = entries
            .iter()
            .find(|entry| entry.name == *name && entry.kind == TKIND_ENUM)
        else {
            unresolved.push(json!({
                "schema_version": 1,
                "id": stable_id("unresolved", &format!("enum.{name}")),
                "category": "enum-missing-from-installed-typelib",
                "target": name,
                "detail": "No enum with this name was located in the installed Excel type library."
            }));
            continue;
        };
        let info = unsafe { library.GetTypeInfo(entry.index) }
            .map_err(|error| format!("cannot load enum {name}: {error}"))?;
        let attr = TypeAttr::acquire(&info)?;
        let type_attr = unsafe { *attr.pointer };
        let mut values = Vec::new();
        for index in 0..u32::from(type_attr.cVars) {
            let descriptor = VarDesc::acquire(&info, index)?;
            let variable = unsafe { *descriptor.pointer };
            let (member_name, documentation) = documentation_for_memid(&info, variable.memid);
            values.push(json!({
                "name": member_name,
                "dispid": variable.memid,
                "value": unsafe { variable.Anonymous.lpvarValue.as_ref() }.map(variant_default).unwrap_or_else(|| json!({"kind": "raw", "detail": "enum value pointer absent"})),
                "documentation": documentation,
            }));
        }
        values.sort_by_key(|value| value_string(value, "name"));
        records.push(json!({
            "schema_version": 1,
            "id": stable_id("enum", name),
            "name": name,
            "guid": entry.guid,
            "values": values,
            "source": "installed-excel-typelib"
        }));
    }
    Ok(records)
}

fn alias_records(
    library: &ITypeLib,
    entries: &BTreeMap<String, TypeEntry>,
    wanted: &BTreeSet<String>,
    unresolved: &mut Vec<Value>,
) -> Result<Vec<Value>, String> {
    let mut records = Vec::new();
    for name in wanted {
        let Some(entry) = entries.get(name).filter(|entry| entry.kind == TKIND_ALIAS) else {
            unresolved.push(json!({
                "schema_version": 1,
                "id": stable_id("unresolved", &format!("alias.{name}")),
                "category": "alias-not-resolved",
                "target": name,
                "detail": "The referenced user-defined type was not a resolvable alias."
            }));
            continue;
        };
        let info = unsafe { library.GetTypeInfo(entry.index) }
            .map_err(|error| format!("cannot load alias {name}: {error}"))?;
        let attr = TypeAttr::acquire(&info)?;
        let value = unsafe { *attr.pointer };
        records.push(json!({
            "schema_version": 1,
            "id": stable_id("alias", name),
            "name": name,
            "guid": entry.guid,
            "target": normalize_type(&info, &value.tdescAlias, entries)?,
            "source": "installed-excel-typelib"
        }));
    }
    Ok(records)
}

fn normalize_type(
    info: &ITypeInfo,
    descriptor: &TYPEDESC,
    entries: &BTreeMap<String, TypeEntry>,
) -> Result<Value, String> {
    if descriptor.vt == VT_PTR {
        let target = unsafe { descriptor.Anonymous.lptdesc.as_ref() }
            .ok_or_else(|| "VT_PTR type description had a null target".to_owned())?;
        return Ok(json!({"kind": "pointer", "target": normalize_type(info, target, entries)?}));
    }
    if descriptor.vt == VT_SAFEARRAY {
        let target = unsafe { descriptor.Anonymous.lptdesc.as_ref() }
            .ok_or_else(|| "VT_SAFEARRAY type description had a null element".to_owned())?;
        return Ok(json!({"kind": "safearray", "element": normalize_type(info, target, entries)?}));
    }
    if descriptor.vt == VT_CARRAY {
        let array = unsafe { descriptor.Anonymous.lpadesc.as_ref() }
            .ok_or_else(|| "VT_CARRAY type description had a null array descriptor".to_owned())?;
        return Ok(json!({
            "kind": "c_array",
            "dimensions": array.cDims,
            "element": normalize_type(info, &array.tdescElem, entries)?,
        }));
    }
    if descriptor.vt == VT_USERDEFINED {
        let reference = unsafe { info.GetRefTypeInfo(descriptor.Anonymous.hreftype) }
            .map_err(|error| format!("cannot resolve user-defined type: {error}"))?;
        let (name, _) = type_documentation(&reference);
        let attr = TypeAttr::acquire(&reference)?;
        let type_attr = unsafe { *attr.pointer };
        let entry = entries.get(&name);
        return Ok(
            match entry.map(|value| value.kind).unwrap_or(type_attr.typekind) {
                TKIND_ENUM => json!({"kind": "enum", "id": format!("Excel.{name}"), "name": name}),
                TKIND_ALIAS => {
                    json!({"kind": "alias", "id": format!("Excel.{name}"), "name": name})
                }
                TKIND_INTERFACE | TKIND_DISPATCH => {
                    json!({"kind": "interface", "id": canonical_owner(&name), "name": name, "guid": guid_string(&type_attr.guid)})
                }
                _ => {
                    json!({"kind": "user_defined", "name": name, "typekind": typekind_name(type_attr.typekind), "guid": guid_string(&type_attr.guid)})
                }
            },
        );
    }
    Ok(primitive_type(descriptor.vt))
}

fn primitive_type(vartype: VARENUM) -> Value {
    let (kind, name) = match vartype {
        VT_VOID => ("void", "VT_VOID"),
        VT_EMPTY => ("empty", "VT_EMPTY"),
        VT_NULL => ("null", "VT_NULL"),
        VT_I1 => ("primitive", "VT_I1"),
        VT_I2 => ("primitive", "VT_I2"),
        VT_I4 => ("primitive", "VT_I4"),
        VT_I8 => ("primitive", "VT_I8"),
        VT_UI1 => ("primitive", "VT_UI1"),
        VT_UI2 => ("primitive", "VT_UI2"),
        VT_UI4 => ("primitive", "VT_UI4"),
        VT_UI8 => ("primitive", "VT_UI8"),
        VT_INT => ("primitive", "VT_INT"),
        VT_UINT => ("primitive", "VT_UINT"),
        VT_R4 => ("primitive", "VT_R4"),
        VT_R8 => ("primitive", "VT_R8"),
        VT_CY => ("primitive", "VT_CY"),
        VT_DATE => ("primitive", "VT_DATE"),
        VT_BSTR => ("bstr", "VT_BSTR"),
        VT_VARIANT => ("variant", "VT_VARIANT"),
        VT_DISPATCH => ("idispatch", "VT_DISPATCH"),
        VT_UNKNOWN => ("iunknown", "VT_UNKNOWN"),
        VT_BOOL => ("primitive", "VT_BOOL"),
        VT_ERROR => ("primitive", "VT_ERROR"),
        VT_HRESULT => ("hresult", "VT_HRESULT"),
        _ => ("unknown", "unrecognized VARENUM"),
    };
    json!({"kind": kind, "vartype": name, "vartype_raw": vartype.0})
}

fn parameter_default(parameter: &PARAMDESC) -> Value {
    if (parameter.wParamFlags.0 & PARAMFLAG_FHASDEFAULT.0) == 0 {
        return Value::Null;
    }
    let Some(default) = (unsafe { parameter.pparamdescex.as_ref() }) else {
        return json!({"kind": "raw", "detail": "PARAMFLAG_FHASDEFAULT without PARAMDESCEX"});
    };
    variant_default(&default.varDefaultValue)
}

fn variant_default(value: &VARIANT) -> Value {
    let inner = unsafe {
        &*(&value.Anonymous.Anonymous as *const std::mem::ManuallyDrop<_>
            as *const windows::Win32::System::Variant::VARIANT_0_0)
    };
    let variant = inner.vt;
    let contents = &inner.Anonymous;
    let exact = unsafe {
        match variant {
            VT_EMPTY => json!(null),
            VT_NULL => json!("VT_NULL"),
            VT_I1 => json!(contents.cVal),
            VT_I2 => json!(contents.iVal),
            VT_I4 | VT_INT => json!(contents.lVal),
            VT_I8 => json!(contents.llVal),
            VT_UI1 => json!(contents.bVal),
            VT_UI2 => json!(contents.uiVal),
            VT_UI4 | VT_UINT => json!(contents.ulVal),
            VT_UI8 => json!(contents.ullVal),
            VT_R4 => json!(contents.fltVal),
            VT_R8 | VT_DATE => json!(contents.dblVal),
            VT_BOOL => json!(contents.boolVal.0),
            VT_ERROR | VT_HRESULT => json!(contents.scode),
            VT_BSTR => {
                let bstr =
                    &*(&contents.bstrVal as *const std::mem::ManuallyDrop<BSTR> as *const BSTR);
                json!(bstr.to_string())
            }
            _ => Value::Null,
        }
    };
    json!({
        "vartype": primitive_type(variant),
        "value": exact,
        "preservation": if exact.is_null() && variant != VT_EMPTY { "typed-raw-unsupported-value" } else { "exact" }
    })
}

fn collect_user_types(value: &Value, enums: &mut BTreeSet<String>, aliases: &mut BTreeSet<String>) {
    match value {
        Value::Object(map) => {
            if map.get("kind").and_then(Value::as_str) == Some("enum")
                && let Some(name) = map.get("name").and_then(Value::as_str)
            {
                enums.insert(name.to_owned());
            }
            if map.get("kind").and_then(Value::as_str) == Some("alias")
                && let Some(name) = map.get("name").and_then(Value::as_str)
            {
                aliases.insert(name.to_owned());
            }
            for nested in map.values() {
                collect_user_types(nested, enums, aliases);
            }
        }
        Value::Array(values) => {
            for nested in values {
                collect_user_types(nested, enums, aliases);
            }
        }
        _ => {}
    }
}

fn load_documentation_members(root: &Path) -> Result<BTreeMap<String, Value>, String> {
    let path = root.join("data").join("members.jsonl");
    let text = fs::read_to_string(&path).map_err(|error| {
        format!(
            "cannot read documentation members {}: {error}",
            path.display()
        )
    })?;
    let mut members = BTreeMap::new();
    for line in text.lines().filter(|line| !line.is_empty()) {
        let value: Value = serde_json::from_str(line)
            .map_err(|error| format!("cannot parse documentation member JSON: {error}"))?;
        if let Some(id) = value.get("id").and_then(Value::as_str) {
            members.insert(id.to_owned(), value);
        }
    }
    Ok(members)
}

fn add_candidate_unresolved(
    members: &[Value],
    documentation: &BTreeMap<String, Value>,
    unresolved: &mut Vec<Value>,
) {
    for &candidate in candidate_ids() {
        let found = members.iter().any(|member| {
            member
                .get("canonical_documentation_id")
                .and_then(Value::as_str)
                == Some(candidate)
        });
        if !found {
            unresolved.push(json!({
                "schema_version": 1,
                "id": stable_id("unresolved", &format!("candidate.{candidate}")),
                "category": "candidate-missing-from-installed-typelib",
                "target": candidate,
                "documentation_present": documentation.contains_key(candidate),
                "detail": "No selected interface member matched this Prompt 03 canonical candidate."
            }));
        }
    }
}

fn artifacts(data: &AuditData) -> Result<BTreeMap<PathBuf, String>, String> {
    let mut files = BTreeMap::new();
    files.insert(
        PathBuf::from("typelib/SOURCE_MANIFEST.toml"),
        data.source_manifest.clone(),
    );
    files.insert(
        PathBuf::from("typelib/library.json"),
        json_text(&data.library)?,
    );
    files.insert(
        PathBuf::from("typelib/coclasses.jsonl"),
        jsonl(&data.coclasses)?,
    );
    files.insert(
        PathBuf::from("typelib/interfaces.jsonl"),
        jsonl(&data.interfaces)?,
    );
    files.insert(
        PathBuf::from("typelib/members.jsonl"),
        jsonl(&data.members)?,
    );
    files.insert(
        PathBuf::from("typelib/parameters.jsonl"),
        jsonl(&data.parameters)?,
    );
    files.insert(PathBuf::from("typelib/enums.jsonl"), jsonl(&data.enums)?);
    files.insert(
        PathBuf::from("typelib/aliases.jsonl"),
        jsonl(&data.aliases)?,
    );
    files.insert(
        PathBuf::from("typelib/unresolved.jsonl"),
        jsonl(&data.unresolved)?,
    );
    for (name, contents) in reports(data) {
        files.insert(PathBuf::from("generated/typelib").join(name), contents);
    }
    Ok(files)
}

fn write_artifacts(root: &Path, data: &AuditData) -> Result<(), String> {
    for (relative, contents) in artifacts(data)? {
        let path = root.join(relative);
        let parent = path
            .parent()
            .ok_or_else(|| format!("no parent for {}", path.display()))?;
        fs::create_dir_all(parent)
            .map_err(|error| format!("cannot create {}: {error}", parent.display()))?;
        fs::write(&path, contents)
            .map_err(|error| format!("cannot write {}: {error}", path.display()))?;
    }
    Ok(())
}

fn reports(data: &AuditData) -> BTreeMap<&'static str, String> {
    let mut reports = BTreeMap::new();
    reports.insert("library-summary.md", library_summary(data));
    reports.insert("architectural-spine.md", spine_report(data));
    reports.insert("candidate-0.1-signatures.md", candidate_report(data));
    reports.insert("collections.md", collections_report(data));
    reports.insert("optional-arguments.md", optional_arguments_report(data));
    reports.insert("range-contracts.md", range_report(data));
    reports.insert("enum-values.md", enum_report(data));
    reports.insert(
        "documentation-differences.md",
        documentation_differences(data),
    );
    reports.insert("unresolved.md", unresolved_report(data));
    reports
}

fn report_header(title: &str, detail: &str) -> String {
    format!(
        "# {title}\n\nGenerated by `excel-com-typelib-audit`; do not edit by hand. {detail}\n\nThis is declared type-library metadata only. It does not establish Excel runtime behavior, activation, collection index bases, omitted-argument behavior, VARIANT runtime types, or SAFEARRAY shape.\n\n"
    )
}

fn library_summary(data: &AuditData) -> String {
    let mut output = report_header(
        "Excel type-library summary",
        "The source manifest records a portable file identity and no raw user path.",
    );
    output.push_str("| Field | Value |\n| --- | --- |\n");
    for field in [
        "name",
        "guid",
        "major_version",
        "minor_version",
        "lcid",
        "syskind",
        "type_info_count",
    ] {
        output.push_str(&format!(
            "| {field} | `{}` |\n",
            value_string(&data.library, field)
        ));
    }
    output.push_str(&format!(
        "\n| Audited coclasses | Interfaces | Members | Parameters | Enums | Aliases | Unresolved |\n| ---: | ---: | ---: | ---: | ---: | ---: | ---: |\n| {} | {} | {} | {} | {} | {} | {} |\n",
        data.coclasses.len(), data.interfaces.len(), data.members.len(), data.parameters.len(), data.enums.len(), data.aliases.len(), data.unresolved.len()
    ));
    output
}

fn spine_report(data: &AuditData) -> String {
    let mut output = report_header(
        "Architectural spine typelib audit",
        "Core interface identities and member counts are taken from the installed Excel library.",
    );
    output.push_str("| Canonical object | Typelib interface | GUID | Type flags | Audited members |\n| --- | --- | --- | ---: | ---: |\n");
    for owner in [
        "Excel.Application",
        "Excel.Workbooks",
        "Excel.Workbook",
        "Excel.Worksheets",
        "Excel.Worksheet",
        "Excel.Range",
    ] {
        let interface = data
            .interfaces
            .iter()
            .find(|record| record.get("canonical_owner").and_then(Value::as_str) == Some(owner));
        let members = data
            .members
            .iter()
            .filter(|record| record.get("owner").and_then(Value::as_str) == Some(owner))
            .count();
        match interface {
            Some(record) => output.push_str(&format!(
                "| `{owner}` | `{}` | `{}` | {} | {members} |\n",
                value_string(record, "name"),
                value_string(record, "guid"),
                value_string(record, "type_flags_raw")
            )),
            None => output.push_str(&format!("| `{owner}` | -- | -- | -- | {members} |\n")),
        }
    }
    output.push_str("\n## Application coclass\n\n");
    for coclass in &data.coclasses {
        output.push_str(&format!(
            "- `{}`: {}\n",
            value_string(coclass, "id"),
            json_inline(coclass.get("implemented_interfaces"))
        ));
    }
    output
}

fn candidate_report(data: &AuditData) -> String {
    let mut output = report_header(
        "Candidate 0.1 signatures",
        "Every Prompt 03 candidate is classified from the installed type library; none is runtime verified.",
    );
    output.push_str("| Canonical documentation ID | Typelib owner/interface | Member | DISPID | INVOKEKIND | Return type | Parameters | Optional | Enum dependencies | Audit result | Remaining runtime gate |\n| --- | --- | --- | ---: | --- | --- | ---: | ---: | --- | --- | --- |\n");
    for &candidate in candidate_ids() {
        let rows = data
            .members
            .iter()
            .filter(|record| {
                record
                    .get("canonical_documentation_id")
                    .and_then(Value::as_str)
                    == Some(candidate)
            })
            .collect::<Vec<_>>();
        if rows.is_empty() {
            output.push_str(&format!("| `{candidate}` | -- | -- | -- | -- | -- | -- | -- | -- | Missing from installed typelib | Runtime verification required if later resolved. |\n"));
            continue;
        }
        for row in rows {
            let enum_dependencies = enum_dependencies(row);
            output.push_str(&format!(
                "| `{candidate}` | `{}/{}` | `{}` | {} | `{}` | `{}` | {} | {} | {} | Typelib verified | Runtime verification required: declared metadata does not establish behavior. |\n",
                value_string(row, "owner"), value_string(row, "typelib_interface"), value_string(row, "name"), value_string(row, "dispid"), value_string(row, "invoke_kind"), type_label(row.get("return_type")), value_string(row, "parameter_count"), value_string(row, "optional_parameter_count"), if enum_dependencies.is_empty() { "--".to_owned() } else { enum_dependencies.join(", ") }
            ));
        }
    }
    output
}

fn collections_report(data: &AuditData) -> String {
    let mut output = report_header(
        "Collection contracts",
        "Collection metadata does not determine runtime indexing or enumeration behavior.",
    );
    output.push_str("| Collection | Member | DISPID | INVOKEKIND | Return type | Parameters | Hidden/restricted | Default member |\n| --- | --- | ---: | --- | --- | ---: | --- | --- |\n");
    for owner in [
        "Excel.Workbooks",
        "Excel.Worksheets",
        "Excel.Sheets",
        "Excel.Windows",
        "Excel.Names",
        "Excel.ListObjects",
        "Excel.ChartObjects",
        "Excel.Shapes",
    ] {
        let rows = data
            .members
            .iter()
            .filter(|record| record.get("owner").and_then(Value::as_str) == Some(owner));
        for row in rows {
            output.push_str(&format!(
                "| `{owner}` | `{}` | {} | `{}` | `{}` | {} | {}/{} | {} |\n",
                value_string(row, "name"),
                value_string(row, "dispid"),
                value_string(row, "invoke_kind"),
                type_label(row.get("return_type")),
                value_string(row, "parameter_count"),
                value_string(row, "hidden"),
                value_string(row, "restricted"),
                value_string(row, "default_member")
            ));
        }
    }
    output.push_str("\n`_NewEnum` and default-member evidence above is declared metadata only; runtime indexing, name lookup, and enumeration lifetime remain runtime gates.\n");
    output
}

fn optional_arguments_report(data: &AuditData) -> String {
    let mut output = report_header(
        "Optional-argument contracts",
        "Rows retain declared parameter order, flags, and default evidence without designing Rust options builders.",
    );
    output.push_str("| Member | Ordinal | Parameter | Type | Flags | Default | Required/optional | Enum | Runtime omission gate |\n| --- | ---: | --- | --- | --- | --- | --- | --- | --- |\n");
    for id in optional_member_ids() {
        let member_ids = data
            .members
            .iter()
            .filter(|member| {
                member
                    .get("canonical_documentation_id")
                    .and_then(Value::as_str)
                    == Some(id)
            })
            .filter_map(|member| member.get("id").and_then(Value::as_str))
            .collect::<BTreeSet<_>>();
        for parameter in data.parameters.iter().filter(|parameter| {
            parameter
                .get("member_id")
                .and_then(Value::as_str)
                .is_some_and(|value| member_ids.contains(value))
        }) {
            let enums = enum_dependencies(parameter);
            output.push_str(&format!("| `{id}` | {} | `{}` | `{}` | `{}` | {} | {} | {} | Runtime verification required. |\n", value_string(parameter, "ordinal"), value_string(parameter, "name"), type_label(parameter.get("type")), json_inline(parameter.get("param_flags")), json_inline(parameter.get("default")), if value_string(parameter, "optional") == "true" { "optional" } else { "required" }, if enums.is_empty() { "--".to_owned() } else { enums.join(", ") }));
        }
    }
    output
}

fn range_report(data: &AuditData) -> String {
    let mut output = report_header(
        "Range declarations",
        "Range VARIANT declarations do not establish the actual runtime VARTYPE or SAFEARRAY dimensions and bounds.",
    );
    output.push_str("| Canonical member | INVOKEKIND | Return type | Parameters | Optional | Default member | Runtime boundary |\n| --- | --- | --- | ---: | ---: | --- | --- |\n");
    for id in range_member_ids() {
        let rows = data.members.iter().filter(|member| {
            member
                .get("canonical_documentation_id")
                .and_then(Value::as_str)
                == Some(id)
        });
        for row in rows {
            output.push_str(&format!(
                "| `{id}` | `{}` | `{}` | {} | {} | {} | Runtime verification required. |\n",
                value_string(row, "invoke_kind"),
                type_label(row.get("return_type")),
                value_string(row, "parameter_count"),
                value_string(row, "optional_parameter_count"),
                value_string(row, "default_member")
            ));
        }
    }
    output
}

fn enum_report(data: &AuditData) -> String {
    let mut output = report_header(
        "Enum values",
        "Values are reflected from the installed library; no numeric Excel constants are hand-authored.",
    );
    output.push_str("| Enum | GUID | Values |\n| --- | --- | ---: |\n");
    for record in &data.enums {
        output.push_str(&format!(
            "| `{}` | `{}` | {} |\n",
            value_string(record, "name"),
            value_string(record, "guid"),
            record
                .get("values")
                .and_then(Value::as_array)
                .map_or(0, Vec::len)
        ));
    }
    output
}

fn documentation_differences(data: &AuditData) -> String {
    let mut output = report_header(
        "Documentation and typelib differences",
        "Differences are reported without changing the documentation-derived evidence layer.",
    );
    output.push_str("| Canonical ID | Result | Detail |\n| --- | --- | --- |\n");
    for &candidate in candidate_ids() {
        let docs = data.documentation.get(candidate);
        let rows = data
            .members
            .iter()
            .filter(|member| {
                member
                    .get("canonical_documentation_id")
                    .and_then(Value::as_str)
                    == Some(candidate)
            })
            .collect::<Vec<_>>();
        let selected = docs.and_then(|documentation| documented_member_row(documentation, &rows));
        let result = match (docs, selected) {
            (None, _) => "Documentation record missing",
            (Some(_), None) => "Missing compatible typelib member",
            (Some(documentation), Some(typelib)) => {
                let documented_count = documentation
                    .get("parameters")
                    .and_then(Value::as_array)
                    .map_or(0, Vec::len);
                let typelib_count = typelib
                    .get("parameter_count")
                    .and_then(Value::as_u64)
                    .unwrap_or_default() as usize;
                if documented_count == typelib_count {
                    "Signature shape agrees"
                } else {
                    "Parameter count differs"
                }
            }
        };
        let detail = match (docs, selected) {
            (Some(documentation), Some(typelib)) => {
                let documented_count = documentation
                    .get("parameters")
                    .and_then(Value::as_array)
                    .map_or(0, Vec::len);
                let documented_kind = value_string(documentation, "kind");
                format!(
                    "documentation {documented_kind} with {documented_count} parameter(s); typelib {} with {} parameter(s), DISPID {}.",
                    value_string(typelib, "invoke_kind"),
                    value_string(typelib, "parameter_count"),
                    value_string(typelib, "dispid")
                )
            }
            (Some(_), None) => "No selected interface member matched; aliases or version differences remain unresolved.".to_owned(),
            (None, Some(_)) => "Typelib member has no selected canonical documentation record.".to_owned(),
            (None, None) => "Neither evidence layer supplied this target.".to_owned(),
        };
        output.push_str(&format!("| `{candidate}` | {result} | {detail} |\n"));
    }
    output
}

fn documented_member_row<'a>(documentation: &Value, rows: &'a [&'a Value]) -> Option<&'a Value> {
    let documented_kind = documentation.get("kind").and_then(Value::as_str);
    let expected_invoke = match documented_kind {
        Some("property") => "INVOKE_PROPERTYGET",
        Some("method") => "INVOKE_FUNC",
        _ => return rows.first().copied(),
    };
    rows.iter()
        .copied()
        .find(|member| value_string(member, "invoke_kind") == expected_invoke)
}

fn unresolved_report(data: &AuditData) -> String {
    let mut output = report_header(
        "Unresolved typelib targets",
        "Unresolved entries are explicit audit follow-ups, not inferred ABI or runtime facts.",
    );
    output.push_str("| Category | Target | Detail |\n| --- | --- | --- |\n");
    if data.unresolved.is_empty() {
        output.push_str("| None | -- | No unresolved target was emitted. |\n");
    } else {
        for record in &data.unresolved {
            output.push_str(&format!(
                "| {} | `{}` | {} |\n",
                value_string(record, "category"),
                value_string(record, "target"),
                value_string(record, "detail")
            ));
        }
    }
    output
}

fn source_manifest(library: &Value, file: &FileEvidence, input: &AuditInput) -> String {
    format!(
        "schema_version = 1\naudit_tool_version = {TOOL_VERSION}\ninput_target_list = \"generated/analysis/typelib-audit-targets.md\"\nlibrary_guid = \"{}\"\nmajor_version = {}\nminor_version = {}\nlcid = {}\nsyskind = \"{}\"\nregistration_source = \"{}\"\nloaded_file_basename = \"{}\"\nloaded_file_sha256 = \"{}\"\nexcel_file_version = \"{}\"\nwindows_version = \"{}\"\noffice_bitness = \"{}\"\nwindows_crate_version = \"0.62.2\"\nwindows_core_version = \"0.62.2\"\ninspection_date = \"2026-07-21\"\n\n[audit_targets]\narchitectural_spine = [\"Excel.Application\", \"Excel.Workbooks\", \"Excel.Workbook\", \"Excel.Worksheets\", \"Excel.Worksheet\", \"Excel.Range\"]\ncollection_families = [\"Excel.Sheets\", \"Excel.Windows\", \"Excel.Names\", \"Excel.ListObjects\", \"Excel.ChartObjects\", \"Excel.Shapes\"]\ncandidate_0_1_count = 50\n",
        value_string(library, "guid"),
        value_string(library, "major_version"),
        value_string(library, "minor_version"),
        value_string(library, "lcid"),
        value_string(library, "syskind"),
        file.registration_category,
        file.basename,
        file.sha256,
        escape_toml(&input.excel_file_version),
        escape_toml(&input.windows_version),
        escape_toml(&input.office_bitness),
    )
}

fn summary(data: &AuditData) -> AuditSummary {
    AuditSummary {
        coclasses: data.coclasses.len(),
        interfaces: data.interfaces.len(),
        members: data.members.len(),
        parameters: data.parameters.len(),
        enums: data.enums.len(),
    }
}

fn sort_data(data: &mut AuditData) {
    for records in [
        &mut data.coclasses,
        &mut data.interfaces,
        &mut data.members,
        &mut data.parameters,
        &mut data.enums,
        &mut data.aliases,
        &mut data.unresolved,
    ] {
        sort_by_id(records);
    }
}

fn sort_by_id(records: &mut [Value]) {
    records.sort_by_key(|value| value_string(value, "id"));
}

fn jsonl(records: &[Value]) -> Result<String, String> {
    let mut output = String::new();
    for record in records {
        output.push_str(
            &serde_json::to_string(record)
                .map_err(|error| format!("cannot encode JSONL: {error}"))?,
        );
        output.push('\n');
    }
    if output.is_empty() {
        output.push('\n');
    }
    Ok(output)
}

fn json_text(value: &Value) -> Result<String, String> {
    let mut text = serde_json::to_string_pretty(value)
        .map_err(|error| format!("cannot encode JSON: {error}"))?;
    text.push('\n');
    Ok(text)
}

fn stable_id(kind: &str, identity: &str) -> String {
    format!("Excel.TypeLib.{kind}.{}", identity.replace(['#', ' '], "."))
}

fn targeted_interface(entry: &TypeEntry) -> bool {
    matches!(entry.kind, TKIND_INTERFACE | TKIND_DISPATCH)
        && matches!(
            entry.name.as_str(),
            "_Application"
                | "Workbooks"
                | "_Workbook"
                | "Worksheets"
                | "_Worksheet"
                | "Range"
                | "Sheets"
                | "Windows"
                | "Names"
                | "ListObjects"
                | "ChartObjects"
                | "Shapes"
        )
}

fn canonical_owner(interface: &str) -> &'static str {
    match interface {
        "Application" | "_Application" => "Excel.Application",
        "Workbooks" => "Excel.Workbooks",
        "Workbook" | "_Workbook" => "Excel.Workbook",
        "Worksheets" => "Excel.Worksheets",
        "Worksheet" | "_Worksheet" => "Excel.Worksheet",
        "Range" => "Excel.Range",
        "Sheets" => "Excel.Sheets",
        "Windows" => "Excel.Windows",
        "Names" => "Excel.Names",
        "ListObjects" => "Excel.ListObjects",
        "ChartObjects" => "Excel.ChartObjects",
        "Shapes" => "Excel.Shapes",
        _ => "Excel.Unknown",
    }
}

fn audited_member(owner: &str, name: &str) -> bool {
    let qualified = format!("{owner}.{name}");
    canonical_member_id(owner, name).is_some()
        || range_member_ids().contains(&qualified.as_str())
        || optional_member_ids().contains(&qualified.as_str())
        || matches!(
            name,
            "Count" | "Item" | "_Default" | "_NewEnum" | "Add" | "Delete" | "Remove"
        )
}

fn canonical_member_id(owner: &str, name: &str) -> Option<String> {
    candidate_ids()
        .iter()
        .find(|candidate| {
            candidate_member_name(candidate) == name && candidate_owner(candidate) == owner
        })
        .map(|candidate| (*candidate).to_owned())
}

fn candidate_owner(id: &str) -> &str {
    id.rsplit_once('.')
        .map_or("", |(owner, _)| owner)
        .trim_end_matches("#method")
}

fn candidate_member_name(id: &str) -> &str {
    id.rsplit('.')
        .next()
        .unwrap_or("")
        .trim_end_matches("#method")
}

fn candidate_ids() -> &'static [&'static str] {
    &[
        "Excel.Application.Visible",
        "Excel.Application.Version",
        "Excel.Application.Workbooks",
        "Excel.Application.ActiveWorkbook",
        "Excel.Application.ActiveSheet",
        "Excel.Application.Calculation",
        "Excel.Application.DisplayAlerts",
        "Excel.Application.ScreenUpdating",
        "Excel.Application.EnableEvents",
        "Excel.Application.Run",
        "Excel.Application.Quit",
        "Excel.Workbooks.Count",
        "Excel.Workbooks.Item",
        "Excel.Workbooks.Add",
        "Excel.Workbooks.Open",
        "Excel.Workbooks.Close",
        "Excel.Workbook.Name",
        "Excel.Workbook.FullName",
        "Excel.Workbook.Path",
        "Excel.Workbook.Worksheets",
        "Excel.Workbook.Save",
        "Excel.Workbook.SaveAs",
        "Excel.Workbook.Close",
        "Excel.Workbook.Saved",
        "Excel.Worksheets.Count",
        "Excel.Worksheets.Item",
        "Excel.Worksheets.Add",
        "Excel.Worksheet.Name",
        "Excel.Worksheet.Range",
        "Excel.Worksheet.Cells",
        "Excel.Worksheet.UsedRange",
        "Excel.Worksheet.Activate#method",
        "Excel.Range.Address",
        "Excel.Range.Value",
        "Excel.Range.Value2",
        "Excel.Range.Formula",
        "Excel.Range.Formula2",
        "Excel.Range.Rows",
        "Excel.Range.Columns",
        "Excel.Range.Offset",
        "Excel.Range.Resize",
        "Excel.Range.ClearContents",
        "Excel.Range.Clear",
        "Excel.Range.Delete",
        "Excel.Range.Copy",
        "Excel.Range.PasteSpecial",
        "Excel.Range.Select",
        "Excel.Range.Activate",
        "Excel.Range.Find",
        "Excel.Range.Sort",
    ]
}

fn optional_member_ids() -> &'static [&'static str] {
    &[
        "Excel.Workbooks.Open",
        "Excel.Workbook.SaveAs",
        "Excel.Workbook.Close",
        "Excel.Worksheets.Add",
        "Excel.Range.Find",
        "Excel.Range.Sort",
        "Excel.Application.Run",
        "Excel.Workbook.ExportAsFixedFormat",
        "Excel.Range.ExportAsFixedFormat",
    ]
}

fn range_member_ids() -> &'static [&'static str] {
    &[
        "Excel.Range.Value",
        "Excel.Range.Value2",
        "Excel.Range.Formula",
        "Excel.Range.Formula2",
        "Excel.Range.Text",
        "Excel.Range.HasFormula",
        "Excel.Range.Item",
        "Excel.Range.Cells",
        "Excel.Range.Rows",
        "Excel.Range.Columns",
        "Excel.Range.Offset",
        "Excel.Range.Resize",
        "Excel.Range.Address",
        "Excel.Range.ClearContents",
        "Excel.Range.Find",
        "Excel.Range.Sort",
        "Excel.Worksheet.Range",
        "Excel.Worksheet.Cells",
        "Excel.Worksheet.UsedRange",
    ]
}

fn type_documentation(info: &ITypeInfo) -> (String, String) {
    documentation_for_memid(info, -1)
}

fn library_documentation(library: &ITypeLib) -> (String, String) {
    let mut name = BSTR::new();
    let mut documentation = BSTR::new();
    let mut context = 0;
    let _ = unsafe {
        library.GetDocumentation(
            -1,
            Some(&mut name),
            Some(&mut documentation),
            &mut context,
            None,
        )
    };
    (name.to_string(), documentation.to_string())
}

fn documentation_for_memid(info: &ITypeInfo, memid: i32) -> (String, String) {
    let mut name = BSTR::new();
    let mut documentation = BSTR::new();
    let mut context = 0;
    let _ = unsafe {
        info.GetDocumentation(
            memid,
            Some(&mut name),
            Some(&mut documentation),
            &mut context,
            None,
        )
    };
    (name.to_string(), documentation.to_string())
}

fn guid_string(guid: &GUID) -> String {
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

fn invkind_name(kind: windows::Win32::System::Com::INVOKEKIND) -> &'static str {
    match kind {
        INVOKE_FUNC => "INVOKE_FUNC",
        INVOKE_PROPERTYGET => "INVOKE_PROPERTYGET",
        INVOKE_PROPERTYPUT => "INVOKE_PROPERTYPUT",
        INVOKE_PROPERTYPUTREF => "INVOKE_PROPERTYPUTREF",
        _ => "INVOKE_UNKNOWN",
    }
}

fn calling_convention(value: i32) -> &'static str {
    if value == CC_STDCALL.0 {
        "CC_STDCALL"
    } else {
        "CALLCONV_UNKNOWN"
    }
}

fn typekind_name(kind: TYPEKIND) -> &'static str {
    match kind {
        TKIND_ENUM => "TKIND_ENUM",
        TKIND_INTERFACE => "TKIND_INTERFACE",
        TKIND_DISPATCH => "TKIND_DISPATCH",
        TKIND_COCLASS => "TKIND_COCLASS",
        TKIND_ALIAS => "TKIND_ALIAS",
        _ => "TKIND_OTHER",
    }
}

fn syskind_name(value: i32) -> &'static str {
    match value {
        1 => "SYS_WIN32",
        3 => "SYS_WIN64",
        _ => "SYSKIND_OTHER",
    }
}

fn parameter_flags(flags: u16) -> Vec<&'static str> {
    let mut names = Vec::new();
    if flags & PARAMFLAG_FIN.0 != 0 {
        names.push("PARAMFLAG_FIN");
    }
    if flags & PARAMFLAG_FOUT.0 != 0 {
        names.push("PARAMFLAG_FOUT");
    }
    if flags & PARAMFLAG_FRETVAL.0 != 0 {
        names.push("PARAMFLAG_FRETVAL");
    }
    if flags & PARAMFLAG_FOPT.0 != 0 {
        names.push("PARAMFLAG_FOPT");
    }
    if flags & PARAMFLAG_FHASDEFAULT.0 != 0 {
        names.push("PARAMFLAG_FHASDEFAULT");
    }
    names
}

fn function_flags(flags: u16) -> Vec<&'static str> {
    let mut names = Vec::new();
    if flags & FUNCFLAG_FDEFAULTBIND.0 != 0 {
        names.push("FUNCFLAG_FDEFAULTBIND");
    }
    if flags & FUNCFLAG_FDEFAULTCOLLELEM.0 != 0 {
        names.push("FUNCFLAG_FDEFAULTCOLLELEM");
    }
    if flags & FUNCFLAG_FHIDDEN.0 != 0 {
        names.push("FUNCFLAG_FHIDDEN");
    }
    if flags & FUNCFLAG_FRESTRICTED.0 != 0 {
        names.push("FUNCFLAG_FRESTRICTED");
    }
    if flags & FUNCFLAG_FNONBROWSABLE.0 != 0 {
        names.push("FUNCFLAG_FNONBROWSABLE");
    }
    names
}

fn value_string(value: &Value, field: &str) -> String {
    value.get(field).map_or_else(
        || "--".to_owned(),
        |field| match field {
            Value::String(value) => value.clone(),
            Value::Null => "--".to_owned(),
            other => other.to_string(),
        },
    )
}

fn json_inline(value: Option<&Value>) -> String {
    value
        .map_or_else(|| "--".to_owned(), Value::to_string)
        .replace('|', "\\|")
}

fn type_label(value: Option<&Value>) -> String {
    value.and_then(Value::as_object).map_or_else(
        || "--".to_owned(),
        |value| {
            value
                .get("kind")
                .and_then(Value::as_str)
                .unwrap_or("unknown")
                .to_owned()
        },
    )
}

fn enum_dependencies(value: &Value) -> Vec<String> {
    let mut enums = BTreeSet::new();
    let mut aliases = BTreeSet::new();
    collect_user_types(value, &mut enums, &mut aliases);
    enums
        .into_iter()
        .map(|name| format!("`Excel.{name}`"))
        .collect()
}

fn escape_toml(value: &str) -> String {
    value.replace('\\', "\\\\").replace('"', "\\\"")
}

struct Apartment;

impl Apartment {
    fn initialize() -> Result<Self, String> {
        unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok() }.map_err(|error| {
            format!("cannot initialize COM for type-library inspection: {error}")
        })?;
        Ok(Self)
    }
}

impl Drop for Apartment {
    fn drop(&mut self) {
        unsafe { CoUninitialize() };
    }
}

struct TypeAttr<'a> {
    info: &'a ITypeInfo,
    pointer: *mut TYPEATTR,
}
impl<'a> TypeAttr<'a> {
    fn acquire(info: &'a ITypeInfo) -> Result<Self, String> {
        let pointer = unsafe { info.GetTypeAttr() }
            .map_err(|error| format!("cannot acquire TYPEATTR: {error}"))?;
        if pointer.is_null() {
            return Err("ITypeInfo returned a null TYPEATTR".to_owned());
        }
        Ok(Self { info, pointer })
    }
}
impl Drop for TypeAttr<'_> {
    fn drop(&mut self) {
        unsafe { self.info.ReleaseTypeAttr(self.pointer) }
    }
}

struct FuncDesc<'a> {
    info: &'a ITypeInfo,
    pointer: *mut FUNCDESC,
}
impl<'a> FuncDesc<'a> {
    fn acquire(info: &'a ITypeInfo, index: u32) -> Result<Self, String> {
        let pointer = unsafe { info.GetFuncDesc(index) }
            .map_err(|error| format!("cannot acquire FUNCDESC {index}: {error}"))?;
        if pointer.is_null() {
            return Err(format!("ITypeInfo returned a null FUNCDESC for {index}"));
        }
        Ok(Self { info, pointer })
    }
}
impl Drop for FuncDesc<'_> {
    fn drop(&mut self) {
        unsafe { self.info.ReleaseFuncDesc(self.pointer) }
    }
}

struct VarDesc<'a> {
    info: &'a ITypeInfo,
    pointer: *mut windows::Win32::System::Com::VARDESC,
}
impl<'a> VarDesc<'a> {
    fn acquire(info: &'a ITypeInfo, index: u32) -> Result<Self, String> {
        let pointer = unsafe { info.GetVarDesc(index) }
            .map_err(|error| format!("cannot acquire VARDESC {index}: {error}"))?;
        if pointer.is_null() {
            return Err(format!("ITypeInfo returned a null VARDESC for {index}"));
        }
        Ok(Self { info, pointer })
    }
}
impl Drop for VarDesc<'_> {
    fn drop(&mut self) {
        unsafe { self.info.ReleaseVarDesc(self.pointer) }
    }
}

struct TypeLibAttr<'a> {
    library: &'a ITypeLib,
    pointer: *mut windows::Win32::System::Com::TLIBATTR,
}
impl<'a> TypeLibAttr<'a> {
    fn acquire(library: &'a ITypeLib) -> Result<Self, String> {
        let pointer = unsafe { library.GetLibAttr() }
            .map_err(|error| format!("cannot acquire TLIBATTR: {error}"))?;
        if pointer.is_null() {
            return Err("ITypeLib returned a null TLIBATTR".to_owned());
        }
        Ok(Self { library, pointer })
    }
}
impl Drop for TypeLibAttr<'_> {
    fn drop(&mut self) {
        unsafe { self.library.ReleaseTLibAttr(self.pointer) }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn candidate_scope_is_stable_and_complete() {
        assert_eq!(candidate_ids().len(), 50);
        assert_eq!(candidate_ids()[0], "Excel.Application.Visible");
        assert_eq!(candidate_ids()[49], "Excel.Range.Sort");
    }

    #[test]
    fn primitive_types_keep_structured_vartype_evidence() {
        assert_eq!(primitive_type(VT_VARIANT)["kind"], "variant");
        assert_eq!(primitive_type(VT_VARIANT)["vartype_raw"], 12);
        assert_eq!(primitive_type(VT_BSTR)["kind"], "bstr");
    }

    #[test]
    fn portable_manifest_escapes_machine_text() {
        assert_eq!(escape_toml("A\\B\"C"), "A\\\\B\\\"C");
    }

    #[test]
    fn empty_jsonl_still_has_a_final_lf() {
        assert_eq!(jsonl(&[]).expect("empty JSONL is valid"), "\n");
    }
}
