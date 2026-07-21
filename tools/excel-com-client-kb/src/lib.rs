#![forbid(unsafe_code)]

use serde::Deserialize;
use serde_json::{Value, json};
use std::collections::{BTreeMap, BTreeSet};
use std::fs;
use std::path::{Path, PathBuf};
use std::process::Command;

const RETRIEVED_ON: &str = "2026-07-21";
const PYWIN32_COMMIT: &str = "a992023bd2d2ef57f8b605b43c1bcc29cdc619e9";
const PYWIN32_311_COMMIT: &str = "8b328dffac71b7afaf2d72f47c4048f27a32f6c8";
const PYWIN32_312_COMMIT: &str = "2a277cb5552756c2b4d42b524dc36d25e0bb6354";
const COMTYPES_COMMIT: &str = "339ea278d85defda3d3c0dba118969021018e5fb";
const MOJIBAKE_PATTERNS: &[&str] = &["â", "ï¿½", "\u{FFFD}"];

#[derive(Debug, Clone, Copy)]
struct CorpusMember {
    owner: &'static str,
    name: &'static str,
    invoke_kind: &'static str,
    expression: &'static str,
    family: &'static str,
    mechanism: &'static str,
}

impl CorpusMember {
    fn canonical(self) -> String {
        format!("{}.{}", self.owner, self.name)
    }

    fn typelib_id(self) -> String {
        format!(
            "Excel.TypeLib.member.{}.{}",
            self.canonical(),
            self.invoke_kind
        )
    }
}

const CORPUS: &[CorpusMember] = &[
    CorpusMember {
        owner: "Excel.Application",
        name: "Version",
        invoke_kind: "INVOKE_PROPERTYGET",
        expression: "excel.Version",
        family: "Application",
        mechanism: "scalar property get",
    },
    CorpusMember {
        owner: "Excel.Application",
        name: "Hwnd",
        invoke_kind: "INVOKE_PROPERTYGET",
        expression: "excel.Hwnd",
        family: "Application",
        mechanism: "scalar property get",
    },
    CorpusMember {
        owner: "Excel.Application",
        name: "Visible",
        invoke_kind: "INVOKE_PROPERTYGET",
        expression: "excel.Visible",
        family: "Application",
        mechanism: "scalar property get",
    },
    CorpusMember {
        owner: "Excel.Application",
        name: "Visible",
        invoke_kind: "INVOKE_PROPERTYPUT",
        expression: "excel.Visible = False",
        family: "Application",
        mechanism: "property put",
    },
    CorpusMember {
        owner: "Excel.Application",
        name: "Workbooks",
        invoke_kind: "INVOKE_PROPERTYGET",
        expression: "excel.Workbooks",
        family: "Application",
        mechanism: "object property get",
    },
    CorpusMember {
        owner: "Excel.Application",
        name: "Run",
        invoke_kind: "INVOKE_FUNC",
        expression: "excel.Run(macro, *args)",
        family: "Application",
        mechanism: "large optional signature",
    },
    CorpusMember {
        owner: "Excel.Application",
        name: "Quit",
        invoke_kind: "INVOKE_FUNC",
        expression: "excel.Quit()",
        family: "Application",
        mechanism: "zero-argument method",
    },
    CorpusMember {
        owner: "Excel.Workbooks",
        name: "Count",
        invoke_kind: "INVOKE_PROPERTYGET",
        expression: "workbooks.Count",
        family: "Workbooks",
        mechanism: "collection count",
    },
    CorpusMember {
        owner: "Excel.Workbooks",
        name: "Item",
        invoke_kind: "INVOKE_PROPERTYGET",
        expression: "workbooks.Item(index)",
        family: "Workbooks",
        mechanism: "default member / collection Item",
    },
    CorpusMember {
        owner: "Excel.Workbooks",
        name: "Add",
        invoke_kind: "INVOKE_FUNC",
        expression: "workbooks.Add()",
        family: "Workbooks",
        mechanism: "zero-argument method",
    },
    CorpusMember {
        owner: "Excel.Workbooks",
        name: "Open",
        invoke_kind: "INVOKE_FUNC",
        expression: "workbooks.Open(filename)",
        family: "Workbooks",
        mechanism: "required-argument method",
    },
    CorpusMember {
        owner: "Excel.Workbook",
        name: "Name",
        invoke_kind: "INVOKE_PROPERTYGET",
        expression: "workbook.Name",
        family: "Workbook",
        mechanism: "scalar property get",
    },
    CorpusMember {
        owner: "Excel.Workbook",
        name: "Worksheets",
        invoke_kind: "INVOKE_PROPERTYGET",
        expression: "workbook.Worksheets",
        family: "Workbook",
        mechanism: "object property get",
    },
    CorpusMember {
        owner: "Excel.Workbook",
        name: "Save",
        invoke_kind: "INVOKE_FUNC",
        expression: "workbook.Save()",
        family: "Workbook",
        mechanism: "zero-argument method",
    },
    CorpusMember {
        owner: "Excel.Workbook",
        name: "SaveAs",
        invoke_kind: "INVOKE_FUNC",
        expression: "workbook.SaveAs(filename, **options)",
        family: "Workbook",
        mechanism: "large optional signature / named arguments",
    },
    CorpusMember {
        owner: "Excel.Workbook",
        name: "Close",
        invoke_kind: "INVOKE_FUNC",
        expression: "workbook.Close(SaveChanges=...)",
        family: "Workbook",
        mechanism: "optional-argument method",
    },
    CorpusMember {
        owner: "Excel.Worksheets",
        name: "Count",
        invoke_kind: "INVOKE_PROPERTYGET",
        expression: "worksheets.Count",
        family: "Worksheets",
        mechanism: "collection count",
    },
    CorpusMember {
        owner: "Excel.Worksheets",
        name: "Item",
        invoke_kind: "INVOKE_PROPERTYGET",
        expression: "worksheets.Item(index)",
        family: "Worksheets",
        mechanism: "default member / collection Item",
    },
    CorpusMember {
        owner: "Excel.Worksheets",
        name: "Add",
        invoke_kind: "INVOKE_FUNC",
        expression: "worksheets.Add()",
        family: "Worksheets",
        mechanism: "method returning object",
    },
    CorpusMember {
        owner: "Excel.Worksheet",
        name: "Name",
        invoke_kind: "INVOKE_PROPERTYGET",
        expression: "worksheet.Name",
        family: "Worksheet",
        mechanism: "scalar property get",
    },
    CorpusMember {
        owner: "Excel.Worksheet",
        name: "Name",
        invoke_kind: "INVOKE_PROPERTYPUT",
        expression: "worksheet.Name = value",
        family: "Worksheet",
        mechanism: "property put",
    },
    CorpusMember {
        owner: "Excel.Worksheet",
        name: "Range",
        invoke_kind: "INVOKE_PROPERTYGET",
        expression: "worksheet.Range(cell1, cell2)",
        family: "Worksheet",
        mechanism: "property get with arguments",
    },
    CorpusMember {
        owner: "Excel.Worksheet",
        name: "Cells",
        invoke_kind: "INVOKE_PROPERTYGET",
        expression: "worksheet.Cells",
        family: "Worksheet",
        mechanism: "object property get",
    },
    CorpusMember {
        owner: "Excel.Worksheet",
        name: "UsedRange",
        invoke_kind: "INVOKE_PROPERTYGET",
        expression: "worksheet.UsedRange",
        family: "Worksheet",
        mechanism: "object property get",
    },
    CorpusMember {
        owner: "Excel.Range",
        name: "Item",
        invoke_kind: "INVOKE_PROPERTYGET",
        expression: "range.Item(row, column)",
        family: "Range",
        mechanism: "default member / property get with arguments",
    },
    CorpusMember {
        owner: "Excel.Range",
        name: "Value",
        invoke_kind: "INVOKE_PROPERTYGET",
        expression: "range.Value",
        family: "Range",
        mechanism: "scalar VARIANT / rectangular array result",
    },
    CorpusMember {
        owner: "Excel.Range",
        name: "Value",
        invoke_kind: "INVOKE_PROPERTYPUT",
        expression: "range.Value = value",
        family: "Range",
        mechanism: "property put with named DISPID",
    },
    CorpusMember {
        owner: "Excel.Range",
        name: "Value2",
        invoke_kind: "INVOKE_PROPERTYGET",
        expression: "range.Value2",
        family: "Range",
        mechanism: "scalar VARIANT / rectangular array result",
    },
    CorpusMember {
        owner: "Excel.Range",
        name: "Value2",
        invoke_kind: "INVOKE_PROPERTYPUT",
        expression: "range.Value2 = value",
        family: "Range",
        mechanism: "property put with named DISPID",
    },
    CorpusMember {
        owner: "Excel.Range",
        name: "Formula",
        invoke_kind: "INVOKE_PROPERTYGET",
        expression: "range.Formula",
        family: "Range",
        mechanism: "formula transport",
    },
    CorpusMember {
        owner: "Excel.Range",
        name: "Formula",
        invoke_kind: "INVOKE_PROPERTYPUT",
        expression: "range.Formula = formula",
        family: "Range",
        mechanism: "property put with named DISPID",
    },
    CorpusMember {
        owner: "Excel.Range",
        name: "Formula2",
        invoke_kind: "INVOKE_PROPERTYGET",
        expression: "range.Formula2",
        family: "Range",
        mechanism: "formula transport",
    },
    CorpusMember {
        owner: "Excel.Range",
        name: "Formula2",
        invoke_kind: "INVOKE_PROPERTYPUT",
        expression: "range.Formula2 = formula",
        family: "Range",
        mechanism: "property put with named DISPID",
    },
    CorpusMember {
        owner: "Excel.Range",
        name: "Text",
        invoke_kind: "INVOKE_PROPERTYGET",
        expression: "range.Text",
        family: "Range",
        mechanism: "scalar VARIANT result",
    },
    CorpusMember {
        owner: "Excel.Range",
        name: "Address",
        invoke_kind: "INVOKE_PROPERTYGET",
        expression: "range.Address()",
        family: "Range",
        mechanism: "optional-argument property get",
    },
    CorpusMember {
        owner: "Excel.Range",
        name: "Offset",
        invoke_kind: "INVOKE_PROPERTYGET",
        expression: "range.Offset(row_offset, column_offset)",
        family: "Range",
        mechanism: "optional-argument property get",
    },
    CorpusMember {
        owner: "Excel.Range",
        name: "Resize",
        invoke_kind: "INVOKE_PROPERTYGET",
        expression: "range.Resize(row_size, column_size)",
        family: "Range",
        mechanism: "optional-argument property get",
    },
    CorpusMember {
        owner: "Excel.Range",
        name: "ClearContents",
        invoke_kind: "INVOKE_FUNC",
        expression: "range.ClearContents()",
        family: "Range",
        mechanism: "zero-argument method",
    },
    CorpusMember {
        owner: "Excel.Range",
        name: "Find",
        invoke_kind: "INVOKE_FUNC",
        expression: "range.Find(what, **options)",
        family: "Range",
        mechanism: "optional arguments / named arguments",
    },
    CorpusMember {
        owner: "Excel.Range",
        name: "Sort",
        invoke_kind: "INVOKE_FUNC",
        expression: "range.Sort(**options)",
        family: "Range",
        mechanism: "large optional signature / named arguments",
    },
    CorpusMember {
        owner: "Excel.Workbooks",
        name: "_NewEnum",
        invoke_kind: "INVOKE_PROPERTYGET",
        expression: "iter(workbooks)",
        family: "Workbooks",
        mechanism: "enumeration",
    },
    CorpusMember {
        owner: "Excel.Worksheets",
        name: "_NewEnum",
        invoke_kind: "INVOKE_PROPERTYGET",
        expression: "iter(worksheets)",
        family: "Worksheets",
        mechanism: "enumeration",
    },
];

#[derive(Debug, Deserialize)]
struct MemberMeta {
    dispid: i32,
    invoke_kind: String,
    parameter_count: usize,
    optional_parameter_count: usize,
    return_type: Value,
    #[serde(default)]
    source: String,
}

#[derive(Debug, Deserialize)]
struct ParameterMeta {
    ordinal: usize,
    name: String,
    optional: bool,
    required: bool,
    #[serde(rename = "type")]
    parameter_type: Value,
    default: Value,
}

#[derive(Debug)]
struct TypelibData {
    members: BTreeMap<String, MemberMeta>,
    parameters: BTreeMap<String, Vec<ParameterMeta>>,
}

#[derive(Debug, Clone, Copy)]
enum ClientMode {
    Pywin32Dynamic,
    Pywin32Generated,
    ComtypesDynamic,
    ComtypesGenerated,
}

impl ClientMode {
    const ALL: [Self; 4] = [
        Self::Pywin32Dynamic,
        Self::Pywin32Generated,
        Self::ComtypesDynamic,
        Self::ComtypesGenerated,
    ];

    fn client(self) -> &'static str {
        match self {
            Self::Pywin32Dynamic | Self::Pywin32Generated => "pywin32",
            Self::ComtypesDynamic | Self::ComtypesGenerated => "comtypes",
        }
    }

    fn mode(self, member: CorpusMember) -> &'static str {
        match self {
            Self::Pywin32Dynamic => "dynamic",
            Self::Pywin32Generated => "makepy-generated",
            Self::ComtypesDynamic => "dynamic dispatch",
            Self::ComtypesGenerated if uses_comtypes_vtable(member) => {
                "generated dual-interface/vtable"
            }
            Self::ComtypesGenerated => "generated dispinterface",
        }
    }

    fn stable_name(self) -> &'static str {
        match self {
            Self::Pywin32Dynamic => "pywin32.dynamic",
            Self::Pywin32Generated => "pywin32.makepy",
            Self::ComtypesDynamic => "comtypes.dynamic",
            Self::ComtypesGenerated => "comtypes.generated",
        }
    }
}

pub struct Summary {
    pub records: usize,
    pub reports: usize,
    pub typelib_joins: usize,
}

pub fn generate(root: &Path) -> Result<Summary, String> {
    let rendered = render(root)?;
    for (relative, content) in &rendered.files {
        let path = root.join(relative);
        if let Some(parent) = path.parent() {
            fs::create_dir_all(parent).map_err(io_error)?;
        }
        fs::write(path, content).map_err(io_error)?;
    }
    Ok(Summary {
        records: rendered.records,
        reports: rendered.reports,
        typelib_joins: CORPUS.len(),
    })
}

pub fn check(root: &Path) -> Result<Summary, String> {
    let rendered = render(root)?;
    for (relative, expected) in &rendered.files {
        reject_mojibake(expected, relative)?;
        let path = root.join(relative);
        let actual = fs::read_to_string(&path).map_err(|_| {
            format!(
                "missing required generated or evidence file: {}",
                path.display()
            )
        })?;
        let actual = normalize_line_endings(&actual);
        if actual != *expected {
            return Err(format!(
                "non-deterministic or hand-edited output: {}",
                path.display()
            ));
        }
        if actual.contains(":\\") {
            return Err(format!(
                "portable-path validation failed: {}",
                path.display()
            ));
        }
        if !actual.ends_with('\n') {
            return Err(format!("missing final newline: {}", path.display()));
        }
    }
    Ok(Summary {
        records: rendered.records,
        reports: rendered.reports,
        typelib_joins: CORPUS.len(),
    })
}

fn normalize_line_endings(text: &str) -> String {
    text.replace("\r\n", "\n")
}

fn reject_mojibake(text: &str, relative: &Path) -> Result<(), String> {
    if let Some(pattern) = MOJIBAKE_PATTERNS
        .iter()
        .find(|pattern| text.contains(**pattern))
    {
        return Err(format!(
            "mojibake pattern {pattern:?} in generated output: {}",
            relative.display()
        ));
    }
    Ok(())
}

pub fn diagnose(root: &Path, mode: &str, python: Option<&str>) -> Result<(), String> {
    let repository = root
        .parent()
        .and_then(Path::parent)
        .ok_or_else(|| "knowledge root must be nested under the repository".to_owned())?;
    let (script, argument) = match mode {
        "pywin32-dynamic" => ("pywin32_diagnostic.py", Some("dynamic")),
        "pywin32-generated" => ("pywin32_diagnostic.py", Some("generated")),
        "comtypes" => ("comtypes_diagnostic.py", Some("dynamic")),
        "comtypes-generated" => ("comtypes_diagnostic.py", Some("generated")),
        _ => {
            return Err(
                "diagnostic mode must be pywin32-dynamic, pywin32-generated, comtypes, or comtypes-generated"
                    .to_owned(),
            );
        }
    };
    let script_path = repository
        .join("tools/excel-com-client-kb/scripts")
        .join(script);
    let mut command = Command::new(python.unwrap_or("py"));
    if python.is_none() {
        command.arg("-3.11");
    }
    command.arg(script_path);
    if let Some(argument) = argument {
        command.args(["--mode", argument]);
    }
    let output = command.output().map_err(io_error)?;
    if !output.status.success() {
        return Err(format!("opt-in {mode} diagnostic did not complete"));
    }
    let stdout = String::from_utf8(output.stdout)
        .map_err(|_| "diagnostic output was not UTF-8".to_owned())?;
    let value: Value = serde_json::from_str(stdout.trim())
        .map_err(|_| "diagnostic output was not JSON".to_owned())?;
    if value.to_string().contains(":\\") {
        return Err("diagnostic unexpectedly contained a local path".to_owned());
    }
    println!(
        "{}",
        serde_json::to_string_pretty(&value).map_err(|error| error.to_string())?
    );
    Ok(())
}

struct Rendered {
    files: BTreeMap<PathBuf, String>,
    records: usize,
    reports: usize,
}

fn render(root: &Path) -> Result<Rendered, String> {
    let typelib = load_typelib(root)?;
    validate_corpus(&typelib)?;
    let mut files = BTreeMap::new();
    insert_client_files(&mut files, "pywin32", &typelib)?;
    insert_client_files(&mut files, "comtypes", &typelib)?;
    insert_reports(&mut files, &typelib)?;
    Ok(Rendered {
        files,
        records: CORPUS.len() * ClientMode::ALL.len() + 37,
        reports: 16,
    })
}

fn load_typelib(root: &Path) -> Result<TypelibData, String> {
    let mut members = BTreeMap::new();
    for line in read_json_lines(&root.join("typelib/members.jsonl"))? {
        let value: Value = serde_json::from_str(&line).map_err(|error| error.to_string())?;
        let id = required_string(&value, "id")?;
        let meta: MemberMeta = serde_json::from_value(value).map_err(|error| error.to_string())?;
        members.insert(id, meta);
    }
    let mut parameters: BTreeMap<String, Vec<ParameterMeta>> = BTreeMap::new();
    for line in read_json_lines(&root.join("typelib/parameters.jsonl"))? {
        let value: Value = serde_json::from_str(&line).map_err(|error| error.to_string())?;
        let member_id = required_string(&value, "member_id")?;
        let meta: ParameterMeta =
            serde_json::from_value(value).map_err(|error| error.to_string())?;
        parameters.entry(member_id).or_default().push(meta);
    }
    for values in parameters.values_mut() {
        values.sort_by_key(|parameter| parameter.ordinal);
    }
    // Application.Hwnd is mandatory for this prompt but is absent from the
    // prior curated member extract. makepy generated code from the installed
    // Excel 1.9 typelib supplies this descriptor; retain that distinct source.
    members
        .entry("Excel.TypeLib.member.Excel.Application.Hwnd.INVOKE_PROPERTYGET".to_owned())
        .or_insert(MemberMeta {
            dispid: 1950,
            invoke_kind: "INVOKE_PROPERTYGET".to_owned(),
            parameter_count: 0,
            optional_parameter_count: 0,
            return_type: json!({"kind":"primitive","vartype":"VT_I4","vartype_raw":3}),
            source: "generated-pywin32-wrapper-from-installed-excel-typelib".to_owned(),
        });
    Ok(TypelibData {
        members,
        parameters,
    })
}

fn read_json_lines(path: &Path) -> Result<Vec<String>, String> {
    let content = fs::read_to_string(path).map_err(io_error)?;
    Ok(content
        .lines()
        .filter(|line| !line.is_empty())
        .map(ToOwned::to_owned)
        .collect())
}

fn required_string(value: &Value, key: &str) -> Result<String, String> {
    value
        .get(key)
        .and_then(Value::as_str)
        .map(ToOwned::to_owned)
        .ok_or_else(|| format!("missing string field {key}"))
}

fn validate_corpus(typelib: &TypelibData) -> Result<(), String> {
    let mut ids = BTreeSet::new();
    for member in CORPUS {
        let id = member.typelib_id();
        if !ids.insert(id.clone()) {
            return Err(format!("duplicate corpus member {id}"));
        }
        let metadata = typelib
            .members
            .get(&id)
            .ok_or_else(|| format!("typelib correlation missing for {id}"))?;
        if metadata.invoke_kind != member.invoke_kind {
            return Err(format!("invoke-kind mismatch for {id}"));
        }
    }
    if CORPUS.len() != 42 {
        return Err("the mandatory corpus must contain 42 invocation forms".to_owned());
    }
    Ok(())
}

fn insert_client_files(
    files: &mut BTreeMap<PathBuf, String>,
    client: &str,
    typelib: &TypelibData,
) -> Result<(), String> {
    let base = PathBuf::from("client-implementations").join(client);
    files.insert(base.join("SOURCE_MANIFEST.toml"), source_manifest(client)?);
    files.insert(base.join("environments.jsonl"), environments(client)?);
    files.insert(base.join("activation.jsonl"), activation(client)?);
    files.insert(base.join("wrappers.jsonl"), wrappers(client)?);
    files.insert(
        base.join("invocations.jsonl"),
        invocations(client, typelib)?,
    );
    files.insert(base.join("conversions.jsonl"), conversions(client)?);
    files.insert(base.join("errors.jsonl"), errors(client)?);
    files.insert(base.join("patterns.jsonl"), patterns(client)?);
    if client == "pywin32" {
        files.insert(
            base.join("version-reconciliation.jsonl"),
            pywin32_version_reconciliation()?,
        );
    }
    Ok(())
}

fn source_manifest(client: &str) -> Result<String, String> {
    let content = match client {
        "pywin32" => format!(
            "schema_version = 1\nclient = \"pywin32\"\nrepository = \"mhammond/pywin32\"\ncommit = \"{PYWIN32_COMMIT}\"\nretrieved_on = \"{RETRIEVED_ON}\"\nsource_version = \"312.1\"\nlicense = \"BSD-3-Clause\"\n\n[reconciliation]\ninstalled_package_version = \"311\"\ninstalled_source_tag = \"b311\"\ninstalled_source_commit = \"{PYWIN32_311_COMMIT}\"\nreleased_comparison_tag = \"b312\"\nreleased_comparison_commit = \"{PYWIN32_312_COMMIT}\"\ninspected_reference_commit = \"{PYWIN32_COMMIT}\"\nstatus = \"complete for the selected Excel Automation paths\"\ninstalled_version_parity = \"source-confirmed for the selected paths; not byte-identical to 312.1\"\nsource_build_312_1 = \"not installed: isolated build exceeded the bounded diagnostic interval\"\n\nfiles_inspected = [\n  \"com/win32com/client/__init__.py\",\n  \"com/win32com/client/dynamic.py\",\n  \"com/win32com/client/build.py\",\n  \"com/win32com/client/genpy.py\",\n  \"com/win32com/client/gencache.py\",\n  \"com/win32com/src/PyIDispatch.cpp\",\n  \"com/win32com/src/oleargs.cpp\",\n  \"com/win32com/src/ErrorUtils.cpp\",\n  \"com/win32com/src/PythonCOM.cpp\",\n]\n"
        ),
        "comtypes" => format!(
            "schema_version = 1\nclient = \"comtypes\"\nrepository = \"enthought/comtypes\"\ncommit = \"{COMTYPES_COMMIT}\"\nretrieved_on = \"{RETRIEVED_ON}\"\nsource_version = \"1.4.16\"\nlicense = \"MIT\"\nlocally_installed_version = \"1.4.16\"\ninstalled_version_matches_inspected_source = true\ninstalled_exact_commit_match = \"unverified: wheel metadata carries no source commit\"\nfiles_inspected = [\n  \"comtypes/__init__.py\",\n  \"comtypes/automation.py\",\n  \"comtypes/safearray.py\",\n  \"comtypes/_memberspec.py\",\n  \"comtypes/client/__init__.py\",\n  \"comtypes/client/_create.py\",\n  \"comtypes/client/_activeobj.py\",\n  \"comtypes/client/_managing.py\",\n  \"comtypes/client/dynamic.py\",\n  \"comtypes/client/lazybind.py\",\n  \"comtypes/client/_generate.py\",\n  \"comtypes/tools/tlbparser.py\",\n]\n"
        ),
        _ => return Err(format!("unknown client {client}")),
    };
    Ok(content)
}

fn environments(client: &str) -> Result<String, String> {
    let values = match client {
        "pywin32" => vec![
            json!({"id":"client.pywin32.05d-env-a","classification":"Control-confirmed","environment":"A","python_version":"3.11.7","interpreter_architecture":"x64","package_architecture":"win_amd64","pywin32_version":"311","source_tag":"b311","source_commit":PYWIN32_311_COMMIT,"excel_version":"16.0","office_bitness":"64-bit","wrapper_modes":["dynamic","makepy-generated"],"generated_cache":"isolated per local control mode; path not committed","com_initialization":{"module":"pythoncom","main_thread":"automatic","sys_coinitialization_flags":null,"source":"PyCOM module initialization reads sys.coinit_flags or defaults to COINIT_APARTMENTTHREADED"},"reconciliation":"source-confirmed for selected Automation paths","raw_paths_recorded":false}),
            json!({"id":"client.pywin32.05d-env-b","classification":"Control-confirmed","environment":"B","python_version":"3.11.7","interpreter_architecture":"x64","package_architecture":"win_amd64","pywin32_version":"312","source_tag":"b312","source_commit":PYWIN32_312_COMMIT,"inspected_reference_version":"312.1","inspected_reference_commit":PYWIN32_COMMIT,"excel_version":"16.0","office_bitness":"64-bit","wrapper_modes":["dynamic","makepy-generated"],"generated_cache":"isolated per local control mode; path not committed","source_build_312_1":"attempted in an isolated environment; bounded build did not complete, no package was retained","reconciliation":"b312 has no material selected Automation-path difference from inspected 312.1 source","raw_paths_recorded":false}),
            json!({"id":"client.pywin32.local-python-311","classification":"Control-confirmed","python_version":"3.11.7","pywin32_version":"311","module_identities":{"win32com":"win32com","pythoncom":"pythoncom"},"com_initialization":{"module":"pythoncom","main_thread":"automatic","sys_coinitialization_flags":null,"source":"PyCOM module initialization reads sys.coinit_flags or defaults to COINIT_APARTMENTTHREADED"},"installed_matches_source_manifest":"reconciled selected paths","raw_paths_recorded":false}),
            json!({"id":"client.pywin32.dynamic-05c","classification":"Inconclusive","mode":"dynamic","wrapper_classes":{"Application":{"class":"CDispatch","module":"win32com.client.dynamic"},"Workbooks":{"class":"CDispatch","module":"win32com.client.dynamic"},"Workbook":"not returned"},"activation":"DispatchEx then dynamic.Dispatch","workbooks_add":{"outer_hresult":"0x80020009","excepinfo_scode":"0x800A03EC"},"note":"Current opt-in control did not supersede the preserved Prompt 05B DispatchEx success."}),
            json!({"id":"client.pywin32.makepy-05c","classification":"Inconclusive","mode":"makepy-generated","wrapper_classes":{"Application":{"class":"_Application","module":"win32com.gen_py.Excel-typelib"},"Workbooks":{"class":"Workbooks","module":"win32com.gen_py.Excel-typelib"},"Workbook":"not returned"},"activation":"gencache.EnsureDispatch","workbooks_add":{"outer_hresult":"0x80020009","excepinfo_scode":"0x800A03EC"},"note":"Generated-wrapper control used pywin32 311; selected Automation paths are reconciled in Prompt 05D, but the historic 05C runtime observation remains separately classified."}),
            json!({"id":"client.pywin32.05d-control-311-dynamic","classification":"Control-confirmed","environment":"A","mode":"dynamic","activation":"DispatchEx then dynamic.Dispatch","wrapper_classes":{"Application":{"class":"CDispatch","module":"win32com.client.dynamic"},"Workbooks":{"class":"CDispatch","module":"win32com.client.dynamic"},"Workbook":{"class":"CDispatch","module":"win32com.client.dynamic"}},"workbooks_add":{"succeeded":true,"created_workbook":"Book1"},"session_state_recorded":true,"owned_process_exit":true,"raw_identity_values_recorded":false}),
            json!({"id":"client.pywin32.05d-control-311-generated","classification":"Control-confirmed","environment":"A","mode":"makepy-generated","activation":"gencache.EnsureDispatch","wrapper_classes":{"Application":{"class":"_Application","module":"win32com.gen_py.Excel-typelib"},"Workbooks":{"class":"Workbooks","module":"win32com.gen_py.Excel-typelib"},"Workbook":{"class":"Workbook","module":"win32com.gen_py.Excel-typelib"}},"workbooks_add":{"succeeded":true,"created_workbook":"Book1"},"session_state_recorded":true,"owned_process_exit":true,"raw_identity_values_recorded":false}),
            json!({"id":"client.pywin32.05d-control-312-dynamic","classification":"Control-confirmed","environment":"B","mode":"dynamic","activation":"DispatchEx then dynamic.Dispatch","wrapper_classes":{"Application":{"class":"CDispatch","module":"win32com.client.dynamic"},"Workbooks":{"class":"CDispatch","module":"win32com.client.dynamic"},"Workbook":{"class":"CDispatch","module":"win32com.client.dynamic"}},"workbooks_add":{"succeeded":true,"created_workbook":"Book1"},"session_state_recorded":true,"owned_process_exit":true,"raw_identity_values_recorded":false}),
            json!({"id":"client.pywin32.05d-control-312-generated","classification":"Control-confirmed","environment":"B","mode":"makepy-generated","activation":"gencache.EnsureDispatch","wrapper_classes":{"Application":{"class":"_Application","module":"win32com.gen_py.Excel-typelib"},"Workbooks":{"class":"Workbooks","module":"win32com.gen_py.Excel-typelib"},"Workbook":{"class":"Workbook","module":"win32com.gen_py.Excel-typelib"}},"workbooks_add":{"succeeded":true,"created_workbook":"Book1"},"session_state_recorded":true,"owned_process_exit":true,"raw_identity_values_recorded":false}),
            json!({"id":"client.pywin32.prompt-05b-preserved-control","classification":"Control-confirmed","activation":"win32com.client.DispatchEx(Excel.Application)","excel_version":"16.0","workbooks_add":"succeeded","created_workbook":"Book1","evidence_boundary":"historical runtime control preserved without reclassification"}),
        ],
        "comtypes" => vec![
            json!({"id":"client.comtypes.05d-env-c","classification":"Control-confirmed","environment":"C","python_version":"3.11.7","interpreter_architecture":"x64","package_architecture":"x64 pure Python","comtypes_version":"1.4.16","excel_version":"16.0","office_bitness":"64-bit","wrapper_modes":["dynamic","generated"],"generated_cache":"isolated per local control mode; path not committed","raw_paths_recorded":false}),
            json!({"id":"client.comtypes.local-python-311","classification":"Control-confirmed","python_version":"3.11.7","comtypes_version":"1.4.16","module_identities":{"comtypes":"comtypes","comtypes_client":"comtypes.client"},"com_initialization":{"source":"comtypes imports call CoInitializeEx with sys.coinit_flags or COINIT_APARTMENTTHREADED and register a shutdown handler"},"installed_version_matches_source_manifest":true,"installed_exact_commit_match":"unverified: wheel metadata carries no source commit","raw_paths_recorded":false}),
            json!({"id":"client.comtypes.dynamic-05c-initial-control","classification":"Control-confirmed","mode":"dynamic","activation":"comtypes.client.CreateObject(Excel.Application, dynamic=True)","wrapper_classes":{"Application":{"class":"Dispatch","module":"comtypes.client.lazybind"},"Workbooks":{"class":"Dispatch","module":"comtypes.client.lazybind"},"Workbook":{"class":"Dispatch","module":"comtypes.client.lazybind"}},"member_resolution":"ITypeComp.Bind followed by the private IDispatch._invoke path","workbooks_add":{"succeeded":true,"zero_argument_dispparams":"cArgs=0, cNamedArgs=0, rgvarg=null, rgdispidNamedArgs=null"},"excel_version":"16.0","created_workbook":"Book1","boundary":"earlier bounded control; preserved separately from the later recheck"}),
            json!({"id":"client.comtypes.dynamic-05c-recheck","classification":"Inconclusive","mode":"dynamic","wrapper_classes":{"Application":{"class":"Dispatch","module":"comtypes.client.lazybind"},"Workbooks":{"class":"Dispatch","module":"comtypes.client.lazybind"},"Workbook":"not returned"},"workbooks_add":{"outer_hresult":"0x80020009","excepinfo_scode":"0x800A03EC"},"note":"Later bounded recheck returned Excel's resource exception; it does not supersede the initial control-confirmed observation."}),
            json!({"id":"client.comtypes.generated-05c-initial-control","classification":"Control-confirmed","mode":"generated","activation":"GetModule(Excel typelib) then comtypes.client.CreateObject(Excel.Application)","wrapper_classes":{"Application":{"class":"POINTER(_Application)","module":"comtypes._post_coinit.unknwn"},"Workbooks":{"class":"POINTER(Workbooks)","module":"comtypes._post_coinit.unknwn"},"Workbook":{"class":"POINTER(_Workbook)","module":"comtypes._post_coinit.unknwn"}},"member_resolution":"generated dual-interface vtable bindings for the observed Application, Workbooks, and Workbook chain","workbooks_add":{"succeeded":true,"dispparams":"not used by the vtable call"},"excel_version":"16.0","created_workbook":"Book1","boundary":"earlier bounded control; other corpus members and selected dispinterface paths were not runtime exercised"}),
            json!({"id":"client.comtypes.generated-05c-recheck","classification":"Inconclusive","mode":"generated","wrapper_classes":{"Application":{"class":"POINTER(_Application)","module":"comtypes._post_coinit.unknwn"},"Workbooks":{"class":"POINTER(Workbooks)","module":"comtypes._post_coinit.unknwn"},"Workbook":"not returned"},"workbooks_add":{"outer_hresult":"0x80020009","excepinfo_scode":"0x800A03EC"},"note":"Later bounded recheck returned Excel's resource exception; it does not supersede the initial control-confirmed observation."}),
            json!({"id":"client.comtypes.05d-control-dynamic","classification":"Control-confirmed","environment":"C","mode":"dynamic","activation":"CreateObject(dynamic=True)","wrapper_classes":{"Application":{"class":"Dispatch","module":"comtypes.client.lazybind"},"Workbooks":{"class":"Dispatch","module":"comtypes.client.lazybind"},"Workbook":{"class":"Dispatch","module":"comtypes.client.lazybind"}},"workbooks_add":{"succeeded":true,"created_workbook":"Book1"},"session_state_recorded":true,"owned_process_exit":true,"raw_identity_values_recorded":false}),
            json!({"id":"client.comtypes.05d-control-generated","classification":"Control-confirmed","environment":"C","mode":"generated","activation":"GetModule then CreateObject","wrapper_classes":{"Application":{"class":"POINTER(_Application)","module":"comtypes._post_coinit.unknwn"},"Workbooks":{"class":"POINTER(Workbooks)","module":"comtypes._post_coinit.unknwn"},"Workbook":{"class":"POINTER(_Workbook)","module":"comtypes._post_coinit.unknwn"}},"workbooks_add":{"succeeded":true,"created_workbook":"Book1"},"session_state_recorded":true,"owned_process_exit":true,"raw_identity_values_recorded":false}),
        ],
        _ => return Err(format!("unknown client {client}")),
    };
    json_lines(values)
}

fn pywin32_version_reconciliation() -> Result<String, String> {
    json_lines(vec![
        json!({"id":"pywin32.reconciliation.dynamic-cdispatch","path":"com/win32com/client/dynamic.py","b311_to_312_1":"identical","b312_to_312_1":"identical","scope":"dynamic.CDispatch member lookup and wrapping","conclusion":"installed 311 source supports the inspected dynamic dispatch mechanics"}),
        json!({"id":"pywin32.reconciliation.pydispatch-invoke","path":"com/win32com/src/PyIDispatch.cpp","b311_to_312_1":"semantically equivalent","b312_to_312_1":"identical","scope":"Invoke, InvokeTypes, reverse rgvarg order, puArgErr, EXCEPINFO, and result conversion","conclusion":"C API modernization does not change the selected Excel Automation behaviour"}),
        json!({"id":"pywin32.reconciliation.dispatch-selection","path":"com/win32com/client/__init__.py","b311_to_312_1":"changed but irrelevant","b312_to_312_1":"identical","scope":"generated-wrapper QueryInterface fast path","conclusion":"does not alter the examined activation or Invoke spine"}),
        json!({"id":"pywin32.reconciliation.generated-descriptors","path":"com/win32com/client/build.py","b311_to_312_1":"changed but irrelevant","b312_to_312_1":"identical","scope":"VT_RECORD generated out-parameter metadata","conclusion":"outside the bounded Excel Range Automation corpus"}),
        json!({"id":"pywin32.reconciliation.generated-cache","path":"com/win32com/client/gencache.py","b311_to_312_1":"semantically equivalent","b312_to_312_1":"identical","scope":"generated-wrapper cache creation and cleanup","conclusion":"cache robustness changes do not alter generated member descriptors"}),
        json!({"id":"pywin32.reconciliation.generator-output","path":"com/win32com/client/genpy.py","b311_to_312_1":"changed but irrelevant","b312_to_312_1":"identical","scope":"temporary filename and atomic generated-wrapper writes","conclusion":"no selected generated invocation semantics changed"}),
        json!({"id":"pywin32.reconciliation.error-utils","path":"com/win32com/src/ErrorUtils.cpp","b311_to_312_1":"semantically equivalent","b312_to_312_1":"identical","scope":"EXCEPINFO materialization, deferred fill-in, and BSTR cleanup","conclusion":"C API modernization preserves the examined error mechanics"}),
        json!({"id":"pywin32.reconciliation.oleargs","path":"com/win32com/src/oleargs.cpp","b311_to_312_1":"materially changed outside bounded corpus","b312_to_312_1":"identical","scope":"PythonOleArgHelper VARIANT and SAFEARRAY conversion","conclusion":"VT_RECORD SAFEARRAY and typed VT_INT changes are material, but ordinary VARIANT, missing, null, and property-put forms used here remain semantically equivalent"}),
        json!({"id":"pywin32.reconciliation.pythoncom","path":"com/win32com/src/PythonCOM.cpp","b311_to_312_1":"semantically equivalent for local Excel activation","b312_to_312_1":"changed but irrelevant","scope":"CoCreateInstanceEx and COM helper loading","conclusion":"modern local Excel activation remains equivalent; the b312 to 312.1 ObjectFromAddress change is outside this automation path"}),
    ])
}

fn activation(client: &str) -> Result<String, String> {
    let values = match client {
        "pywin32" => vec![
            json!({"id":"pywin32.activation.DispatchEx","classification":"Source-established","entry":"win32com.client.DispatchEx","progid_or_clsid":"Excel.Application","activation":"pythoncom.CoCreateInstanceEx","clsctx":"CLSCTX_SERVER","requested_iid":"IID_IDispatch","server_info":"machine is translated to server-info when supplied","outer_unknown":"returned IDispatch is wrapped by Dispatch","type_info_probing":"Dispatch delegates to __WrapDispatch and may select a generated class"}),
            json!({"id":"pywin32.activation.Dispatch","classification":"Source-established","entry":"win32com.client.Dispatch","activation":"wrap existing IDispatch or resolve and create through __GetGoodDispatch","wrapper_selection":"__WrapDispatch probes type information and gencache before falling back to dynamic.CDispatch"}),
        ],
        "comtypes" => vec![
            json!({"id":"comtypes.activation.CreateObject.dynamic","classification":"Source-established","entry":"comtypes.client.CreateObject(dynamic=True)","activation":"comtypes.CoCreateInstance or CoCreateInstanceEx","clsctx":"default from CLSCTX_SERVER when omitted","requested_interface":"IDispatch","wrapper_selection":"client.dynamic.Dispatch selects lazybind.Dispatch when type information is present; otherwise client.dynamic._Dispatch"}),
            json!({"id":"comtypes.activation.CreateObject.generated","classification":"Source-established","entry":"comtypes.client.CreateObject(dynamic=False)","activation":"comtypes.CoCreateInstance or CoCreateInstanceEx","requested_interface":"coclass default interface when available","wrapper_selection":"GetBestInterface loads type information, generates wrapper module, then queries the selected interface"}),
            json!({"id":"comtypes.activation.GetActiveObject","classification":"Source-established","entry":"comtypes.client.GetActiveObject","activation":"running-object lookup then dynamic or GetBestInterface wrapping","server_creation":false}),
        ],
        _ => return Err(format!("unknown client {client}")),
    };
    json_lines(values)
}

fn wrappers(client: &str) -> Result<String, String> {
    let values = match client {
        "pywin32" => vec![
            json!({"id":"pywin32.wrapper.dynamic","classification":"Source-established","class":"dynamic.CDispatch","member_maps":"build.BuildDispatch / CDispatch _olerepr_ maps with GetIDsOfNames fallback and caching","call_paths":"Invoke fallback; Build.MakeFuncMethod may create an in-memory typed InvokeTypes method when function descriptors are available","object_results":"_get_good_object_ wraps IDispatch results via Dispatch"}),
            json!({"id":"pywin32.wrapper.makepy","classification":"Generated-code-established","class":"DispatchBaseClass subclasses under win32com.gen_py","member_maps":"generated _prop_map_get_ and _prop_map_put_ plus generated methods","call_paths":"fixed signature methods generally InvokeTypes; generated Item special method and property put tuple use Invoke","selection":"gencache.EnsureDispatch upgrades a dynamic wrapper when type information is available"}),
        ],
        "comtypes" => vec![
            json!({"id":"comtypes.wrapper.dynamic","classification":"Source-established","class":"client.lazybind.Dispatch when type information is present; client.dynamic._Dispatch fallback otherwise","member_maps":"lazybind caches ITypeComp.Bind results; dynamic fallback caches name-to-DISPID values from IDispatch.GetIDsOfNames","call_paths":"lazybind calls private IDispatch._invoke; fallback calls public IDispatch.Invoke and converts a bad-context property get into MethodCaller","enumeration":"DISPID_NEWENUM then IEnumVARIANT.Next","control_observed":"CreateObject(dynamic=True) selected comtypes.client.lazybind.Dispatch"}),
            json!({"id":"comtypes.wrapper.generated","classification":"Generated-code-established","generator":"client.GetModule / tools.tlbparser","selection":"a TKIND_DISPATCH with a custom dual reference is parsed as a dual interface; otherwise DispMemberGenerator creates IDispatch.Invoke methods","Excel_typelib_correlation":"Application, Workbooks, Workbook, and Worksheet are dual candidates; Worksheets and Range are dispatch-only in the audited interface metadata"}),
        ],
        _ => return Err(format!("unknown client {client}")),
    };
    json_lines(values)
}

fn invocations(client: &str, typelib: &TypelibData) -> Result<String, String> {
    let modes: &[ClientMode] = match client {
        "pywin32" => &[ClientMode::Pywin32Dynamic, ClientMode::Pywin32Generated],
        "comtypes" => &[ClientMode::ComtypesDynamic, ClientMode::ComtypesGenerated],
        _ => return Err(format!("unknown client {client}")),
    };
    let mut values = Vec::new();
    for mode in modes {
        for member in CORPUS {
            values.push(invocation_record(*mode, *member, typelib)?);
        }
    }
    json_lines(values)
}

fn invocation_record(
    mode: ClientMode,
    member: CorpusMember,
    typelib: &TypelibData,
) -> Result<Value, String> {
    let id = member.typelib_id();
    let metadata = typelib
        .members
        .get(&id)
        .ok_or_else(|| format!("missing typelib record {id}"))?;
    let parameters = typelib
        .parameters
        .get(&id)
        .map(Vec::as_slice)
        .unwrap_or(&[]);
    let parameter_records: Vec<Value> = parameters
        .iter()
        .map(|parameter| {
            json!({
                "ordinal": parameter.ordinal,
                "name": parameter.name,
                "required": parameter.required,
                "optional": parameter.optional,
                "default": parameter.default,
                "type": parameter.parameter_type,
            })
        })
        .collect();
    let (
        invoke_api,
        lcid,
        resolution,
        wrapper_class,
        wrapper_module,
        optional,
        named,
        result,
        cleanup,
        error,
    ) = invocation_strategy(mode, member, metadata);
    let flags = dispatch_flags(member.invoke_kind);
    Ok(json!({
        "id": format!("{}.{}.{}", mode.stable_name(), member.canonical(), member.invoke_kind),
        "member": member.canonical(),
        "client": mode.client(),
        "mode": mode.mode(member),
        "expression": member.expression,
        "wrapper_class": wrapper_class,
        "wrapper_module": wrapper_module,
        "activation_context": "Excel.Application local-server activation; client implementation source only unless the environment record says Control-confirmed",
        "member_resolution": resolution,
        "dispid": {"source": metadata.source, "value": metadata.dispid, "typelib_member_id": id},
        "lcid": lcid,
        "invoke": {
            "api": invoke_api,
            "flags": flags,
            "argument_order": if invoke_api.contains("IDispatch::Invoke") || invoke_api == "PyIDispatch.InvokeTypes" || invoke_api == "PyIDispatch.Invoke" { "logical positional arguments are reversed into rgvarg" } else { "vtable parameter order; DISPPARAMS not used" },
            "argument_descriptors": parameter_records,
            "optional_argument_handling": optional,
            "named_dispids": named,
            "result_requested": result,
        },
        "typelib_correlation": {
            "invoke_kind": metadata.invoke_kind,
            "parameter_count": metadata.parameter_count,
            "optional_parameter_count": metadata.optional_parameter_count,
            "return_type": metadata.return_type,
        },
        "returned_wrapper": returned_wrapper(member, mode),
        "cleanup_path": cleanup,
        "error_translation": error,
        "upstream_source_references": source_references(mode),
        "classification": classification_for(mode, member),
    }))
}

#[allow(clippy::type_complexity)]
fn invocation_strategy(
    mode: ClientMode,
    member: CorpusMember,
    metadata: &MemberMeta,
) -> (
    &'static str,
    Value,
    Value,
    &'static str,
    &'static str,
    Value,
    Value,
    Value,
    &'static str,
    &'static str,
) {
    let property_put = matches!(
        member.invoke_kind,
        "INVOKE_PROPERTYPUT" | "INVOKE_PROPERTYPUTREF"
    );
    let py_optional = if metadata.optional_parameter_count == 0 {
        json!("no optional parameters")
    } else {
        json!(
            "dynamic mode sends only supplied trailing arguments; pythoncom.Missing converts to VT_ERROR/DISP_E_PARAMNOTFOUND when explicitly supplied"
        )
    };
    match mode {
        ClientMode::Pywin32Dynamic => (
            if member.name == "_NewEnum" {
                "PyIDispatch.InvokeTypes"
            } else {
                "PyIDispatch.Invoke or in-memory typed InvokeTypes"
            },
            json!({"value":0,"kind":"explicitly passed"}),
            json!(
                "CDispatch uses build/type-information maps when available; otherwise it caches GetIDsOfNames(0, name)"
            ),
            "dynamic.CDispatch",
            "win32com.client.dynamic",
            py_optional,
            if property_put {
                json!(["DISPID_PROPERTYPUT"])
            } else {
                json!([])
            },
            json!("requested for reads/methods; property puts request no result"),
            "PyCom_MakeUntypedDISPPARAMS initializes null pointers for zero arguments; helpers and VARIANTARGs are cleared after Invoke",
            "pythoncom.com_error; PyIDispatch initializes puArgErr to UINT_MAX and translates EXCEPINFO with deferred fill-in and BSTR cleanup",
        ),
        ClientMode::Pywin32Generated => (
            if property_put || member.name == "Item" {
                "PyIDispatch.Invoke"
            } else {
                "PyIDispatch.InvokeTypes"
            },
            json!({"value":0,"kind":"generated LCID constant"}),
            json!(
                "generated makepy member map supplies the DISPID; Item has a generated special Invoke projection"
            ),
            "DispatchBaseClass subclass",
            "win32com.gen_py.Excel-typelib",
            if metadata.optional_parameter_count == 0 {
                json!("no optional parameters")
            } else {
                json!(
                    "generated signature uses pythoncom.Missing defaults; InvokeTypes omits trailing non-byref Missing arguments"
                )
            },
            if property_put {
                json!(["DISPID_PROPERTYPUT"])
            } else {
                json!([])
            },
            json!(
                "derived from generated return descriptor; fixed simple returns are converted by InvokeTypes, object returns are wrapped with Dispatch"
            ),
            "InvokeTypes clears typed argument/result helpers; generated property put delegates to Invoke cleanup",
            "pythoncom.com_error with the same PyIDispatch EXCEPINFO and puArgErr normalization",
        ),
        ClientMode::ComtypesDynamic => (
            "lazybind.Dispatch → private IDispatch::_invoke → IDispatch::Invoke",
            json!({"value":0,"kind":"default keyword _lcid"}),
            json!(
                "the observed lazybind path caches ITypeComp.Bind(name, invkind); client.dynamic._Dispatch is the no-type-information fallback and caches GetIDsOfNames(name)"
            ),
            "client.lazybind.Dispatch when type information is present; client.dynamic._Dispatch fallback otherwise",
            "comtypes.client.lazybind",
            if metadata.optional_parameter_count == 0 {
                json!("no optional parameters")
            } else {
                json!(
                    "only positional arguments are supported by IDispatch.Invoke; named keyword parameters raise ValueError"
                )
            },
            if property_put {
                json!(["DISPID_PROPERTYPUT"])
            } else {
                json!([])
            },
            json!(
                "lazybind _invoke constructs a result VARIANT; returned interface values are wrapped dynamically"
            ),
            "automation._invoke owns the temporary DISPPARAMS/VARIANT storage; for zero arguments it leaves cArgs/cNamedArgs at zero and both pointers null",
            "COMError; public IDispatch.Invoke exposes EXCEPINFO for DISP_E_EXCEPTION and physical puArgErr handling for applicable parameter HRESULTs",
        ),
        ClientMode::ComtypesGenerated if uses_comtypes_vtable(member) => (
            "generated dual-interface vtable call",
            json!("not applicable to vtable dispatch"),
            json!(
                "comtypes tlbparser follows GetRefTypeOfImplType(-1) for a dual dispinterface and generates COMMETHOD bindings"
            ),
            "generated interface pointer",
            "comtypes.gen.Excel-typelib",
            json!(
                "typed optional defaults come from the generated signature; optional VARIANT defaults without a typelib default become VARIANT.missing"
            ),
            json!("not applicable to vtable dispatch"),
            json!("ctypes out-parameter conversion is driven by the generated COMMETHOD signature"),
            "ctypes pointer and out-parameter ownership paths; no DISPPARAMS result VARIANT is constructed by the call site",
            "ctypes COMError from HRESULT; no IDispatch EXCEPINFO frame is used on this selected path",
        ),
        ClientMode::ComtypesGenerated => (
            "IDispatch::Invoke",
            json!({"value":0,"kind":"default keyword _lcid"}),
            json!(
                "DispMemberGenerator embeds the typelib memid in a generated method/property and calls IDispatch.Invoke"
            ),
            "generated dispinterface pointer",
            "comtypes.gen.Excel-typelib",
            json!(
                "generated optional VARIANT parameters without a default use VARIANT.missing; named keyword arguments remain unsupported in IDispatch.Invoke"
            ),
            if property_put {
                json!(["DISPID_PROPERTYPUT"])
            } else {
                json!([])
            },
            json!("IDispatch.Invoke creates a VARIANT result and unwraps dynamically"),
            "DISPPARAMS and VARIANT ctypes destructors own cleanup; SAFEARRAY output ownership transfers through __ctypes_from_outparam__",
            "COMError with EXCEPINFO for DISP_E_EXCEPTION and physical puArgErr handling only for applicable parameter HRESULTs",
        ),
    }
}

fn dispatch_flags(invoke_kind: &str) -> Value {
    match invoke_kind {
        "INVOKE_FUNC" => json!(["DISPATCH_METHOD", "0x0001"]),
        "INVOKE_PROPERTYGET" => json!(["DISPATCH_PROPERTYGET", "0x0002"]),
        "INVOKE_PROPERTYPUT" => json!(["DISPATCH_PROPERTYPUT", "0x0004"]),
        "INVOKE_PROPERTYPUTREF" => json!(["DISPATCH_PROPERTYPUTREF", "0x0008"]),
        _ => json!(["unknown"]),
    }
}

fn uses_comtypes_vtable(member: CorpusMember) -> bool {
    matches!(
        member.owner,
        "Excel.Application" | "Excel.Workbooks" | "Excel.Workbook" | "Excel.Worksheet"
    )
}

fn returned_wrapper(member: CorpusMember, mode: ClientMode) -> &'static str {
    if member.invoke_kind == "INVOKE_PROPERTYPUT" {
        return "no result";
    }
    if matches!(
        member.name,
        "Workbooks"
            | "Worksheets"
            | "Range"
            | "Cells"
            | "UsedRange"
            | "Item"
            | "Add"
            | "Offset"
            | "Resize"
    ) {
        match mode {
            ClientMode::Pywin32Dynamic | ClientMode::Pywin32Generated => {
                "pywin32 dispatch wrapper when the result is IDispatch; scalar Python value otherwise"
            }
            ClientMode::ComtypesDynamic => {
                "dynamic Dispatch wrapper for VT_DISPATCH; ctypes interface wrapper for VT_UNKNOWN"
            }
            ClientMode::ComtypesGenerated => {
                "generated interface pointer when typed; otherwise comtypes VARIANT conversion"
            }
        }
    } else {
        "client scalar conversion or None for VT_EMPTY/VT_NULL"
    }
}

fn source_references(mode: ClientMode) -> Vec<&'static str> {
    match mode {
        ClientMode::Pywin32Dynamic => vec![
            "client/dynamic.py",
            "client/build.py",
            "src/PyIDispatch.cpp",
            "src/oleargs.cpp",
        ],
        ClientMode::Pywin32Generated => vec![
            "client/__init__.py",
            "client/genpy.py",
            "client/gencache.py",
            "src/PyIDispatch.cpp",
        ],
        ClientMode::ComtypesDynamic => {
            vec!["client/lazybind.py", "client/dynamic.py", "automation.py"]
        }
        ClientMode::ComtypesGenerated => vec![
            "client/_generate.py",
            "tools/tlbparser.py",
            "_memberspec.py",
            "automation.py",
        ],
    }
}

fn classification_for(mode: ClientMode, member: CorpusMember) -> &'static str {
    match mode {
        ClientMode::ComtypesGenerated if uses_comtypes_vtable(member) => "Typelib-correlated",
        ClientMode::ComtypesGenerated => "Generated-code-established",
        _ => "Source-established",
    }
}

fn conversions(client: &str) -> Result<String, String> {
    let values = match client {
        "pywin32" => vec![
            json!({"id":"pywin32.conversion.scalars","classification":"Source-established","input":{"None":"VT_NULL","bool":"VT_BOOL","int":"VT_I4/UI4/I8 as needed","float":"VT_R8","str":"VT_BSTR","datetime":"VT_DATE","decimal":"VT_CY","dispatch":"VT_DISPATCH","unknown":"VT_UNKNOWN","missing":"VT_ERROR/DISP_E_PARAMNOTFOUND","empty":"VT_EMPTY"},"output":{"VT_EMPTY":"None","VT_NULL":"None","VT_BOOL":"bool","VT_I2/VT_I4":"int","VT_R4/VT_R8":"float","VT_DATE":"datetime","VT_CY":"Decimal","VT_BSTR":"str","VT_ERROR":"int scode","VT_DISPATCH":"PyIDispatch","VT_UNKNOWN":"PyIUnknown"},"cleanup":"VariantClear after PyCom_PyObjectFromVariant"}),
            json!({"id":"pywin32.conversion.safearray","classification":"Source-established","input":"Python sequences become SAFEARRAY(VARIANT); buffers use VT_ARRAY|VT_UI1","output":"PyCom_PyObjectFromSAFEARRAY converts dimensions and elements to nested Python sequences","rank_bounds":"source walks SAFEARRAY dimensions; this does not establish Excel Range bounds","ownership":"VARIANT cleanup owns returned SAFEARRAY after conversion"}),
        ],
        "comtypes" => vec![
            json!({"id":"comtypes.conversion.scalars","classification":"Source-established","input":{"None":"VT_NULL","empty_sequence":"VT_EMPTY special case","bool":"VT_BOOL","int":"VT_I4","float":"VT_R8","str":"VT_BSTR","datetime":"VT_DATE","decimal":"VT_CY","dispatch":"VT_DISPATCH","unknown":"VT_UNKNOWN","missing":"VT_ERROR/DISP_E_PARAMNOTFOUND"},"output":{"VT_EMPTY":"None","VT_NULL":"None","VT_BOOL":"bool","VT_I2/VT_I4":"int","VT_R4/VT_R8":"float","VT_DATE":"datetime","VT_CY":"Decimal","VT_BSTR":"str","VT_DISPATCH":"wrapped interface","VT_UNKNOWN":"wrapped interface"},"note":"automation.py explicitly marks some getter combinations as unimplemented; no unsupported conversion is claimed"}),
            json!({"id":"comtypes.conversion.safearray","classification":"Source-established","input":"list/tuple creates SAFEARRAY(VARIANT); array.array and compatible ndarray can choose typed primitive SAFEARRAYs","output":"_midlSAFEARRAY(...).unpack() provides the Python result","rank_bounds":"source exposes unpacking but this prompt does not claim Excel result rank/bounds","ownership":"SAFEARRAY out-param marks _needsfree and destroys on object cleanup; VARIANT out-param clears after extracting value"}),
        ],
        _ => return Err(format!("unknown client {client}")),
    };
    json_lines(values)
}

fn errors(client: &str) -> Result<String, String> {
    let values = match client {
        "pywin32" => vec![
            json!({"id":"pywin32.errors.invoke","classification":"Source-established","hresult_detection":"PyIDispatch checks Invoke HRESULT","exception":"DISP_E_EXCEPTION calls PyCom_BuildPyExceptionFromEXCEPINFO","deferred_fill_in":"ErrorUtils invokes pfnDeferredFillIn before materializing EXCEPINFO","arg_error":"initialized to UINT_MAX; only parameter-not-found/type-mismatch paths are translated from reverse rgvarg order","bstr_cleanup":"PyCom_CleanupExcepInfo frees EXCEPINFO BSTRs after Python exception construction","result_cleanup":"VariantClear runs after result conversion"}),
        ],
        "comtypes" => vec![
            json!({"id":"comtypes.errors.invoke","classification":"Source-established","hresult_detection":"ctypes raises COMError from IDispatch.Invoke","exception":"DISP_E_EXCEPTION replaces details with description, source, help file, help context, and scode from EXCEPINFO","deferred_fill_in":"not invoked by the inspected IDispatch.Invoke implementation","arg_error":"DISP_E_PARAMNOTFOUND returns argerr; DISP_E_TYPEMISMATCH reports argerr plus one","bstr_cleanup":"ctypes BSTR/EXCEPINFO ownership is managed by the declared structures","result_cleanup":"VARIANT and DISPPARAMS destructors own post-call cleanup"}),
        ],
        _ => return Err(format!("unknown client {client}")),
    };
    json_lines(values)
}

fn patterns(client: &str) -> Result<String, String> {
    let values = match client {
        "pywin32" => vec![
            json!({"id":"pywin32.pattern.activation","classification":"Source-established","rule":"DispatchEx requests IID_IDispatch through CoCreateInstanceEx and then selects a dynamic or generated wrapper"}),
            json!({"id":"pywin32.pattern.frames","classification":"Source-established","rule":"Invoke reverses positional arguments; zero arguments leave both DISPPARAMS pointers null; puts attach DISPID_PROPERTYPUT"}),
            json!({"id":"pywin32.pattern.generated","classification":"Generated-code-established","rule":"makepy emits type descriptors and usually InvokeTypes; omitted generated Missing arguments shrink trailing non-byref arguments"}),
            json!({"id":"pywin32.pattern.collections","classification":"Generated-code-established","rule":"generated and dynamic wrappers use DISPID_NEWENUM/InvokeTypes for iteration and special Item projection for collection indexing"}),
        ],
        "comtypes" => vec![
            json!({"id":"comtypes.pattern.activation","classification":"Source-established","rule":"CreateObject(dynamic=True) requests IDispatch; generated mode uses GetBestInterface/type-library generation"}),
            json!({"id":"comtypes.pattern.frames","classification":"Source-established","rule":"IDispatch.Invoke reverses VARIANT arguments and attaches DISPID_PROPERTYPUT for property assignment; arbitrary named keywords are unsupported"}),
            json!({"id":"comtypes.pattern.generated","classification":"Generated-code-established","rule":"TKIND_DISPATCH with a dual reference is generated as a typed vtable interface; ordinary dispinterfaces retain IDispatch.Invoke"}),
            json!({"id":"comtypes.pattern.collections","classification":"Source-established","rule":"dynamic iteration invokes DISPID_NEWENUM and consumes IEnumVARIANT.Next; indexing is an enumeration operation, not proof of Excel Item indexing"}),
        ],
        _ => return Err(format!("unknown client {client}")),
    };
    json_lines(values)
}

fn json_lines(values: Vec<Value>) -> Result<String, String> {
    let mut result = String::new();
    for value in values {
        result.push_str(&serde_json::to_string(&value).map_err(|error| error.to_string())?);
        result.push('\n');
    }
    Ok(result)
}

fn insert_reports(
    files: &mut BTreeMap<PathBuf, String>,
    typelib: &TypelibData,
) -> Result<(), String> {
    let base = PathBuf::from("generated/client-implementations");
    files.insert(base.join("activation-comparison.md"), activation_report());
    files.insert(base.join("dynamic-vs-generated.md"), dynamic_report());
    files.insert(
        base.join("representative-member-matrix.md"),
        member_matrix(typelib)?,
    );
    files.insert(
        base.join("invocation-patterns.md"),
        invocation_patterns_report(),
    );
    files.insert(base.join("property-put-patterns.md"), property_put_report());
    files.insert(
        base.join("optional-argument-patterns.md"),
        optional_report(),
    );
    files.insert(base.join("collection-patterns.md"), collection_report());
    files.insert(
        base.join("value-conversion-patterns.md"),
        conversion_report(),
    );
    files.insert(
        base.join("safearray-conversion-patterns.md"),
        safearray_report(),
    );
    files.insert(base.join("error-handling-patterns.md"), error_report());
    files.insert(base.join("pywin32-vs-comtypes.md"), comparison_report());
    files.insert(base.join("pywin32-311-vs-312.md"), pywin32_version_report());
    files.insert(
        base.join("source-version-reconciliation.md"),
        source_version_reconciliation_report(),
    );
    files.insert(
        base.join("implemented-rust-parity.md"),
        implemented_rust_parity_report(),
    );
    files.insert(base.join("rust-parity-backlog.md"), parity_report());
    files.insert(base.join("unresolved.md"), unresolved_report());
    Ok(())
}

fn report_header(title: &str) -> String {
    format!(
        "# {title}\n\nGenerated by `excel-com-client-kb`; do not edit by hand. This is source-derived client-implementation evidence, not Excel runtime evidence unless explicitly labelled Control-confirmed.\n\n"
    )
}

fn activation_report() -> String {
    let mut output = report_header("Activation comparison");
    output.push_str("| Client path | Activation | Requested interface | Wrapper selection | Classification |\n| --- | --- | --- | --- | --- |\n");
    output.push_str("| pywin32 `DispatchEx` | `CoCreateInstanceEx` with `CLSCTX_SERVER` | `IID_IDispatch` | `Dispatch` probes type information, then generated cache or `CDispatch` | Source-established |\n");
    output.push_str("| pywin32 `Dispatch` | wraps existing dispatch or resolves/creates through internal helper | `IDispatch` | generated wrapper if available; otherwise `dynamic.CDispatch` | Source-established |\n");
    output.push_str("| comtypes dynamic | `CreateObject(dynamic=True)` calls `CoCreateInstance` or `CoCreateInstanceEx` | `IDispatch` | `lazybind.Dispatch` with type information; `_Dispatch` fallback otherwise | Source-established; lazybind control-confirmed |\n");
    output.push_str("| comtypes generated | `CreateObject` and `GetBestInterface` load type information | coclass default interface | generated dispinterface or dual-interface wrapper | Source-established |\n\n");
    output.push_str("The current Rust probe uses `CoCreateInstance(CLSCTX_LOCAL_SERVER)` and LCID `0x0400`; this report does not select a replacement activation or apartment policy.\n");
    output
}

fn dynamic_report() -> String {
    let mut output = report_header("Dynamic versus generated dispatch");
    output.push_str("| Client mode | Member map / DISPID source | Primary call path | LCID | Optional arguments |\n| --- | --- | --- | --- | --- |\n");
    output.push_str("| pywin32 dynamic | type-information map when present; cached `GetIDsOfNames(0, name)` fallback | `Invoke`; dynamic in-memory methods may use `InvokeTypes` when descriptors are available | 0 | visible trailing arguments; explicit `pythoncom.Missing` becomes missing marker |\n");
    output.push_str("| pywin32 makepy | generated maps and method constants | fixed signatures mostly `InvokeTypes`; generated `Item` and property puts use `Invoke` | 0 | generated Missing defaults suppress trailing non-byref arguments |\n");
    output.push_str("| comtypes dynamic | `lazybind` caches `ITypeComp.Bind`; `_Dispatch` fallback caches `GetIDsOfNames(name)` | private `_invoke` then `IDispatch::Invoke`; zero args leave both frame pointers null | 0 | positional only; arbitrary named keywords are rejected |\n");
    output.push_str("| comtypes generated | embedded typelib memids; parser chooses dispatch or dual interface | `IDispatch::Invoke` for dispinterfaces; vtable for dual interfaces | 0 for dispatch calls | generated `VARIANT` optional defaults use `VARIANT.missing` when no default exists |\n\n");
    output.push_str("The opt-in pywin32 311 controls identified `dynamic.CDispatch` in dynamic mode and `win32com.gen_py` classes in generated mode. Both current `Workbooks.Add` controls were Inconclusive with `0x80020009` / `0x800A03EC`; they do not replace the preserved Prompt 05B success. Initial isolated comtypes 1.4.16 controls succeeded: dynamic selected `lazybind.Dispatch` and the generated Application/Workbooks/Workbook chain selected vtable pointers. Later comtypes rechecks reached the same wrappers but were Inconclusive with the same HRESULT pair; both observations remain separate.\n");
    output
}

fn member_matrix(typelib: &TypelibData) -> Result<String, String> {
    let mut output = report_header("Representative member matrix");
    output.push_str("The 42 forms exceed the nominal 20–30 target because the mandatory corpus explicitly includes get/put forms and `_NewEnum` for both collections. Excluded families: charts, pivots, shapes, events, connections, and the remainder of the object model.\n\n");
    output.push_str("| Member form | Family | Mechanism | DISPID | Parameters / optional | pywin32 dynamic | pywin32 generated | comtypes dynamic | comtypes generated |\n| --- | --- | --- | ---: | --- | --- | --- | --- | --- |\n");
    for member in CORPUS {
        let metadata = typelib
            .members
            .get(&member.typelib_id())
            .expect("validated corpus");
        let generated = if uses_comtypes_vtable(*member) {
            "dual vtable"
        } else {
            "dispinterface Invoke"
        };
        output.push_str(&format!(
            "| `{}` `{}` | {} | {} | {} | {} / {} | Invoke or dynamic InvokeTypes | {} | IDispatch.Invoke | {} |\n",
            member.canonical(), member.invoke_kind, member.family, member.mechanism, metadata.dispid,
            metadata.parameter_count, metadata.optional_parameter_count,
            if member.invoke_kind == "INVOKE_PROPERTYPUT" || member.name == "Item" { "Invoke" } else { "InvokeTypes" },
            generated,
        ));
    }
    Ok(output)
}

fn invocation_patterns_report() -> String {
    let mut output = report_header("Invocation patterns");
    output.push_str("| Pattern | pywin32 dynamic | pywin32 generated | comtypes | Shared rule / uncertainty |\n| --- | --- | --- | --- | --- |\n");
    output.push_str("| Property get | `CDispatch` map then `Invoke` or in-memory typed method | generated map, normally `InvokeTypes` | dynamic/dispinterface `Invoke`; selected dual interface can use vtable | invoke kind follows typelib; do not equate source path with Excel outcome |\n");
    output.push_str("| Method | dynamic callable maps member | fixed generated methods use descriptors | `Invoke` or generated `COMMETHOD` | positional Automation arguments are reverse-ordered only on IDispatch paths |\n");
    output.push_str("| Object result | wrap `IDispatch` through `Dispatch` | generated class lookup then wrap | dynamic `Dispatch` or typed interface out-param | retain client-owned wrapper lifetime independently of raw result `VARIANT` |\n");
    output.push_str("| Default / Item | source has special Item projection | generated Item special projection uses `Invoke` | dynamic method/property helper; generated wrapper depends on interface kind | source does not prove Excel indexing base |\n");
    output.push_str("| Enumeration | `DISPID_NEWENUM` with `InvokeTypes` | generated iterator does the same | `DISPID_NEWENUM`, then `IEnumVARIANT.Next` | iteration is separate from `Item` |\n");
    output
}

fn property_put_report() -> String {
    let mut output = report_header("Property-put patterns");
    output.push_str("Applies to `Application.Visible`, `Worksheet.Name`, `Range.Value2`, and `Range.Formula`.\n\n");
    output.push_str("| Client | Flag selection | named DISPID | order | conversion |\n| --- | --- | --- | --- | --- |\n");
    output.push_str("| pywin32 dynamic | `DISPATCH_PROPERTYPUT`; putref when applicable | `DISPID_PROPERTYPUT` | value is last logical argument and first `rgvarg` element after reversal | untyped `PythonOleArgHelper` |\n");
    output.push_str("| pywin32 generated | generated setter calls `Invoke` | `DISPID_PROPERTYPUT` | generated setter appends value to stored Invoke tuple | generated wrapper chooses normal/typed member metadata |\n");
    output.push_str("| comtypes dynamic / dispinterface | `DISPATCH_PROPERTYPUT` or `PROPERTYPUTREF` by object test | `DISPID_PROPERTYPUT` | reversed `VARIANT` array | `VARIANT.value` conversion |\n");
    output.push_str("| comtypes generated dual | generated setter/vtable binding | not an `IDispatch` frame | declared parameter order | ctypes type declaration |\n\n");
    output.push_str("Property put is therefore a distinct parity item: it cannot be inferred from zero-argument `Workbooks.Add`.\n");
    output
}

fn optional_report() -> String {
    let mut output = report_header("Optional-argument patterns");
    output.push_str("| Form | pywin32 | comtypes | Corpus impact |\n| --- | --- | --- | --- |\n");
    output.push_str("| Omitted trailing Python argument | dynamic sends fewer arguments; generated `pythoncom.Missing` stops non-byref `InvokeTypes` argument emission | generated `VARIANT` optional defaults become `VARIANT.missing`; dynamic only accepts positional arguments | `Add`, `Open`, `Close`, `SaveAs`, `Worksheets.Add`, `Find`, `Sort`, `Run` |\n");
    output.push_str("| Explicit missing marker | `pythoncom.Missing` converts to `VT_ERROR` / `DISP_E_PARAMNOTFOUND` | `VARIANT.missing` is `VT_ERROR` / `DISP_E_PARAMNOTFOUND` | distinguish omission from an explicit marker |\n");
    output.push_str(
        "| `None` | `VT_NULL` | `VT_NULL` | never infer omission from Python spelling alone |\n",
    );
    output.push_str("| empty | pywin32 `Empty` special object is `VT_EMPTY` | `VARIANT.empty` is `VT_EMPTY` | distinct from null and missing |\n");
    output.push_str("| named keyword | generated Python signature can bind a Python name but still emits positional COM argument layout | dynamic IDispatch.Invoke rejects arbitrary keywords | `SaveAs`, `Find`, and `Sort` require a separate raw named-DISPID decision in Prompt 05D |\n\n");
    output.push_str("No client-source rule establishes which of these Excel accepts for every member. The existing `Workbooks.Add` runtime variants remain preserved.\n");
    output
}

fn collection_report() -> String {
    let mut output = report_header("Collection patterns");
    output.push_str("| Collection feature | pywin32 | comtypes | Evidence boundary |\n| --- | --- | --- | --- |\n");
    output.push_str("| `Count` | generated property/method map or dynamic map | dispatch or dual interface based on typeinfo | typelib-correlated DISPID/return type |\n");
    output.push_str("| `Item` | generated special `__getitem__` uses `Invoke`; dynamic wrapper builds a callable | dynamic enumeration indexing differs from `Item`; generated strategy follows interface | source does not establish 1-based Excel semantics |\n");
    output.push_str("| `_NewEnum` | uses `DISPID_NEWENUM` and an iterator projection | calls `Invoke(-4)` then `IEnumVARIANT.Next` | source-established mechanics |\n");
    output.push_str("| wrapping | returns Dispatch-wrapped object when applicable | dynamic dispatch or typed generated pointer | ownership is client runtime behaviour, not Excel object identity proof |\n");
    output
}

fn conversion_report() -> String {
    let mut output = report_header("Value-conversion patterns");
    output.push_str("Both clients implement client-specific mappings rather than exposing a raw public `VARIANT` model.\n\n");
    output.push_str("| VARTYPE | pywin32 result | comtypes result |\n| --- | --- | --- |\n");
    output.push_str("| `VT_EMPTY`, `VT_NULL` | `None` | `None` |\n| `VT_BOOL` | `bool` | `bool` |\n| `VT_I2`, `VT_I4`, integral variants | Python `int` | Python `int` |\n| `VT_R4`, `VT_R8` | `float` | `float` |\n| `VT_DATE` | pywin date/time object | `datetime` |\n| `VT_CY` | `Decimal` conversion | `Decimal` |\n| `VT_BSTR` | `str` | `str` |\n| `VT_ERROR` | integer scode | conversion is incomplete for some getter forms; error path is `COMError` |\n| `VT_DISPATCH`, `VT_UNKNOWN` | COM wrapper then Dispatch selection | AddRef then dynamic/generated interface wrapper |\n\n");
    output.push_str("`Range.Value`, `Value2`, `Formula`, and `Formula2` are included as transport members; this table does not claim their Excel runtime result types.\n");
    output
}

fn safearray_report() -> String {
    let mut output = report_header("SAFEARRAY conversion patterns");
    output.push_str(
        "| Topic | pywin32 | comtypes | Rust parity implication |\n| --- | --- | --- | --- |\n",
    );
    output.push_str("| Python input | sequences produce `SAFEARRAY(VARIANT)`; byte buffers produce `VT_UI1` arrays | list/tuple use variant arrays; `array.array` and ndarray may choose typed primitive arrays | model element type and ownership separately from worksheet shape |\n");
    output.push_str("| Result | native converter walks SAFEARRAY dimensions into nested Python sequences | `_midlSAFEARRAY(...).unpack()` returns Python value | preserve rank and bounds before choosing a public shape |\n");
    output.push_str("| Ownership | `VariantClear` releases result array after conversion | out-param array marks itself for destruction; VARIANT extracts then clears | an owned result wrapper is required |\n");
    output.push_str("| Limit | source does not establish Excel's rectangular Range result bounds | source explicitly has unsupported getter combinations | validate against Excel only after workbook creation works |\n");
    output
}

fn error_report() -> String {
    let mut output = report_header("Error-handling patterns");
    output.push_str("| Topic | pywin32 | comtypes | Corrected Rust interpretation |\n| --- | --- | --- | --- |\n");
    output.push_str("| `puArgErr` initialization | `UINT_MAX` sentinel | zero-initialized `c_uint` | initialize Rust with a sentinel; never interpret zero as a source parameter for a zero-argument exception |\n");
    output.push_str("| `DISP_E_EXCEPTION` | materializes EXCEPINFO and invalidates arg index except type-mismatch/param-not-found inner scodes | returns EXCEPINFO details tuple | `0x800A03EC` is application error; `puArgErr` is not applicable |\n");
    output.push_str("| deferred fill-in | source invokes it before building Python exception | inspected Invoke path does not invoke it | retain Rust deferred-fill-in support; record absence/presence |\n");
    output.push_str("| BSTR cleanup | explicit EXCEPINFO cleanup after conversion | ctypes ownership | clear every error/result allocation once |\n\n");
    output.push_str("Historical Prompt 05B records retain their physical `rgvarg_index: 0`; this report adds the source-derived interpretation that it is not a meaningful source parameter for zero-argument `DISP_E_EXCEPTION`.\n");
    output
}

fn comparison_report() -> String {
    let mut output = report_header("pywin32 versus comtypes");
    output.push_str("Shared, source-supported rules:\n\n- Automation positional arguments are reversed on `IDispatch::Invoke` paths.\n- Property puts carry `DISPID_PROPERTYPUT` on `IDispatch` paths.\n- Both use LCID 0 by default for their exposed dispatch helpers.\n- Both translate `None` to `VT_NULL`, distinguish an explicit missing marker, and wrap object returns.\n- Both own temporary `VARIANT`/`DISPPARAMS` storage and translate failing HRESULTs to Python exceptions.\n\nImplementation-specific rules:\n\n- pywin32 makepy emits `InvokeTypes` descriptors and can omit trailing `pythoncom.Missing` arguments; dynamic `CDispatch` may still synthesize typed calls from type information.\n- The observed comtypes dynamic path is `lazybind.Dispatch`: it resolves members with `ITypeComp.Bind` and invokes an empty frame with null pointers. The no-type-information `_Dispatch` fallback uses `GetIDsOfNames`.\n- comtypes generated code may switch from IDispatch to a dual-interface vtable binding, and its dynamic `Invoke` rejects arbitrary named keyword arguments.\n- comtypes exposes explicit `VARIANT.missing`, `VARIANT.empty`, and `VARIANT.null` objects; pywin32 exposes `pythoncom.Missing` and special OLE marker objects.\n\nThe successful comtypes controls rule out neither an Excel/session difference nor a Rust defect; no rule here changes the Rust probe.\n");
    output
}

fn pywin32_version_report() -> String {
    let mut output = report_header("pywin32 311 versus 312.1 selected Automation paths");
    output.push_str(&format!("The installed 311 package was reconciled against its exact upstream `b311` tag `{PYWIN32_311_COMMIT}`, then compared with inspected 312.1 reference `{PYWIN32_COMMIT}`. The released `b312` tag `{PYWIN32_312_COMMIT}` is also compared because a 312.1 wheel was not published for the isolated control environment.\n\n"));
    output.push_str("| Path / behaviour | b311 to 312.1 | b312 to 312.1 | Bounded conclusion |\n| --- | --- | --- | --- |\n");
    output.push_str("| `dynamic.py` / `CDispatch` | identical | identical | Dynamic wrapper mechanics are source-matched. |\n");
    output.push_str("| `PyIDispatch.cpp` / Invoke and InvokeTypes | semantically equivalent | identical | Reverse arguments, `puArgErr`, EXCEPINFO, and result conversion are matched for the bounded corpus. |\n");
    output.push_str("| `__init__.py`, `build.py`, `genpy.py` | changed but irrelevant | identical | Generated wrapper fast paths, record descriptors, and atomic writer details do not affect the selected spine. |\n");
    output.push_str("| `gencache.py`, `ErrorUtils.cpp` | semantically equivalent | identical | Cache robustness and C API modernization retain selected semantics. |\n");
    output.push_str("| `oleargs.cpp` | materially changed outside bounded corpus | identical | VT_RECORD SAFEARRAY and typed VT_INT changed; ordinary VARIANT, missing, null, and property-put forms remain source-equivalent here. |\n");
    output.push_str("| `PythonCOM.cpp` | semantically equivalent for local Excel activation | changed but irrelevant | The b312 to 312.1 `ObjectFromAddress` change is not in the local Excel activation path. |\n\n");
    output.push_str("This is a source reconciliation, not a claim that an intermittent Excel result has one source-level cause.\n");
    output
}

fn source_version_reconciliation_report() -> String {
    let mut output = report_header("Client source-version reconciliation");
    output.push_str("| Client environment | Package / source | Reconciliation result | Boundary |\n| --- | --- | --- | --- |\n");
    output.push_str(&format!("| pywin32 A | 311 / `b311` `{PYWIN32_311_COMMIT}` | Exact upstream 311 source inspected and classified against 312.1. | Selected Automation paths only. |\n"));
    output.push_str(&format!("| pywin32 B | 312 / `b312` `{PYWIN32_312_COMMIT}` | Released 312 selected paths are identical to 312.1 except an irrelevant helper change. | 312.1 source build was attempted only in an isolated environment and was not retained after its bounded build interval. |\n"));
    output.push_str(&format!("| pywin32 reference | 312.1 / `{PYWIN32_COMMIT}` | Retained as the inspected source reference from Prompt 05C. | No unverified wheel-to-commit claim. |\n"));
    output.push_str("| comtypes C | 1.4.16 | Installed package version matches the inspected source version; wheel commit metadata remains unverified. | Existing Prompt 05C source boundary remains unchanged. |\n\n");
    output.push_str("The controls use isolated Python 3.11 x64 environments and isolated generated-wrapper caches. Their local paths, raw HWND values, and raw process identities are deliberately absent from committed evidence.\n");
    output
}

fn implemented_rust_parity_report() -> String {
    let mut output = report_header("Implemented Rust parity configurations");
    output.push_str("Prompt 05D adds research-only configurations to `excel-com-range-probe`; they do not expose a production Excel COM API.\n\n");
    output.push_str("| Mode | Activation / CLSCTX / IID | LCID | Invocation boundary | Type-info or dual-interface boundary |\n| --- | --- | ---: | --- | --- |\n");
    output.push_str("| rust-baseline | `CoCreateInstance` / `CLSCTX_LOCAL_SERVER` / `IID_IDispatch` | `0x0400` | raw `IDispatch::Invoke` | retains historical baseline. |\n");
    output.push_str("| pywin32 dynamic | `CoCreateInstanceEx` / `CLSCTX_SERVER` / `IID_IDispatch` | 0 | raw dispatch comparison with pywin32 source behaviour recorded | no generated descriptor ABI is improvised. |\n");
    output.push_str("| pywin32 generated | `CoCreateInstance` / `CLSCTX_SERVER` / `IID_IDispatch` | 0 | raw fallback while generated InvokeTypes descriptors are recorded | generated descriptors are evidence, not a hand-emitted ABI. |\n");
    output.push_str("| comtypes dynamic | `CoCreateInstance` / `CLSCTX_SERVER` / `IID_IDispatch` | 0 | raw dispatch comparison with terminal lazybind `_invoke` mechanics recorded | raw dynamic path only. |\n");
    output.push_str("| comtypes generated | `CoCreateInstance` / `CLSCTX_SERVER` / `IID_IDispatch` | 0 | raw fallback | generated Rust Excel bindings are unavailable, so no vtable layout is hand-written. |\n\n");
    output.push_str("All modes initialize the apartment as STA, preserve zero-argument null frame pointers, distinguish omitted/missing/empty/null values, initialize `puArgErr` with `UINT_MAX`, normalize EXCEPINFO with BSTR cleanup, clone returned dispatch before `VariantClear`, and retain `DISPID_PROPERTYPUT` frames. `PROPERTYPUTREF` is unit-tested as a distinct frame form.\n");
    output
}

fn parity_report() -> String {
    let mut output = report_header("Rust parity backlog");
    output.push_str("| Priority | Pattern / member | confirmed divergence or question | Recommended Prompt 05D work | Runtime validation needed |\n| --- | --- | --- | --- | --- |\n");
    output.push_str("| P0 | `puArgErr` / `Workbooks.Add` | existing historical zero index is not meaningful for `DISP_E_EXCEPTION` with `0x800A03EC` | preserve raw value, add applicability field, sentinel initialization, and no parameter mapping | rerun diagnostic only; do not infer a workbook result |\n");
    output.push_str("| P0 | activation / COM initialization | pywin32 and comtypes auto-initialize main thread with configurable `sys.coinit_flags`; Rust explicitly selects STA | compare current Rust activation with `DispatchEx` without choosing final apartment policy | controlled activation comparison |\n");
    output.push_str("| P0 | LCID | Python dispatch helpers use 0; Rust uses `0x0400` | make LCID an explicit experiment boundary, not a speculative default change | same member under both LCIDs |\n");
    output.push_str("| P1 | zero-argument frame | pywin32 native Invoke and observed comtypes `lazybind` null-initialize empty pointers; comtypes public Invoke has a distinct zero-length-array helper | retain Rust null frame; inspect only if controlled ABI evidence requires comparison | `Workbooks.Add` after activation issue is isolated |\n");
    output.push_str("| P1 | property puts | clients attach `DISPID_PROPERTYPUT`; generated/vtable paths differ | add tested property-put frame builder | `Visible`, `Worksheet.Name`, `Value2`, `Formula` |\n");
    output.push_str("| P1 | optional arguments | omission, missing, empty, and null are distinct | represent all forms explicitly; no public API decision yet | `Add`, `Close`, `SaveAs`, `Find`, `Sort`, `Run` |\n");
    output.push_str("| P1 | returned dispatch | clients wrap after conversion and own temporary VARIANTs | validate owned Rust wrapper/refcount boundary | object get and `Worksheets.Add` |\n");
    output.push_str("| P2 | SAFEARRAY conversion | client shapes/ownership differ and neither source proves Excel shape | preserve rank/bounds/type evidence before value model design | multi-cell `Value2` after smoke succeeds |\n");
    output.push_str("| P2 | EXCEPINFO | pywin32 fills deferred error and cleans BSTRs; comtypes exposes details | keep cleanup telemetry, add semantic applicability | invalid range/type mismatch controls |\n");
    output.push_str("| P2 | named arguments | Python surface syntax does not prove raw named-DISPID layout | isolate named-DISPID builder separately | `Find`, `SaveAs`, `Sort` |\n");
    output
}

fn unresolved_report() -> String {
    let mut output = report_header("Unresolved client-implementation questions");
    output.push_str("| Question | Status | Next boundary |\n| --- | --- | --- |\n");
    output.push_str("| Why did the original Prompt 05B pywin32 `DispatchEx` control succeed while the current pywin32 311 dynamic and makepy controls returned `0x800A03EC`? | Inconclusive; historical success is preserved and current controls are separately recorded | compare activation/session state without overwriting either observation |\n");
    output.push_str("| Does comtypes generated mode select a vtable or dispinterface for every selected Excel interface on this host? | Initial Application, Workbooks, and Workbook vtable chain is Control-confirmed, but the later recheck is Inconclusive; remaining corpus paths are only Typelib-correlated | extend bounded controls to Worksheet, Worksheets, and Range without conflating source and runtime evidence |\n");
    output.push_str("| Does LCID 0 versus `0x0400` affect this Excel failure? | Not tested | bounded same-frame experiment in Prompt 05D |\n");
    output.push_str("| What Range Value2 SAFEARRAY shapes does this Excel return? | Not tested because no Rust workbook was created | resume Prompt 05 after workbook creation succeeds |\n");
    output
}

fn io_error(error: std::io::Error) -> String {
    error.to_string()
}

#[cfg(test)]
mod tests {
    use super::*;

    fn knowledge_root() -> PathBuf {
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../knowledge/excel-object-model")
    }

    #[test]
    fn mandatory_corpus_has_stable_unique_ids() {
        assert_eq!(CORPUS.len(), 42);
        let ids: BTreeSet<_> = CORPUS.iter().map(|member| member.typelib_id()).collect();
        assert_eq!(ids.len(), CORPUS.len());
        assert!(ids.contains("Excel.TypeLib.member.Excel.Workbooks.Add.INVOKE_FUNC"));
        assert!(ids.contains("Excel.TypeLib.member.Excel.Range.Value2.INVOKE_PROPERTYPUT"));
    }

    #[test]
    fn corpus_joins_installed_typelib_metadata() {
        let typelib = load_typelib(&knowledge_root()).expect("typelib metadata");
        validate_corpus(&typelib).expect("all selected members join");
        let add = typelib
            .members
            .get("Excel.TypeLib.member.Excel.Workbooks.Add.INVOKE_FUNC")
            .expect("Workbooks.Add");
        assert_eq!(add.dispid, 181);
        assert_eq!(add.optional_parameter_count, 1);
    }

    #[test]
    fn mode_classification_distinguishes_invoke_and_vtable_paths() {
        let item = CORPUS
            .iter()
            .find(|member| member.canonical() == "Excel.Workbooks.Item")
            .expect("Item");
        assert_eq!(ClientMode::Pywin32Generated.mode(*item), "makepy-generated");
        assert!(uses_comtypes_vtable(*item));
        let range = CORPUS
            .iter()
            .find(|member| member.canonical() == "Excel.Range.Value2")
            .expect("Value2");
        assert_eq!(
            ClientMode::ComtypesGenerated.mode(*range),
            "generated dispinterface"
        );
    }

    #[test]
    fn invoke_flags_and_property_put_named_dispid_are_normalized() {
        assert_eq!(
            dispatch_flags("INVOKE_FUNC"),
            json!(["DISPATCH_METHOD", "0x0001"])
        );
        assert_eq!(
            dispatch_flags("INVOKE_PROPERTYPUT"),
            json!(["DISPATCH_PROPERTYPUT", "0x0004"])
        );
        let typelib = load_typelib(&knowledge_root()).expect("typelib metadata");
        let member = CORPUS
            .iter()
            .find(|member| {
                member.canonical() == "Excel.Range.Value2"
                    && member.invoke_kind == "INVOKE_PROPERTYPUT"
            })
            .expect("Value2 put");
        let record =
            invocation_record(ClientMode::Pywin32Generated, *member, &typelib).expect("record");
        assert_eq!(record["invoke"]["api"], "PyIDispatch.Invoke");
        assert_eq!(
            record["invoke"]["named_dispids"],
            json!(["DISPID_PROPERTYPUT"])
        );
    }

    #[test]
    fn generated_invoke_and_invoke_types_paths_are_extracted() {
        let typelib = load_typelib(&knowledge_root()).expect("typelib metadata");
        let version = CORPUS
            .iter()
            .find(|member| member.canonical() == "Excel.Application.Version")
            .expect("Version");
        let item = CORPUS
            .iter()
            .find(|member| member.canonical() == "Excel.Workbooks.Item")
            .expect("Item");
        assert_eq!(
            invocation_record(ClientMode::Pywin32Generated, *version, &typelib).expect("Version")["invoke"]
                ["api"],
            "PyIDispatch.InvokeTypes"
        );
        assert_eq!(
            invocation_record(ClientMode::Pywin32Generated, *item, &typelib).expect("Item")["invoke"]
                ["api"],
            "PyIDispatch.Invoke"
        );
    }

    #[test]
    fn dispatch_lcid_and_type_descriptors_are_normalized() {
        let typelib = load_typelib(&knowledge_root()).expect("typelib metadata");
        let save_as = CORPUS
            .iter()
            .find(|member| member.canonical() == "Excel.Workbook.SaveAs")
            .expect("SaveAs");
        let record =
            invocation_record(ClientMode::Pywin32Generated, *save_as, &typelib).expect("record");
        assert_eq!(record["lcid"]["value"], 0);
        let descriptors = record["invoke"]["argument_descriptors"]
            .as_array()
            .expect("descriptor array");
        assert!(!descriptors.is_empty());
        assert_eq!(descriptors[0]["ordinal"], 0);
        assert!(
            descriptors
                .iter()
                .all(|descriptor| descriptor.get("type").is_some())
        );
    }

    #[test]
    fn optional_parameter_metadata_is_preserved() {
        let typelib = load_typelib(&knowledge_root()).expect("typelib metadata");
        let member = CORPUS
            .iter()
            .find(|member| member.canonical() == "Excel.Workbook.SaveAs")
            .expect("SaveAs");
        let record =
            invocation_record(ClientMode::Pywin32Generated, *member, &typelib).expect("record");
        assert!(
            record["typelib_correlation"]["optional_parameter_count"]
                .as_u64()
                .unwrap_or_default()
                > 0
        );
        assert!(
            record["invoke"]["optional_argument_handling"]
                .as_str()
                .unwrap_or_default()
                .contains("Missing")
        );
    }

    #[test]
    fn source_manifests_are_pinned_and_portable() {
        let pywin = source_manifest("pywin32").expect("manifest");
        let comtypes = source_manifest("comtypes").expect("manifest");
        assert!(pywin.contains(PYWIN32_COMMIT));
        assert!(comtypes.contains(COMTYPES_COMMIT));
        assert!(!pywin.contains(":\\"));
        assert!(!comtypes.contains(":\\"));
        assert!(comtypes.contains("client/lazybind.py"));
        assert!(comtypes.contains("locally_installed_version = \"1.4.16\""));
        assert!(pywin.contains(PYWIN32_311_COMMIT));
        assert!(pywin.contains(PYWIN32_312_COMMIT));
        assert!(pywin.contains("complete for the selected Excel Automation paths"));
    }

    #[test]
    fn comtypes_dynamic_records_the_observed_lazybind_path() {
        let typelib = load_typelib(&knowledge_root()).expect("typelib metadata");
        let add = CORPUS
            .iter()
            .find(|member| member.canonical() == "Excel.Workbooks.Add")
            .expect("Workbooks.Add");
        let record =
            invocation_record(ClientMode::ComtypesDynamic, *add, &typelib).expect("record");
        assert_eq!(record["wrapper_module"], "comtypes.client.lazybind");
        assert!(
            record["invoke"]["api"]
                .as_str()
                .unwrap_or_default()
                .contains("_invoke")
        );
    }

    #[test]
    fn reports_are_deterministic() {
        let first = render(&knowledge_root()).expect("first render").files;
        let second = render(&knowledge_root()).expect("second render").files;
        assert_eq!(first, second);
        assert_eq!(first.len(), 33);
        let property_put = first
            .get(&PathBuf::from(
                "generated/client-implementations/property-put-patterns.md",
            ))
            .expect("property put report");
        assert!(property_put.contains("DISPID_PROPERTYPUT"));
    }

    #[test]
    fn pywin32_version_reconciliation_is_complete_and_portable() {
        let records = pywin32_version_reconciliation().expect("records");
        assert!(records.contains("dynamic-cdispatch"));
        assert!(records.contains("materially changed outside bounded corpus"));
        assert!(records.contains("semantically equivalent for local Excel activation"));
        assert!(!records.contains(":\\"));
        assert!(!records.contains("C:\\"));
        let report = pywin32_version_report();
        assert!(report.contains(PYWIN32_311_COMMIT));
        assert!(report.contains("not a claim that an intermittent Excel result"));
    }

    #[test]
    fn generated_reports_reject_known_mojibake_patterns() {
        let reports = render(&knowledge_root()).expect("reports").files;
        for (path, content) in reports {
            if path.starts_with("generated/client-implementations") {
                reject_mojibake(&content, &path).expect("generated report is valid UTF-8 text");
            }
        }
        assert!(reject_mojibake("Application â†’ Workbooks", Path::new("report.md")).is_err());
        assert!(reject_mojibake("replacement ï¿½ marker", Path::new("report.md")).is_err());
    }

    #[test]
    fn historical_prompt_05b_control_is_preserved() {
        let runtime = fs::read_to_string(knowledge_root().join("runtime/observations.jsonl"))
            .expect("runtime evidence");
        assert!(runtime.contains("runtime.control.pywin32-dispatchex-workbooks-add"));
        assert!(runtime.contains("\"created_workbook\":\"Book1\""));
        let environments = environments("pywin32").expect("environment record");
        assert!(environments.contains("prompt-05b-preserved-control"));
        assert!(environments.contains("historical runtime control preserved"));
    }
}
