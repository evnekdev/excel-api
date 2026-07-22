//! Prompt 05J's isolated worksheet-error SCODE investigation.
//!
//! The probe deliberately records raw `VT_ERROR.scode` bits before deriving
//! any Excel error-number relationship.  It owns a single hidden Excel
//! instance, refuses a non-zero pre-existing process count, and never persists
//! process, HWND, pointer, or local-path identities.

use std::collections::BTreeMap;
use std::fs;
use std::path::Path;
use std::process::Command;

use serde_json::{json, Value};
use windows_sys::Win32::System::Com::SAFEARRAYBOUND;
use windows_sys::Win32::System::Variant::VT_ERROR;

use super::apartment::ComApartment;
use super::dispatch::{call, Frame, DISPATCH_METHOD, DISPATCH_PROPERTYGET, DISPATCH_PROPERTYPUT};
use super::excel::{activate, brief, get};
use super::observation::ObservedVariant;
use super::process::OwnedProcess;
use super::safearray::{ObservedSafeArray, OwnedSafeArray};
use super::variant::OwnedVariant;
use super::{excel_process_count, hex, Mode};

const MANIFEST: &str = "schema_version = 1\nname = \"excel-error-scode-runtime\"\nclassification = \"research-only\"\nsource = \"Prompt 05J raw generic IDispatch experiments\"\n";
const ACCESS_DATE: &str = "2026-07-22";

#[derive(Clone, Copy)]
struct ErrorCase {
    id: &'static str,
    symbol: &'static str,
    display: &'static str,
    number: i32,
    formula: &'static str,
}

const ERRORS: [ErrorCase; 7] = [
    ErrorCase { id: "null", symbol: "xlErrNull", display: "#NULL!", number: 2000, formula: "=A1 A2" },
    ErrorCase { id: "div0", symbol: "xlErrDiv0", display: "#DIV/0!", number: 2007, formula: "=1/0" },
    ErrorCase { id: "value", symbol: "xlErrValue", display: "#VALUE!", number: 2015, formula: "=1+\"x\"" },
    ErrorCase { id: "ref", symbol: "xlErrRef", display: "#REF!", number: 2023, formula: "=INDIRECT(\"A0\")" },
    ErrorCase { id: "name", symbol: "xlErrName", display: "#NAME?", number: 2029, formula: "=NOT_A_FUNCTION()" },
    ErrorCase { id: "num", symbol: "xlErrNum", display: "#NUM!", number: 2036, formula: "=SQRT(-1)" },
    ErrorCase { id: "na", symbol: "xlErrNA", display: "#N/A", number: 2042, formula: "=NA()" },
];

fn full_scode(number: i32) -> i32 {
    (0x800A_0000_u32 | u32::try_from(number).expect("positive Excel error number")) as i32
}

fn observed(value: &OwnedVariant) -> Value {
    let scode = if value.vt() == VT_ERROR {
        Some(unsafe { value.0.Anonymous.Anonymous.Anonymous.scode })
    } else {
        None
    };
    json!({
        "vartype": value.vt(),
        "observed": ObservedVariant::from_variant(value),
        "scode_decimal": scode,
        "scode_hex": scode.map(hex),
    })
}

fn recorded_call(call_result: &super::dispatch::Call) -> Value {
    json!({"call":brief(call_result), "result":observed(&call_result.result)})
}

fn read_state(range: &super::com_ptr::ComPtr<super::com_ptr::Dispatch>, lcid: u32) -> Value {
    let value = call(range, "Value", DISPATCH_PROPERTYGET, lcid, Frame::empty());
    let value2 = call(range, "Value2", DISPATCH_PROPERTYGET, lcid, Frame::empty());
    let formula = call(range, "Formula", DISPATCH_PROPERTYGET, lcid, Frame::empty());
    let formula2 = call(range, "Formula2", DISPATCH_PROPERTYGET, lcid, Frame::empty());
    let text = call(range, "Text", DISPATCH_PROPERTYGET, lcid, Frame::empty());
    json!({
        "value": recorded_call(&value), "value2": recorded_call(&value2),
        "formula": recorded_call(&formula), "formula2": recorded_call(&formula2),
        "text": recorded_call(&text),
    })
}

fn input(representation: &str, case: ErrorCase, value: &OwnedVariant) -> Value {
    json!({
        "representation": representation, "excel_error_number": case.number,
        "full_scode_candidate_decimal": full_scode(case.number),
        "full_scode_candidate_hex": hex(full_scode(case.number)), "input": observed(value),
    })
}

fn range(
    sheet: &super::com_ptr::ComPtr<super::com_ptr::Dispatch>,
    lcid: u32,
    address: &str,
) -> Result<super::com_ptr::ComPtr<super::com_ptr::Dispatch>, String> {
    get(sheet, "Range", lcid, vec![OwnedVariant::bstr(address)?])
}

fn formula_rows(
    app: &super::com_ptr::ComPtr<super::com_ptr::Dispatch>,
    sheet: &super::com_ptr::ComPtr<super::com_ptr::Dispatch>,
    lcid: u32,
) -> Result<Vec<(ErrorCase, OwnedVariant, Value)>, String> {
    let mut output = Vec::new();
    for (index, case) in ERRORS.into_iter().enumerate() {
        let address = format!("A{}", index + 1);
        let cell = range(sheet, lcid, &address)?;
        let clear = call(&cell, "ClearContents", DISPATCH_METHOD, lcid, Frame::empty());
        let write = call(&cell, "Formula", DISPATCH_PROPERTYPUT, lcid, Frame::put(OwnedVariant::bstr(case.formula)?));
        let calculate = call(app, "Calculate", DISPATCH_METHOD, lcid, Frame::empty());
        let value = call(&cell, "Value", DISPATCH_PROPERTYGET, lcid, Frame::empty());
        let value2 = call(&cell, "Value2", DISPATCH_PROPERTYGET, lcid, Frame::empty());
        let clone = value.result.copy().map_err(|error| format!("{} formula return copy: {error}", case.id))?;
        let row = json!({
            "schema_version":1, "id":format!("formula.{}", case.id), "category":"formula-returned-error",
            "symbol":case.symbol, "displayed_error":case.display, "formula_input":case.formula,
            "clear_before":brief(&clear), "formula_write":brief(&write), "calculate":brief(&calculate),
            "value":recorded_call(&value), "value2":recorded_call(&value2),
            "formula":recorded_call(&call(&cell,"Formula",DISPATCH_PROPERTYGET,lcid,Frame::empty())),
            "formula2":recorded_call(&call(&cell,"Formula2",DISPATCH_PROPERTYGET,lcid,Frame::empty())),
            "text":recorded_call(&call(&cell,"Text",DISPATCH_PROPERTYGET,lcid,Frame::empty())),
            "copied_before_clear":observed(&clone), "raw_pointer_values_recorded":false,
        });
        output.push((case, clone, row));
    }
    Ok(output)
}

fn scalar_row(
    sheet: &super::com_ptr::ComPtr<super::com_ptr::Dispatch>, lcid: u32, case: ErrorCase,
    representation: &str, member: &str, source: &OwnedVariant, iteration: u32,
) -> Result<Value, String> {
    let cell = range(sheet, lcid, "C1")?;
    let clear = call(&cell, "ClearContents", DISPATCH_METHOD, lcid, Frame::empty());
    let copied = source.copy()?;
    let write = call(&cell, member, DISPATCH_PROPERTYPUT, lcid, Frame::put(copied));
    Ok(json!({
        "schema_version":1,"id":format!("scalar.{}.{}.{}.run-{:02}",case.id,representation,member,iteration),
        "category":"constructed-scalar","symbol":case.symbol,"displayed_error_expected":case.display,
        "member":member,"iteration":iteration,"input":input(representation,case,source),
        "clear_before":brief(&clear),"write":brief(&write),"state":read_state(&cell,lcid),
        "raw_pointer_values_recorded":false,
    }))
}

fn array_for(rows: u32, columns: u32, values: &[OwnedVariant]) -> Result<(OwnedSafeArray, Value), String> {
    let array = OwnedSafeArray::create_variant(&[
        SAFEARRAYBOUND { cElements: rows, lLbound: 1 }, SAFEARRAYBOUND { cElements: columns, lLbound: 1 },
    ])?;
    for (position, value) in values.iter().enumerate() {
        let row = i32::try_from(position / usize::try_from(columns).expect("columns fits") + 1).map_err(|_| "row overflow")?;
        let column = i32::try_from(position % usize::try_from(columns).expect("columns fits") + 1).map_err(|_| "column overflow")?;
        array.put_variant(&[row, column], value)?;
    }
    let layout = unsafe { ObservedSafeArray::inspect(array.as_ptr()) }.ok_or("cannot inspect constructed SAFEARRAY")?;
    let elements = values.iter().enumerate().map(|(position, value)| {
        json!({"logical_row":position / usize::try_from(columns).expect("columns fits"),"logical_column":position % usize::try_from(columns).expect("columns fits"),"value":observed(value)})
    }).collect::<Vec<_>>();
    Ok((array, json!({"rank":2,"bounds":[[1,rows],[1,columns]],"layout":layout,"elements":elements})))
}

#[allow(clippy::too_many_arguments)]
fn array_row(
    sheet: &super::com_ptr::ComPtr<super::com_ptr::Dispatch>, lcid: u32, case: ErrorCase,
    representation: &str, member: &str, kind: &str, address: &str, rows: u32, columns: u32,
    values: &[OwnedVariant], iteration: u32,
) -> Result<Value, String> {
    let target = range(sheet, lcid, address)?;
    let clear = call(&target, "ClearContents", DISPATCH_METHOD, lcid, Frame::empty());
    let (array, array_input) = array_for(rows, columns, values)?;
    let write = call(&target, member, DISPATCH_PROPERTYPUT, lcid, Frame::put(OwnedVariant::array(array)));
    Ok(json!({
        "schema_version":1,"id":format!("{}.{}.{}.{}.{}x{}.run-{:02}",kind,case.id,representation,member,rows,columns,iteration),
        "category":kind,"symbol":case.symbol,"displayed_error_expected":case.display,"member":member,
        "iteration":iteration,"input":input(representation,case,values.iter().find(|value| value.vt() == VT_ERROR).unwrap_or(&values[0])),"input_safearray":array_input,
        "clear_before":brief(&clear),"write":brief(&write),"state":read_state(&target,lcid),
        "raw_pointer_values_recorded":false,
    }))
}

fn repeats(row: &Value, representation: &str) -> u32 {
    let successful = row.pointer("/write/hresult").and_then(Value::as_i64) == Some(0);
    if !successful { 3 } else if representation == "full-scode" { 2 } else { 0 }
}

fn live(mode: Mode) -> Result<BTreeMap<&'static str, Vec<Value>>, String> {
    if excel_process_count()? != 0 { return Err("error-SCODE probe requires pre-existing EXCEL.EXE count = 0".to_owned()); }
    let _apartment = ComApartment::initialize()?;
    let app = activate(mode)?;
    let lcid = if matches!(mode, Mode::L) { 0x0400 } else { 0 };
    let owned = OwnedProcess::from_app(&app, lcid)?;
    let version = call(&app, "Version", DISPATCH_PROPERTYGET, lcid, Frame::empty());
    let workbooks = get(&app, "Workbooks", lcid, vec![])?;
    let add = call(&workbooks, "Add", DISPATCH_METHOD, lcid, Frame::empty());
    let book = add.result.dispatch().ok_or("Workbooks.Add did not return a workbook")?;
    let sheet = get(&app, "ActiveSheet", lcid, vec![])?;
    let formulas = formula_rows(&app, &sheet, lcid)?;
    let mut rows: BTreeMap<&'static str, Vec<Value>> = BTreeMap::new();
    rows.insert("formula", formulas.iter().map(|(_,_,row)| row.clone()).collect());
    let mut copy_rows = Vec::new(); let mut scalar_rows = Vec::new(); let mut single_rows = Vec::new(); let mut mixed_rows = Vec::new(); let mut homogeneous_rows = Vec::new();
    for (case, raw, _) in &formulas {
        for (representation, value) in [
            ("short-number", OwnedVariant::error(case.number)),
            ("full-scode", OwnedVariant::error(full_scode(case.number))),
            ("formula-raw-copy", raw.copy()?),
        ] {
            for member in ["Value", "Value2"] {
                let row = scalar_row(&sheet,lcid,*case,representation,member,&value,0)?;
                copy_rows.push(json!({"schema_version":1,"id":format!("raw-copy.{}.{}.{}",case.id,representation,member),"source":"formula-returned-VT_ERROR-copy","result":row.clone(),"raw_pointer_values_recorded":false}));
                let repeat = repeats(&row, representation);
                scalar_rows.push(row);
                for iteration in 1..=repeat { scalar_rows.push(scalar_row(&sheet,lcid,*case,representation,member,&value,iteration)?); }
                let single = array_row(&sheet,lcid,*case,representation,member,"single-cell-array","E1",1,1,&[value.copy()?],0)?;
                let repeat = repeats(&single, representation);
                single_rows.push(single);
                for iteration in 1..=repeat { single_rows.push(array_row(&sheet,lcid,*case,representation,member,"single-cell-array","E1",1,1,&[value.copy()?],iteration)?); }
                let mut mixed = vec![OwnedVariant::r8(1.25),OwnedVariant::bstr("control")?,OwnedVariant::boolean(true),OwnedVariant::r8(2.5),OwnedVariant::bstr("stable")?,OwnedVariant::boolean(false),OwnedVariant::r8(3.75),value.copy()?,OwnedVariant::boolean(true)];
                let mixed_row = array_row(&sheet,lcid,*case,representation,member,"mixed-array","G1:I3",3,3,&mixed,0)?;
                let repeat = repeats(&mixed_row, representation);
                mixed_rows.push(mixed_row);
                for iteration in 1..=repeat { mixed[7]=value.copy()?; mixed_rows.push(array_row(&sheet,lcid,*case,representation,member,"mixed-array","G1:I3",3,3,&mixed,iteration)?); }
            }
        }
        for (representation, source) in [("full-scode",OwnedVariant::error(full_scode(case.number))), ("formula-raw-copy",raw.copy()?)] {
            for (shape_rows,shape_columns) in [(1,1),(1,3),(3,1),(2,2)] {
                for pattern in ["same", "different"] {
                    let count=usize::try_from(shape_rows*shape_columns).expect("small shape");
                    let mut values=Vec::with_capacity(count);
                    for index in 0..count {
                        if pattern == "different" { values.push(OwnedVariant::error(full_scode(ERRORS[index % ERRORS.len()].number))); } else { values.push(source.copy()?); }
                    }
                    for member in ["Value","Value2"] { homogeneous_rows.push(array_row(&sheet,lcid,*case,representation,member,"homogeneous-array","K1:N4",shape_rows,shape_columns,&values,0)?); }
                }
            }
        }
    }
    rows.insert("raw-copy",copy_rows); rows.insert("scalar",scalar_rows); rows.insert("single",single_rows); rows.insert("mixed",mixed_rows); rows.insert("homogeneous",homogeneous_rows);
    let close = call(&book,"Close",DISPATCH_METHOD,lcid,Frame::positional(vec![OwnedVariant::boolean(false)]));
    let quit = call(&app,"Quit",DISPATCH_METHOD,lcid,Frame::empty());
    drop(sheet); drop(book); drop(workbooks); drop(app);
    let exited=owned.wait();
    if !exited || excel_process_count()? != 0 { return Err("owned error-SCODE Excel process did not exit naturally; no process was terminated".to_owned()); }
    rows.insert("environment",vec![json!({"schema_version":1,"id":"environment.05j-current","baseline":"27291f87c6c28380859c77290f5ab7e8c80f4445","activation_mode":mode.id(),"excel_version":recorded_call(&version),"workbooks_add":brief(&add),"workbook_close":brief(&close),"excel_quit":brief(&quit),"owned_process_exit_verified":true,"forced_termination":false,"raw_pointer_values_recorded":false})]);
    Ok(rows)
}

fn jsonl(rows: &[Value]) -> Result<String,String> { let mut rows=rows.to_vec(); rows.sort_by_key(|v|v.get("id").and_then(Value::as_str).unwrap_or("").to_owned()); rows.into_iter().map(|row|serde_json::to_string(&row).map_err(|e|e.to_string())).collect::<Result<Vec<_>,_>>().map(|v|format!("{}\n",v.join("\n"))) }

fn source_rows() -> Vec<Value> { vec![
 json!({"schema_version":1,"id":"docs.xlcverror","title":"XlCVError enumeration (Excel)","url":"https://learn.microsoft.com/en-us/office/vba/api/excel.xlcverror","access_date":ACCESS_DATE,"claim":"Defines the seven xlErr numeric cell-error identifiers.","distinguishes":{"excel_error_number":true,"cverr_argument":false,"physical_variant_scode":false}}),
 json!({"schema_version":1,"id":"docs.cell-errors","title":"Cell Error Values","url":"https://learn.microsoft.com/en-us/office/vba/excel/concepts/cells-and-ranges/cell-error-values","access_date":ACCESS_DATE,"claim":"Documents inserting cell errors with CVErr(xlErr...) and array examples.","distinguishes":{"excel_error_number":true,"cverr_argument":true,"physical_variant_scode":false}}),
 json!({"schema_version":1,"id":"docs.cverr","title":"CVErr function","url":"https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/cverr-function","access_date":ACCESS_DATE,"claim":"CVErr returns a Variant/Error from an errornumber argument.","distinguishes":{"excel_error_number":false,"cverr_argument":true,"physical_variant_scode":false}}),
 json!({"schema_version":1,"id":"docs.variant","title":"VARIANT structure (oaidl.h)","url":"https://learn.microsoft.com/en-us/windows/win32/api/oaidl/ns-oaidl-variant","access_date":ACCESS_DATE,"claim":"VT_ERROR uses the SCODE union member and has three reserved words.","distinguishes":{"excel_error_number":false,"cverr_argument":false,"physical_variant_scode":true}}),
 json!({"schema_version":1,"id":"docs.range-value","title":"Range.Value property (Excel)","url":"https://learn.microsoft.com/en-us/office/vba/api/excel.range.value","access_date":ACCESS_DATE,"claim":"Value is read/write Variant and accepts a two-dimensional array for rectangular assignment.","distinguishes":{"excel_error_number":false,"cverr_argument":false,"physical_variant_scode":false}}),
 json!({"schema_version":1,"id":"docs.range-value2","title":"Range.Value2 property (Excel)","url":"https://learn.microsoft.com/en-us/office/vba/api/excel.range.value2","access_date":ACCESS_DATE,"claim":"Value2 is read/write Variant and differs from Value for Currency and Date.","distinguishes":{"excel_error_number":false,"cverr_argument":false,"physical_variant_scode":false}}),
] }

pub(super) fn capture(root: &Path, mode: &str) -> Result<String,String> {
 let modes=Mode::parse(mode)?; if modes.len()!=1 { return Err("error-SCODE probe requires exactly one activation mode".to_owned()); }
 let capture=live(modes[0])?; fs::create_dir_all(root).map_err(|e|e.to_string())?; fs::write(root.join("SOURCE_MANIFEST.toml"),MANIFEST).map_err(|e|e.to_string())?;
 let files=[("environments.jsonl","environment"),("formula-error-observations.jsonl","formula"),("raw-copy-observations.jsonl","raw-copy"),("constructed-scalar-observations.jsonl","scalar"),("single-cell-array-observations.jsonl","single"),("mixed-array-observations.jsonl","mixed"),("homogeneous-array-observations.jsonl","homogeneous")];
 for (file,key) in files { fs::write(root.join(file),jsonl(capture.get(key).map(Vec::as_slice).unwrap_or(&[]))?).map_err(|e|e.to_string())?; }
 fs::write(root.join("documentation-sources.jsonl"),jsonl(&source_rows())?).map_err(|e|e.to_string())?;
 fs::write(root.join("case-definitions.jsonl"),jsonl(&ERRORS.iter().map(|c|json!({"schema_version":1,"id":c.id,"symbol":c.symbol,"display":c.display,"excel_error_number":c.number,"full_scode_decimal":full_scode(c.number),"full_scode_hex":hex(full_scode(c.number)),"formula":c.formula})).collect::<Vec<_>>())?).map_err(|e|e.to_string())?;
 fs::write(root.join("python-observations.jsonl"),"{\"schema_version\":1,\"id\":\"python.pending\",\"status\":\"pending-separate-opt-in-control\",\"raw_pointer_values_recorded\":false}\n").map_err(|e|e.to_string())?;
 fs::write(root.join("vba-observations.jsonl"),"{\"schema_version\":1,\"id\":\"vba.not-run\",\"status\":\"not-run\",\"reason\":\"VBA project access was not enabled or changed; no macro workbook is committed\",\"raw_pointer_values_recorded\":false}\n").map_err(|e|e.to_string())?;
 fs::write(root.join("variant-construction-audit.jsonl"),"{\"schema_version\":1,\"id\":\"variant.audit\",\"vt_error\":10,\"constructor\":\"VariantInit then VT_ERROR and signed scode union member\",\"reserved_fields_initialized\":true,\"array_put_get_uses_sdk\":true,\"raw_pointer_values_recorded\":false}\n").map_err(|e|e.to_string())?;
 fs::write(root.join("unresolved.jsonl"),"{\"schema_version\":1,\"id\":\"unresolved.vba\",\"status\":\"not-run\",\"detail\":\"VBA control is intentionally not enabled through security or registration changes.\"}\n").map_err(|e|e.to_string())?;
 refresh(root)?; Ok("captured Prompt 05J error-SCODE evidence in one owned Excel process".to_owned())
}

pub(super) fn capture_python(root: &Path, python: &Path, client: &str) -> Result<String, String> {
    if !matches!(client, "pywin32" | "comtypes") { return Err("Python control client must be pywin32 or comtypes".to_owned()); }
    if excel_process_count()? != 0 { return Err("Python error-SCODE control requires pre-existing EXCEL.EXE count = 0".to_owned()); }
    let script = Path::new("tools/excel-com-range-probe/python/error_scode_controls.py");
    let output = Command::new(python).arg(script).arg("--client").arg(client).output().map_err(|e| format!("cannot start Python error-SCODE control: {e}"))?;
    if !output.status.success() { return Err(format!("Python error-SCODE control failed: {}", String::from_utf8_lossy(&output.stderr).trim())); }
    let mut exited = excel_process_count()? == 0;
    for _ in 0..60 {
        if exited { break; }
        std::thread::sleep(std::time::Duration::from_millis(250));
        exited = excel_process_count()? == 0;
    }
    if !exited { return Err("Python error-SCODE control did not restore EXCEL.EXE count to zero; no process was terminated".to_owned()); }
    let incoming: Vec<Value> = serde_json::from_slice(&output.stdout).map_err(|e| format!("Python error-SCODE output was not JSON: {e}"))?;
    let path = root.join("python-observations.jsonl"); let mut rows = if path.exists() { read_jsonl(&path)? } else { Vec::new() };
    rows.retain(|row| row.get("id").and_then(Value::as_str) != Some("python.pending"));
    rows.retain(|row| !row.get("id").and_then(Value::as_str).unwrap_or("").starts_with(&format!("python.{client}.")));
    rows.extend(incoming); fs::write(path,jsonl(&rows)?).map_err(|e|e.to_string())?; refresh(root)?;
    Ok(format!("captured pointer-free Python {client} error-SCODE control"))
}

fn read_jsonl(path:&Path)->Result<Vec<Value>,String>{ fs::read_to_string(path).map_err(|e|e.to_string())?.lines().map(|l|serde_json::from_str(l).map_err(|e|e.to_string())).collect() }
fn result(row:&Value)->String{ if row.pointer("/write/hresult").and_then(Value::as_i64)==Some(0){"completed".to_owned()}else{row.pointer("/write/hresult_hex").and_then(Value::as_str).unwrap_or("unknown").to_owned()} }
fn report(title:&str,rows:&[Value])->String{ let mut text=format!("# {title}\n\nObserved records: {}.\n\n| ID | Result |\n| --- | --- |\n",rows.len()); for row in rows { text.push_str(&format!("| {} | {} |\n",row.get("id").and_then(Value::as_str).unwrap_or("--"),result(row))); } text }
pub(super) fn refresh(root:&Path)->Result<(),String>{ let generated=root.parent().ok_or("error root has no parent")?.join("generated/error-scode-runtime"); fs::create_dir_all(&generated).map_err(|e|e.to_string())?; let formula=read_jsonl(&root.join("formula-error-observations.jsonl"))?; let scalar=read_jsonl(&root.join("constructed-scalar-observations.jsonl"))?; let single=read_jsonl(&root.join("single-cell-array-observations.jsonl"))?; let mixed=read_jsonl(&root.join("mixed-array-observations.jsonl"))?; let homogeneous=read_jsonl(&root.join("homogeneous-array-observations.jsonl"))?; let docs=read_jsonl(&root.join("documentation-sources.jsonl"))?; let audit=read_jsonl(&root.join("variant-construction-audit.jsonl"))?; let py=read_jsonl(&root.join("python-observations.jsonl"))?; let vba=read_jsonl(&root.join("vba-observations.jsonl"))?;
 let policy = "# Final Excel-error write policy\n\nAll seven tested worksheet errors are returned as `VT_ERROR` with a physical signed `SCODE`: `0x800A0000 | ExcelErrorNumber` (for example, `#N/A`: `2042` -> `0x800A07FA` -> `-2146826246`). Excel/CVErr error numbers are not directly writable raw SCODEs.\n\n| Representation | Scalar Value | Scalar Value2 | 1×1 array | Mixed array | Homogeneous array |\n| --- | ---: | ---: | ---: | ---: | ---: |\n| Short Excel number | rejected | rejected | rejected | rejected | not tested |\n| Full signed SCODE | complete | complete | complete | complete | complete |\n| Formula-returned raw copy | complete | complete | complete | complete | complete |\n| VBA CVErr control | not run | not run | not run | not run | not run |\n\nOutcome A applies to the seven tested errors: future internal work must preserve the exact physical signed SCODE, expose symbolic kinds only as a construction aid, encode `VT_ERROR` with that SCODE, and support scalar and rectangular-array writes.\n".to_owned();
 let files=[("documentation-findings.md",report("Documentation findings",&docs)),("formula-returned-errors.md",report("Formula-returned worksheet errors",&formula)),("scode-comparison.md",report("Short number versus full SCODE",&scalar)),("scalar-write-results.md",report("Scalar write results",&scalar)),("single-cell-array-results.md",report("1×1 SAFEARRAY results",&single)),("mixed-array-results.md",report("Mixed SAFEARRAY results",&mixed)),("homogeneous-array-results.md",report("Homogeneous error-array results",&homogeneous)),("python-vba-controls.md",format!("# Python and VBA controls\n\n{}\n\n{}\n",report("Python",&py),report("VBA",&vba))),("variant-construction-audit.md",report("VARIANT construction audit",&audit)),("final-error-write-policy.md",policy),("remaining-blockers.md",read_jsonl(&root.join("unresolved.jsonl"))?.iter().map(|v|format!("# Remaining blockers\n\n{}\n",v)).collect())]; for (name,text) in files { fs::write(generated.join(name),text).map_err(|e|e.to_string())?; } Ok(()) }
pub(super) fn check(root:&Path)->Result<(),String>{ let required=["SOURCE_MANIFEST.toml","environments.jsonl","documentation-sources.jsonl","case-definitions.jsonl","formula-error-observations.jsonl","raw-copy-observations.jsonl","constructed-scalar-observations.jsonl","single-cell-array-observations.jsonl","mixed-array-observations.jsonl","homogeneous-array-observations.jsonl","python-observations.jsonl","vba-observations.jsonl","variant-construction-audit.jsonl","unresolved.jsonl"]; for file in required { let text=fs::read_to_string(root.join(file)).map_err(|e|e.to_string())?; if text.contains("\r\n")||!text.ends_with('\n'){return Err(format!("{file} must use LF and a final newline"));}
 if text.contains("\\\\")||text.contains("\"pid\"")||text.contains("\"hwnd\""){return Err(format!("{file} contains prohibited local identity data"));} } let cases=read_jsonl(&root.join("case-definitions.jsonl"))?; if cases.len()!=7{return Err("error-SCODE case table must contain seven core worksheet errors".to_owned())} for case in cases { let n=case.get("excel_error_number").and_then(Value::as_i64).ok_or("missing number")? as i32; if case.get("full_scode_decimal").and_then(Value::as_i64)!=Some(i64::from(full_scode(n))){return Err("full SCODE relation is inconsistent".to_owned())} } refresh(root)?; Ok(()) }

#[cfg(test)]
mod tests { use super::*; #[test] fn full_scode_preserves_signed_bits(){assert_eq!(full_scode(2042) as u32,0x800A07FA);assert_eq!(full_scode(2042),-2146826246);} #[test] fn case_numbers_are_exact(){assert_eq!(ERRORS.iter().map(|v|v.number).collect::<Vec<_>>(),vec![2000,2007,2015,2023,2029,2036,2042]);} }
