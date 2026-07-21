//! Canonical raw `Range.Value2` scalar observations for Prompt 05H.

use core::ffi::c_void;

use serde_json::{json, Value};
use windows_sys::Win32::System::Com::SAFEARRAYBOUND;
use windows_sys::Win32::System::Ole::SafeArrayGetElement;
use windows_sys::Win32::System::Variant::{VARIANT, VT_ARRAY};

use super::apartment::ComApartment;
use super::dispatch::{call, Frame, DISPATCH_METHOD, DISPATCH_PROPERTYGET, DISPATCH_PROPERTYPUT};
use super::excel::{activate, brief, get};
use super::observation::ObservedVariant;
use super::process::OwnedProcess;
use super::safearray::{
    checked_element_count, row_column_indices, ObservedSafeArray, OwnedSafeArray,
};
use super::variant::OwnedVariant;
use super::Mode;

struct ScalarCase {
    id: &'static str,
    semantic_input: &'static str,
    input: OwnedVariant,
}

/// Runs the required Value2 scalar writes in one owned workbook and returns
/// only copied, pointer-free observations.
pub(super) fn scalar_value2_case_ids() -> &'static [&'static str] {
    &[
        "S-V2-01", "S-V2-02", "S-V2-03", "S-V2-04", "S-V2-05", "S-V2-06", "S-V2-07",
        "S-V2-08", "S-V2-09", "S-V2-10", "S-V2-11", "S-V2-12", "S-V2-13", "S-V2-14",
        "S-V2-15", "S-V2-16", "S-V2-17", "S-V2-18", "S-V2-19", "S-V2-20", "S-V2-21",
        "S-V2-22", "S-V2-23", "S-V2-24", "S-V2-25", "S-V2-26", "S-V2-27", "S-V2-28",
    ]
}

pub(super) fn scalar_value2(mode: Mode, only_case: Option<&str>) -> Result<Value, String> {
    scalar_values(mode, "Value2", "S-V2", only_case)
}

pub(super) fn scalar_value_case_ids() -> Vec<String> {
    scalar_value2_case_ids()
        .iter()
        .map(|id| id.replacen("S-V2", "S-V", 1))
        .collect()
}

pub(super) fn scalar_value(mode: Mode, only_case: Option<&str>) -> Result<Value, String> {
    scalar_values(mode, "Value", "S-V", only_case)
}

fn scalar_values(
    mode: Mode,
    input_member: &str,
    id_prefix: &str,
    only_case: Option<&str>,
) -> Result<Value, String> {
    let _apartment = ComApartment::initialize()?;
    let app = activate(mode)?;
    let lcid = if matches!(mode, Mode::L) { 0x0400 } else { 0 };
    let owned = OwnedProcess::from_app(&app, lcid)?;
    let workbooks = get(&app, "Workbooks", lcid, vec![])?;
    let add = call(&workbooks, "Add", DISPATCH_METHOD, lcid, Frame::empty());
    let workbook = add
        .result
        .dispatch()
        .ok_or_else(|| format!("Workbooks.Add did not return VT_DISPATCH: {}", add.hr))?;
    let sheet = get(&app, "ActiveSheet", lcid, vec![])?;
    let target = get(&sheet, "Range", lcid, vec![OwnedVariant::bstr("A1")?])?;
    let cases = vec![
        ScalarCase { id: "S-V2-01", semantic_input: "VT_EMPTY", input: OwnedVariant::empty() },
        ScalarCase { id: "S-V2-02", semantic_input: "VT_NULL", input: OwnedVariant::null() },
        ScalarCase { id: "S-V2-03", semantic_input: "VT_BOOL(false)", input: OwnedVariant::boolean(false) },
        ScalarCase { id: "S-V2-04", semantic_input: "VT_BOOL(true)", input: OwnedVariant::boolean(true) },
        ScalarCase { id: "S-V2-05", semantic_input: "VT_I2(-123)", input: OwnedVariant::i2(-123) },
        ScalarCase { id: "S-V2-06", semantic_input: "VT_I2(123)", input: OwnedVariant::i2(123) },
        ScalarCase { id: "S-V2-07", semantic_input: "VT_I4(-42)", input: OwnedVariant::i4(-42) },
        ScalarCase { id: "S-V2-08", semantic_input: "VT_I4(0)", input: OwnedVariant::i4(0) },
        ScalarCase { id: "S-V2-09", semantic_input: "VT_I4(42)", input: OwnedVariant::i4(42) },
        ScalarCase { id: "S-V2-10", semantic_input: "VT_I4(2147483647)", input: OwnedVariant::i4(i32::MAX) },
        ScalarCase { id: "S-V2-11", semantic_input: "VT_I8(42)", input: OwnedVariant::i8(42) },
        ScalarCase { id: "S-V2-12", semantic_input: "VT_I8(9007199254740993)", input: OwnedVariant::i8(9_007_199_254_740_993) },
        ScalarCase { id: "S-V2-13", semantic_input: "VT_R4(1.25)", input: OwnedVariant::r4(1.25) },
        ScalarCase { id: "S-V2-14", semantic_input: "VT_R8(1.25)", input: OwnedVariant::r8(1.25) },
        ScalarCase { id: "S-V2-15", semantic_input: "VT_R8(-1.25)", input: OwnedVariant::r8(-1.25) },
        ScalarCase { id: "S-V2-16", semantic_input: "VT_R8(-0)", input: OwnedVariant::r8(-0.0) },
        ScalarCase { id: "S-V2-17", semantic_input: "VT_R8(NaN)", input: OwnedVariant::r8(f64::NAN) },
        ScalarCase { id: "S-V2-18", semantic_input: "VT_R8(+Infinity)", input: OwnedVariant::r8(f64::INFINITY) },
        ScalarCase { id: "S-V2-19", semantic_input: "VT_R8(-Infinity)", input: OwnedVariant::r8(f64::NEG_INFINITY) },
        ScalarCase { id: "S-V2-20", semantic_input: "VT_BSTR(empty)", input: OwnedVariant::bstr("")? },
        ScalarCase { id: "S-V2-21", semantic_input: "VT_BSTR(ASCII)", input: OwnedVariant::bstr("plain ASCII")? },
        ScalarCase { id: "S-V2-22", semantic_input: "VT_BSTR(BMP Unicode)", input: OwnedVariant::bstr("Grüße Ω")? },
        ScalarCase { id: "S-V2-23", semantic_input: "VT_BSTR(supplementary Unicode)", input: OwnedVariant::bstr("snowman 😀")? },
        ScalarCase { id: "S-V2-24", semantic_input: "VT_BSTR(embedded NUL)", input: OwnedVariant::bstr("left\0right")? },
        ScalarCase { id: "S-V2-25", semantic_input: "VT_ERROR(2042 #N/A)", input: OwnedVariant::error(2042) },
        ScalarCase { id: "S-V2-26", semantic_input: "VT_ERROR(DISP_E_PARAMNOTFOUND)", input: OwnedVariant::error(-2_147_352_572) },
        ScalarCase { id: "S-V2-27", semantic_input: "VT_DATE(45292.5)", input: OwnedVariant::date(45_292.5) },
        ScalarCase { id: "S-V2-28", semantic_input: "VT_CY(123.4500)", input: OwnedVariant::currency(1_234_500) },
    ];
    let mut observations = Vec::with_capacity(cases.len());
    let mut found = only_case.is_none();
    for case in cases {
        let observation_id = case.id.replacen("S-V2", id_prefix, 1);
        if only_case.is_some_and(|id| id != observation_id) {
            continue;
        }
        found = true;
        let input_vartype = case.input.vt();
        let clear_before = call(&target, "ClearContents", DISPATCH_METHOD, lcid, Frame::empty());
        let write = call(
            &target,
            input_member,
            DISPATCH_PROPERTYPUT,
            lcid,
            Frame::put(case.input),
        );
        let read_value2 = call(&target, "Value2", DISPATCH_PROPERTYGET, lcid, Frame::empty());
        let read_value = call(&target, "Value", DISPATCH_PROPERTYGET, lcid, Frame::empty());
        let clear_after = call(&target, "ClearContents", DISPATCH_METHOD, lcid, Frame::empty());
        observations.push(json!({
            "schema_version": 1,
            "id": observation_id,
            "activation_mode": mode.id(),
            "member": input_member,
            "input_vartype": input_vartype,
            "input_semantic_value": case.semantic_input,
            "clear_before": brief(&clear_before),
            "write": brief(&write),
            "value2_read": ObservedVariant::from_variant(&read_value2.result),
            "value2_read_call": brief(&read_value2),
            "value_read": ObservedVariant::from_variant(&read_value.result),
            "value_read_call": brief(&read_value),
            "clear_after": brief(&clear_after),
            "raw_pointer_values_recorded": false,
        }));
    }
    if !found {
        return Err(format!(
            "unknown scalar Value2 case: {}",
            only_case.unwrap_or("--")
        ));
    }
    let close = call(
        &workbook,
        "Close",
        DISPATCH_METHOD,
        lcid,
        Frame::positional(vec![OwnedVariant::boolean(false)]),
    );
    let quit = call(&app, "Quit", DISPATCH_METHOD, lcid, Frame::empty());
    drop(sheet);
    drop(workbook);
    drop(workbooks);
    drop(app);
    let exited = owned.wait();
    Ok(json!({
        "schema_version": 1,
        "backend": "raw-windows-sys",
        "activation_mode": mode.id(),
        "workbooks_add": brief(&add),
        "observations": observations,
        "cleanup": {
            "workbook_close": brief(&close),
            "excel_quit": brief(&quit),
            "owned_process_exit_verified": exited,
            "forced_termination": false,
        },
        "success": exited,
    }))
}

/// Reads the required rectangular shapes through both `Value2` and `Value`.
/// Cells are populated individually with row/column marker values before the
/// rectangular read, so the observed SAFEARRAY index order can be proven
/// rather than assumed from a descriptor layout.
pub(super) fn rectangular_reads(mode: Mode) -> Result<Value, String> {
    let _apartment = ComApartment::initialize()?;
    let app = activate(mode)?;
    let lcid = if matches!(mode, Mode::L) { 0x0400 } else { 0 };
    let owned = OwnedProcess::from_app(&app, lcid)?;
    let workbooks = get(&app, "Workbooks", lcid, vec![])?;
    let add = call(&workbooks, "Add", DISPATCH_METHOD, lcid, Frame::empty());
    let workbook = add
        .result
        .dispatch()
        .ok_or_else(|| format!("Workbooks.Add did not return VT_DISPATCH: {}", add.hr))?;
    let sheet = get(&app, "ActiveSheet", lcid, vec![])?;
    let mut observations = Vec::new();

    for (id, rows, columns) in [
        ("R-01", 1_u32, 1_u32),
        ("R-02", 1, 2),
        ("R-03", 1, 5),
        ("R-04", 2, 1),
        ("R-05", 5, 1),
        ("R-06", 2, 2),
        ("R-07", 2, 3),
        ("R-08", 3, 2),
        ("R-09", 3, 4),
    ] {
        let address = range_address(1, 1, rows, columns);
        let target = get(&sheet, "Range", lcid, vec![OwnedVariant::bstr(&address)?])?;
        let clear_before = call(&target, "ClearContents", DISPATCH_METHOD, lcid, Frame::empty());
        for row in 0..rows {
            for column in 0..columns {
                let cell = cell_address(row + 1, column + 1);
                let cell = get(&sheet, "Range", lcid, vec![OwnedVariant::bstr(&cell)?])?;
                let marker = i32::try_from((row + 1) * 1000 + column + 1)
                    .map_err(|_| "marker conversion overflow")?;
                let write = call(
                    &cell,
                    "Value2",
                    DISPATCH_PROPERTYPUT,
                    lcid,
                    Frame::put(OwnedVariant::i4(marker)),
                );
                if write.hr != 0 {
                    return Err(format!("{id} marker write failed: {}", write.hr));
                }
            }
        }
        let value2 = call(&target, "Value2", DISPATCH_PROPERTYGET, lcid, Frame::empty());
        let value = call(&target, "Value", DISPATCH_PROPERTYGET, lcid, Frame::empty());
        let clear_after = call(&target, "ClearContents", DISPATCH_METHOD, lcid, Frame::empty());
        observations.push(json!({
            "schema_version": 1,
            "id": id,
            "activation_mode": mode.id(),
            "fresh_process": true,
            "range_address": address,
            "rows": rows,
            "columns": columns,
            "content_pattern": "unique_row_column_markers",
            "clear_before": brief(&clear_before),
            "value2_read_call": brief(&value2),
            "value2_read": rectangular_observation(&value2.result),
            "value_read_call": brief(&value),
            "value_read": rectangular_observation(&value.result),
            "clear_after": brief(&clear_after),
            "raw_pointer_values_recorded": false,
        }));
    }

    for (id, values) in [
        (
            "R-P-text",
            vec![
                CellValue::Text("alpha"),
                CellValue::Text("beta"),
                CellValue::Text("gamma"),
                CellValue::Text("delta"),
            ],
        ),
        (
            "R-P-unicode",
            vec![
                CellValue::Text("Grüße Ω"),
                CellValue::Text("😀"),
                CellValue::Text("e\u{301}"),
                CellValue::Text("雪"),
            ],
        ),
        (
            "R-P-mixed",
            vec![
                CellValue::I4(1),
                CellValue::Text("text"),
                CellValue::Bool(true),
                CellValue::Empty,
            ],
        ),
        (
            "R-P-errors",
            vec![
                CellValue::Formula("=1/0"),
                CellValue::Formula("=NA()"),
                CellValue::Formula("=SQRT(-1)"),
                CellValue::Empty,
            ],
        ),
    ] {
        observations.push(read_pattern(&sheet, lcid, mode, id, 2, 2, values)?);
    }

    let close = call(
        &workbook,
        "Close",
        DISPATCH_METHOD,
        lcid,
        Frame::positional(vec![OwnedVariant::boolean(false)]),
    );
    let quit = call(&app, "Quit", DISPATCH_METHOD, lcid, Frame::empty());
    drop(sheet);
    drop(workbook);
    drop(workbooks);
    drop(app);
    let exited = owned.wait();
    Ok(json!({
        "schema_version": 1,
        "backend": "raw-windows-sys",
        "activation_mode": mode.id(),
        "workbooks_add": brief(&add),
        "observations": observations,
        "cleanup": {"workbook_close": brief(&close), "excel_quit": brief(&quit), "owned_process_exit_verified": exited, "forced_termination": false},
        "success": exited,
    }))
}

/// Constructs `SAFEARRAY(VARIANT)` inputs through SDK APIs and assigns them
/// to matching Excel ranges. The complete range and every individual cell are
/// then read back before clearing the owned workbook.
pub(super) fn rectangular_writes(mode: Mode) -> Result<Value, String> {
    let _apartment = ComApartment::initialize()?;
    let app = activate(mode)?;
    let lcid = if matches!(mode, Mode::L) { 0x0400 } else { 0 };
    let owned = OwnedProcess::from_app(&app, lcid)?;
    let workbooks = get(&app, "Workbooks", lcid, vec![])?;
    let add = call(&workbooks, "Add", DISPATCH_METHOD, lcid, Frame::empty());
    let workbook = add
        .result
        .dispatch()
        .ok_or_else(|| format!("Workbooks.Add did not return VT_DISPATCH: {}", add.hr))?;
    let sheet = get(&app, "ActiveSheet", lcid, vec![])?;
    let mut observations = Vec::new();
    for (id, rows, columns, lower_bound) in [
        ("W-01", 1_u32, 1_u32, 1_i32),
        ("W-02", 1, 5, 1),
        ("W-03", 5, 1, 1),
        ("W-04", 2, 2, 1),
        ("W-05", 2, 3, 1),
        ("W-06", 3, 2, 1),
        ("W-07", 3, 4, 1),
        ("W-B-01", 2, 2, 0),
        ("W-B-02", 2, 2, 7),
    ] {
        observations.push(write_marker_array(
            &sheet,
            lcid,
            mode,
            id,
            rows,
            columns,
            lower_bound,
        )?);
    }

    let close = call(
        &workbook,
        "Close",
        DISPATCH_METHOD,
        lcid,
        Frame::positional(vec![OwnedVariant::boolean(false)]),
    );
    let quit = call(&app, "Quit", DISPATCH_METHOD, lcid, Frame::empty());
    drop(sheet);
    drop(workbook);
    drop(workbooks);
    drop(app);
    let exited = owned.wait();
    Ok(json!({
        "schema_version":1, "backend":"raw-windows-sys", "activation_mode":mode.id(),
        "workbooks_add":brief(&add), "observations":observations,
        "cleanup":{"workbook_close":brief(&close), "excel_quit":brief(&quit), "owned_process_exit_verified":exited, "forced_termination":false},
        "success":exited,
    }))
}

/// Combines the established rectangular read and SDK-array write groups into
/// one follow-up envelope.  Each inner group owns and cleans up its own Excel
/// process; the combined result makes the 2x3 read/write controls available to
/// the Prompt 05I S/X repeatability capture without modifying Prompt 05H rows.
pub(super) fn rectangular_differential(mode: Mode) -> Result<Value, String> {
    let reads = rectangular_reads(mode)?;
    let writes = rectangular_writes(mode)?;
    let mut observations = reads
        .get("observations")
        .and_then(Value::as_array)
        .cloned()
        .unwrap_or_default();
    observations.extend(
        writes
            .get("observations")
            .and_then(Value::as_array)
            .cloned()
            .unwrap_or_default(),
    );
    for observation in &mut observations {
        observation["category"] = Value::String("rectangular-differential".to_owned());
    }
    Ok(json!({
        "schema_version":1,
        "backend":"raw-windows-sys",
        "activation_mode":mode.id(),
        "observations":observations,
        "cleanup":{
            "rectangular_reads":reads.get("cleanup").cloned().unwrap_or(Value::Null),
            "rectangular_writes":writes.get("cleanup").cloned().unwrap_or(Value::Null),
            "owned_process_exit_verified":reads.get("success").and_then(Value::as_bool)==Some(true) && writes.get("success").and_then(Value::as_bool)==Some(true),
            "forced_termination":false,
        },
        "success":reads.get("success").and_then(Value::as_bool)==Some(true) && writes.get("success").and_then(Value::as_bool)==Some(true),
        "raw_pointer_values_recorded":false,
    }))
}

fn write_marker_array(
    sheet: &super::com_ptr::ComPtr<super::com_ptr::Dispatch>,
    lcid: u32,
    mode: Mode,
    id: &str,
    rows: u32,
    columns: u32,
    lower_bound: i32,
) -> Result<Value, String> {
    let address = range_address(1, 1, rows, columns);
    let target = get(sheet, "Range", lcid, vec![OwnedVariant::bstr(&address)?])?;
    let clear_before = call(&target, "ClearContents", DISPATCH_METHOD, lcid, Frame::empty());
    let bounds = [
        SAFEARRAYBOUND {
            cElements: rows,
            lLbound: lower_bound,
        },
        SAFEARRAYBOUND {
            cElements: columns,
            lLbound: lower_bound,
        },
    ];
    let array = OwnedSafeArray::create_variant(&bounds)?;
    for row in 0..rows {
        for column in 0..columns {
            let marker = i32::try_from((row + 1) * 1000 + column + 1)
                .map_err(|_| "marker conversion overflow")?;
            let value = OwnedVariant::i4(marker);
            let row_index = lower_bound
                .checked_add(i32::try_from(row).map_err(|_| "row index overflow")?)
                .ok_or("row index overflow")?;
            let column_index = lower_bound
                .checked_add(i32::try_from(column).map_err(|_| "column index overflow")?)
                .ok_or("column index overflow")?;
            array.put_variant(&[row_index, column_index], &value)?;
        }
    }
    let input_layout = unsafe { ObservedSafeArray::inspect(array.as_ptr()) }
        .ok_or_else(|| "cannot inspect newly constructed SAFEARRAY".to_owned())?;
    let write = call(
        &target,
        "Value2",
        DISPATCH_PROPERTYPUT,
        lcid,
        Frame::put(OwnedVariant::array(array)),
    );
    let value2 = call(&target, "Value2", DISPATCH_PROPERTYGET, lcid, Frame::empty());
    let value = call(&target, "Value", DISPATCH_PROPERTYGET, lcid, Frame::empty());
    let mut individual = Vec::new();
    for row in 0..rows {
        for column in 0..columns {
            let cell = cell_address(row + 1, column + 1);
            let cell = get(sheet, "Range", lcid, vec![OwnedVariant::bstr(&cell)?])?;
            let read = call(&cell, "Value2", DISPATCH_PROPERTYGET, lcid, Frame::empty());
            individual.push(json!({
                "logical_row":row,
                "logical_column":column,
                "value2_read_call":brief(&read),
                "value2_read":ObservedVariant::from_variant(&read.result),
            }));
        }
    }
    let clear_after = call(&target, "ClearContents", DISPATCH_METHOD, lcid, Frame::empty());
    Ok(json!({
        "schema_version":1, "id":id, "activation_mode":mode.id(), "fresh_process":true,
        "range_address":address, "rows":rows, "columns":columns, "content_pattern":"unique_row_column_markers",
        "input_safearray":input_layout, "clear_before":brief(&clear_before), "write":brief(&write),
        "value2_read_call":brief(&value2), "value2_read":rectangular_observation(&value2.result),
        "value_read_call":brief(&value), "value_read":rectangular_observation(&value.result),
        "individual_cells":individual, "clear_after":brief(&clear_after), "raw_pointer_values_recorded":false,
    }))
}

#[derive(Clone, Copy)]
enum CellValue {
    Empty,
    Bool(bool),
    I4(i32),
    Text(&'static str),
    Formula(&'static str),
}

fn read_pattern(
    sheet: &super::com_ptr::ComPtr<super::com_ptr::Dispatch>,
    lcid: u32,
    mode: Mode,
    id: &str,
    rows: u32,
    columns: u32,
    values: Vec<CellValue>,
) -> Result<Value, String> {
    let address = range_address(1, 1, rows, columns);
    let target = get(sheet, "Range", lcid, vec![OwnedVariant::bstr(&address)?])?;
    let clear_before = call(&target, "ClearContents", DISPATCH_METHOD, lcid, Frame::empty());
    for row in 0..rows {
        for column in 0..columns {
            let index = usize::try_from(row * columns + column).map_err(|_| "pattern index overflow")?;
            let cell = cell_address(row + 1, column + 1);
            let cell = get(sheet, "Range", lcid, vec![OwnedVariant::bstr(&cell)?])?;
            let write = match values[index] {
                CellValue::Empty => call(&cell, "Value2", DISPATCH_PROPERTYPUT, lcid, Frame::put(OwnedVariant::empty())),
                CellValue::Bool(value) => call(&cell, "Value2", DISPATCH_PROPERTYPUT, lcid, Frame::put(OwnedVariant::boolean(value))),
                CellValue::I4(value) => call(&cell, "Value2", DISPATCH_PROPERTYPUT, lcid, Frame::put(OwnedVariant::i4(value))),
                CellValue::Text(value) => call(&cell, "Value2", DISPATCH_PROPERTYPUT, lcid, Frame::put(OwnedVariant::bstr(value)?)),
                CellValue::Formula(value) => call(&cell, "Formula", DISPATCH_PROPERTYPUT, lcid, Frame::put(OwnedVariant::bstr(value)?)),
            };
            if write.hr != 0 {
                return Err(format!("{id} write failed: {}", write.hr));
            }
        }
    }
    let value2 = call(&target, "Value2", DISPATCH_PROPERTYGET, lcid, Frame::empty());
    let value = call(&target, "Value", DISPATCH_PROPERTYGET, lcid, Frame::empty());
    let clear_after = call(&target, "ClearContents", DISPATCH_METHOD, lcid, Frame::empty());
    Ok(json!({
        "schema_version": 1, "id": id, "activation_mode":mode.id(), "fresh_process":true,
        "range_address":address, "rows":rows, "columns":columns, "content_pattern":id,
        "clear_before":brief(&clear_before), "value2_read_call":brief(&value2), "value2_read":rectangular_observation(&value2.result),
        "value_read_call":brief(&value), "value_read":rectangular_observation(&value.result), "clear_after":brief(&clear_after),
        "raw_pointer_values_recorded":false,
    }))
}

fn rectangular_observation(value: &OwnedVariant) -> Value {
    let owner = ObservedVariant::from_variant(value);
    if value.vt() & VT_ARRAY == 0 {
        return json!({"owner":owner, "shape":"scalar"});
    }
    let array = unsafe { value.0.Anonymous.Anonymous.Anonymous.parray };
    let Some(layout) = (unsafe { ObservedSafeArray::inspect(array) }) else {
        return json!({"owner":owner, "layout_error":"SDK metadata inspection failed"});
    };
    let mut cells = Vec::new();
    if layout.rank == 2 {
        let first = &layout.dimensions[0];
        let second = &layout.dimensions[1];
        for row in 0..first.element_count {
            for column in 0..second.element_count {
                let Some(indices) = row_column_indices(&layout.dimensions, row, column) else {
                    continue;
                };
                let mut element = OwnedVariant::empty();
                let result = unsafe {
                    SafeArrayGetElement(
                        array,
                        indices.as_ptr(),
                        &mut element.0 as *mut VARIANT as *mut c_void,
                    )
                };
                cells.push(json!({
                    "physical_indices": indices,
                    "logical_row":row,
                    "logical_column":column,
                    "get_element_hresult":format!("0x{:08X}", result as u32),
                    "value": if result == 0 { serde_json::to_value(ObservedVariant::from_variant(&element)).unwrap_or(Value::Null) } else { Value::Null },
                }));
            }
        }
    }
    json!({
        "owner":owner,
        "checked_element_count":checked_element_count(&layout.dimensions),
        "layout":layout,
        "physical_traversal":"increasing physical dimensions [1, 2], queried via SafeArrayGetElement",
        "logical_mapping":"row markers prove physical dimension 1 maps to Excel rows and physical dimension 2 maps to Excel columns when values match row*1000+column",
        "cells":cells,
    })
}

fn cell_address(row: u32, column: u32) -> String {
    let mut column = column;
    let mut name = String::new();
    while column > 0 {
        let remainder = (column - 1) % 26;
        name.insert(0, char::from_u32(u32::from(b'A') + remainder).unwrap_or('?'));
        column = (column - 1) / 26;
    }
    format!("{name}{row}")
}

fn range_address(row: u32, column: u32, rows: u32, columns: u32) -> String {
    let first = cell_address(row, column);
    let last = cell_address(row + rows - 1, column + columns - 1);
    format!("{first}:{last}")
}

/// Captures the scalar semantics that cannot be inferred from the scalar
/// write sweep alone: blank/cleared/formula-empty cells, date/currency format
/// effects, formula-produced Excel errors, and Formula/Formula2 results.
pub(super) fn miscellaneous_semantics(mode: Mode) -> Result<Value, String> {
    let _apartment = ComApartment::initialize()?;
    let app = activate(mode)?;
    let lcid = if matches!(mode, Mode::L) { 0x0400 } else { 0 };
    let owned = OwnedProcess::from_app(&app, lcid)?;
    let workbooks = get(&app, "Workbooks", lcid, vec![])?;
    let add = call(&workbooks, "Add", DISPATCH_METHOD, lcid, Frame::empty());
    let workbook = add
        .result
        .dispatch()
        .ok_or_else(|| format!("Workbooks.Add did not return VT_DISPATCH: {}", add.hr))?;
    let sheet = get(&app, "ActiveSheet", lcid, vec![])?;
    let mut observations = Vec::new();

    for (id, setup) in [
        ("B-01", BlankSetup::NeverWritten),
        ("B-02", BlankSetup::Value2(OwnedVariant::empty())),
        ("B-03", BlankSetup::Value2(OwnedVariant::null())),
        ("B-04", BlankSetup::Value2(OwnedVariant::bstr("")?)),
        ("B-05", BlankSetup::Cleared),
        ("B-06", BlankSetup::Formula("=\"\"")),
    ] {
        let cell = get(&sheet, "Range", lcid, vec![OwnedVariant::bstr("A1")?])?;
        let clear_before = call(&cell, "ClearContents", DISPATCH_METHOD, lcid, Frame::empty());
        let setup_name = blank_setup_name(&setup);
        let write = match setup {
            BlankSetup::NeverWritten => None,
            BlankSetup::Value2(value) => Some(call(
                &cell,
                "Value2",
                DISPATCH_PROPERTYPUT,
                lcid,
                Frame::put(value),
            )),
            BlankSetup::Cleared => {
                let initial = call(
                    &cell,
                    "Value2",
                    DISPATCH_PROPERTYPUT,
                    lcid,
                    Frame::put(OwnedVariant::i4(1)),
                );
                if initial.hr != 0 {
                    return Err(format!("{id} initial write failed"));
                }
                Some(call(&cell, "ClearContents", DISPATCH_METHOD, lcid, Frame::empty()))
            }
            BlankSetup::Formula(text) => Some(call(
                &cell,
                "Formula",
                DISPATCH_PROPERTYPUT,
                lcid,
                Frame::put(OwnedVariant::bstr(text)?),
            )),
        };
        observations.push(json!({
            "schema_version":1, "id":id, "category":"blank-null-empty", "activation_mode":mode.id(), "fresh_process":true,
            "setup":setup_name, "clear_before":brief(&clear_before), "write_or_setup":write.as_ref().map(brief),
            "reads":read_members(&cell, lcid), "clear_after":brief(&call(&cell, "ClearContents", DISPATCH_METHOD, lcid, Frame::empty())),
            "raw_pointer_values_recorded":false,
        }));
    }

    for (id, serial) in [
        ("D-01", -1.0), ("D-02", 0.0), ("D-03", 1.0), ("D-04", 2.0),
        ("D-05", 0.5), ("D-06", 45_292.0), ("D-07", 45_292.5),
    ] {
        observations.push(date_or_currency_case(
            &sheet, lcid, mode, id, "date-currency", "Value", OwnedVariant::date(serial), "m/d/yyyy h:mm",
        )?);
    }
    for (id, scaled) in [
        ("C-01", 0_i64), ("C-02", 1_234_500), ("C-03", -1_234_500), ("C-04", 9_876_543_210_123),
    ] {
        observations.push(date_or_currency_case(
            &sheet, lcid, mode, id, "date-currency", "Value", OwnedVariant::currency(scaled), "$#,##0.0000",
        )?);
    }

    for (id, formula, expected) in [
        ("E-01", "=A1 A2", "#NULL!"),
        ("E-02", "=1/0", "#DIV/0!"),
        ("E-03", "=1+\"x\"", "#VALUE!"),
        ("E-04", "=INDIRECT(\"A0\")", "#REF!"),
        ("E-05", "=NOT_A_FUNCTION()", "#NAME?"),
        ("E-06", "=SQRT(-1)", "#NUM!"),
        ("E-07", "=NA()", "#N/A"),
    ] {
        let cell = get(&sheet, "Range", lcid, vec![OwnedVariant::bstr("A1")?])?;
        let clear_before = call(&cell, "ClearContents", DISPATCH_METHOD, lcid, Frame::empty());
        let write = call(&cell, "Formula", DISPATCH_PROPERTYPUT, lcid, Frame::put(OwnedVariant::bstr(formula)?));
        observations.push(json!({
            "schema_version":1,"id":id,"category":"errors","activation_mode":mode.id(),"fresh_process":true,
            "expected_display_error":expected,"formula_input":formula,"clear_before":brief(&clear_before),"write":brief(&write),"reads":read_members(&cell,lcid),
            "clear_after":brief(&call(&cell,"ClearContents",DISPATCH_METHOD,lcid,Frame::empty())),"raw_pointer_values_recorded":false,
        }));
    }

    let reference = get(&sheet, "Range", lcid, vec![OwnedVariant::bstr("B1")?])?;
    let _ = call(&reference, "Value2", DISPATCH_PROPERTYPUT, lcid, Frame::put(OwnedVariant::i4(40)));
    for (id, formula) in [
        ("F-01", "=1+2"), ("F-02", "=B1+1"), ("F-03", "=\"text\""), ("F-04", "=1=1"),
        ("F-05", "=1/0"), ("F-06", "=DATE(2024,1,2)"), ("F-07", "=\"\""), ("F-08", "=\"Grüße Ω\""),
    ] {
        let cell = get(&sheet, "Range", lcid, vec![OwnedVariant::bstr("A1")?])?;
        let clear_before = call(&cell, "ClearContents", DISPATCH_METHOD, lcid, Frame::empty());
        let write = call(&cell, "Formula", DISPATCH_PROPERTYPUT, lcid, Frame::put(OwnedVariant::bstr(formula)?));
        observations.push(json!({
            "schema_version":1,"id":id,"category":"formulas","activation_mode":mode.id(),"fresh_process":true,
            "formula_property":"Formula","formula_input":formula,"clear_before":brief(&clear_before),"write":brief(&write),"reads":read_members(&cell,lcid),
            "clear_after":brief(&call(&cell,"ClearContents",DISPATCH_METHOD,lcid,Frame::empty())),"raw_pointer_values_recorded":false,
        }));
    }

    let close = call(&workbook, "Close", DISPATCH_METHOD, lcid, Frame::positional(vec![OwnedVariant::boolean(false)]));
    let quit = call(&app, "Quit", DISPATCH_METHOD, lcid, Frame::empty());
    drop(reference);
    drop(sheet);
    drop(workbook);
    drop(workbooks);
    drop(app);
    let exited = owned.wait();
    Ok(json!({
        "schema_version":1,"backend":"raw-windows-sys","activation_mode":mode.id(),"workbooks_add":brief(&add),"observations":observations,
        "cleanup":{"workbook_close":brief(&close),"excel_quit":brief(&quit),"owned_process_exit_verified":exited,"forced_termination":false},"success":exited,
    }))
}

enum BlankSetup {
    NeverWritten,
    Value2(OwnedVariant),
    Cleared,
    Formula(&'static str),
}

fn blank_setup_name(setup: &BlankSetup) -> &'static str {
    match setup {
        BlankSetup::NeverWritten => "never_written_blank",
        BlankSetup::Value2(value) if value.vt() == 0 => "written_VT_EMPTY",
        BlankSetup::Value2(value) if value.vt() == 1 => "written_VT_NULL",
        BlankSetup::Value2(_) => "written_empty_string",
        BlankSetup::Cleared => "cleared_with_ClearContents",
        BlankSetup::Formula(_) => "formula_returning_empty_string",
    }
}

fn read_members(
    cell: &super::com_ptr::ComPtr<super::com_ptr::Dispatch>,
    lcid: u32,
) -> Value {
    let value = call(cell, "Value", DISPATCH_PROPERTYGET, lcid, Frame::empty());
    let value2 = call(cell, "Value2", DISPATCH_PROPERTYGET, lcid, Frame::empty());
    let formula = call(cell, "Formula", DISPATCH_PROPERTYGET, lcid, Frame::empty());
    let formula2 = call(cell, "Formula2", DISPATCH_PROPERTYGET, lcid, Frame::empty());
    json!({
        "Value":{"call":brief(&value),"value":ObservedVariant::from_variant(&value.result)},
        "Value2":{"call":brief(&value2),"value":ObservedVariant::from_variant(&value2.result)},
        "Formula":{"call":brief(&formula),"value":ObservedVariant::from_variant(&formula.result)},
        "Formula2":{"call":brief(&formula2),"value":ObservedVariant::from_variant(&formula2.result)},
    })
}

fn date_or_currency_case(
    sheet: &super::com_ptr::ComPtr<super::com_ptr::Dispatch>,
    lcid: u32,
    mode: Mode,
    id: &str,
    category: &str,
    member: &str,
    input: OwnedVariant,
    number_format: &str,
) -> Result<Value, String> {
    let cell = get(sheet, "Range", lcid, vec![OwnedVariant::bstr("A1")?])?;
    let clear_before = call(&cell, "ClearContents", DISPATCH_METHOD, lcid, Frame::empty());
    let format_write = call(&cell, "NumberFormat", DISPATCH_PROPERTYPUT, lcid, Frame::put(OwnedVariant::bstr(number_format)?));
    let input_vartype = input.vt();
    let write = call(&cell, member, DISPATCH_PROPERTYPUT, lcid, Frame::put(input));
    let format_read = call(&cell, "NumberFormat", DISPATCH_PROPERTYGET, lcid, Frame::empty());
    Ok(json!({
        "schema_version":1,"id":id,"category":category,"activation_mode":mode.id(),"fresh_process":true,
        "input_member":member,"input_vartype":input_vartype,"number_format_input":number_format,"clear_before":brief(&clear_before),
        "number_format_write":brief(&format_write),"write":brief(&write),"number_format_read":{"call":brief(&format_read),"value":ObservedVariant::from_variant(&format_read.result)},
        "reads":read_members(&cell,lcid),"clear_after":brief(&call(&cell,"ClearContents",DISPATCH_METHOD,lcid,Frame::empty())),"raw_pointer_values_recorded":false,
    }))
}

/// Covers precision boundaries, BSTR edge cases, a mixed `SAFEARRAY(VARIANT)`
/// write, and the installed Excel's Formula2/dynamic-array capability.
pub(super) fn edge_cases(mode: Mode) -> Result<Value, String> {
    let _apartment = ComApartment::initialize()?;
    let app = activate(mode)?;
    let lcid = if matches!(mode, Mode::L) { 0x0400 } else { 0 };
    let owned = OwnedProcess::from_app(&app, lcid)?;
    let workbooks = get(&app, "Workbooks", lcid, vec![])?;
    let add = call(&workbooks, "Add", DISPATCH_METHOD, lcid, Frame::empty());
    let workbook = add
        .result
        .dispatch()
        .ok_or_else(|| format!("Workbooks.Add did not return VT_DISPATCH: {}", add.hr))?;
    let sheet = get(&app, "ActiveSheet", lcid, vec![])?;
    let mut observations = Vec::new();

    for (id, input, semantic) in [
        ("P-01", OwnedVariant::i8(9_007_199_254_991), "VT_I8(2^53-1)"),
        ("P-02", OwnedVariant::i8(9_007_199_254_992), "VT_I8(2^53)"),
        ("P-03", OwnedVariant::i8(9_007_199_254_993), "VT_I8(2^53+1)"),
        ("P-04", OwnedVariant::i4(i32::MAX), "VT_I4(i32::MAX)"),
        ("P-05", OwnedVariant::i4(i32::MIN), "VT_I4(i32::MIN)"),
        ("P-06", OwnedVariant::r8(123_456_789_012_345.67), "VT_R8(many significant digits)"),
        ("P-07", OwnedVariant::r8(f64::MIN_POSITIVE), "VT_R8(min positive normal)"),
        ("P-08", OwnedVariant::r8(f64::from_bits(1)), "VT_R8(min subnormal)"),
        ("P-09", OwnedVariant::r8(1.0e307), "VT_R8(large finite)"),
        ("P-10", OwnedVariant::r8(-0.0), "VT_R8(negative zero)"),
    ] {
        let input_vartype = input.vt();
        let input_bits = observed_numeric_bits(&input);
        let cell = get(&sheet, "Range", lcid, vec![OwnedVariant::bstr("A1")?])?;
        let clear_before = call(&cell, "ClearContents", DISPATCH_METHOD, lcid, Frame::empty());
        let write = call(&cell, "Value2", DISPATCH_PROPERTYPUT, lcid, Frame::put(input));
        let value2 = call(&cell, "Value2", DISPATCH_PROPERTYGET, lcid, Frame::empty());
        observations.push(json!({
            "schema_version":1,"id":id,"category":"precision","activation_mode":mode.id(),"fresh_process":true,
            "input_semantic_value":semantic,"input_vartype":input_vartype,"input_numeric_bits":input_bits,
            "clear_before":brief(&clear_before),"write":brief(&write),"value2_read_call":brief(&value2),"value2_read":ObservedVariant::from_variant(&value2.result),
            "value_read":read_members(&cell,lcid).get("Value").cloned().unwrap_or(Value::Null),"clear_after":brief(&call(&cell,"ClearContents",DISPATCH_METHOD,lcid,Frame::empty())),"raw_pointer_values_recorded":false,
        }));
    }

    let mut strings = vec![
        ("T-01", String::new(), "empty"),
        ("T-02", "  leading".to_owned(), "leading spaces"),
        ("T-03", "trailing  ".to_owned(), "trailing spaces"),
        ("T-04", "line\nbreak".to_owned(), "line break"),
        ("T-05", "tab\tvalue".to_owned(), "tab"),
        ("T-06", "left\0right".to_owned(), "embedded NUL"),
        ("T-07", "Grüße Ω".to_owned(), "BMP Unicode"),
        ("T-08", "supplementary 😀".to_owned(), "supplementary Unicode"),
        ("T-09", "e\u{301}".to_owned(), "combining characters"),
        ("T-10", "=1+2".to_owned(), "formula-like text"),
        ("T-11", "'=1+2".to_owned(), "apostrophe-prefixed text"),
        ("T-12", "x".repeat(32_766), "below Excel cell limit"),
        ("T-13", "x".repeat(32_767), "at Excel cell limit"),
        ("T-14", "x".repeat(32_768), "beyond Excel cell limit"),
    ];
    for (id, input, semantic) in strings.drain(..) {
        let input_length = input.encode_utf16().count();
        let cell = get(&sheet, "Range", lcid, vec![OwnedVariant::bstr("A1")?])?;
        let clear_before = call(&cell, "ClearContents", DISPATCH_METHOD, lcid, Frame::empty());
        let write = call(&cell, "Value2", DISPATCH_PROPERTYPUT, lcid, Frame::put(OwnedVariant::bstr(&input)?));
        let value2 = call(&cell, "Value2", DISPATCH_PROPERTYGET, lcid, Frame::empty());
        let formula = call(&cell, "Formula", DISPATCH_PROPERTYGET, lcid, Frame::empty());
        observations.push(json!({
            "schema_version":1,"id":id,"category":"strings","activation_mode":mode.id(),"fresh_process":true,
            "semantic_input":semantic,"input_utf16_length":input_length,"clear_before":brief(&clear_before),"write":brief(&write),
            "value2_read":{"call":brief(&value2),"summary":string_summary(&value2.result)},"formula_read":{"call":brief(&formula),"summary":string_summary(&formula.result)},
            "clear_after":brief(&call(&cell,"ClearContents",DISPATCH_METHOD,lcid,Frame::empty())),"raw_pointer_values_recorded":false,
        }));
    }

    observations.push(mixed_array_case(&sheet, lcid, mode)?);
    observations.push(dynamic_array_case(&sheet, lcid, mode)?);

    let close = call(&workbook, "Close", DISPATCH_METHOD, lcid, Frame::positional(vec![OwnedVariant::boolean(false)]));
    let quit = call(&app, "Quit", DISPATCH_METHOD, lcid, Frame::empty());
    drop(sheet);
    drop(workbook);
    drop(workbooks);
    drop(app);
    let exited = owned.wait();
    Ok(json!({
        "schema_version":1,"backend":"raw-windows-sys","activation_mode":mode.id(),"workbooks_add":brief(&add),"observations":observations,
        "cleanup":{"workbook_close":brief(&close),"excel_quit":brief(&quit),"owned_process_exit_verified":exited,"forced_termination":false},"success":exited,
    }))
}

fn observed_numeric_bits(value: &OwnedVariant) -> Value {
    match ObservedVariant::from_variant(value) {
        ObservedVariant::F64 { value_bits, .. } => json!(value_bits),
        ObservedVariant::F32 { value_bits, .. } => json!(value_bits),
        ObservedVariant::I32 { value, .. } => json!(value),
        ObservedVariant::I64 { value, .. } => json!(value),
        _ => Value::Null,
    }
}

fn string_summary(value: &OwnedVariant) -> Value {
    match ObservedVariant::from_variant(value) {
        ObservedVariant::String { vartype, utf16 } => json!({
            "kind":"string", "vartype":vartype, "utf16_length":utf16.len(),
            "utf16_preview": if utf16.len() <= 80 { json!(utf16) } else { Value::Null },
        }),
        other => serde_json::to_value(other).unwrap_or(Value::Null),
    }
}

fn mixed_array_case(
    sheet: &super::com_ptr::ComPtr<super::com_ptr::Dispatch>,
    lcid: u32,
    mode: Mode,
) -> Result<Value, String> {
    let address = "A1:C3";
    let target = get(sheet, "Range", lcid, vec![OwnedVariant::bstr(address)?])?;
    let clear_before = call(&target, "ClearContents", DISPATCH_METHOD, lcid, Frame::empty());
    let bounds = [
        SAFEARRAYBOUND { cElements: 3, lLbound: 1 },
        SAFEARRAYBOUND { cElements: 3, lLbound: 1 },
    ];
    let array = OwnedSafeArray::create_variant(&bounds)?;
    let values = vec![
        OwnedVariant::empty(), OwnedVariant::null(), OwnedVariant::boolean(true),
        OwnedVariant::i4(42), OwnedVariant::r8(1.25), OwnedVariant::bstr("text")?,
        OwnedVariant::error(2042), OwnedVariant::date(45_292.5), OwnedVariant::currency(1_234_500),
    ];
    for (index, value) in values.iter().enumerate() {
        let row = i32::try_from(index / 3 + 1).map_err(|_| "mixed row index overflow")?;
        let column = i32::try_from(index % 3 + 1).map_err(|_| "mixed column index overflow")?;
        array.put_variant(&[row, column], value)?;
    }
    let input_layout = unsafe { ObservedSafeArray::inspect(array.as_ptr()) }
        .ok_or_else(|| "cannot inspect mixed SAFEARRAY".to_owned())?;
    let write = call(&target, "Value2", DISPATCH_PROPERTYPUT, lcid, Frame::put(OwnedVariant::array(array)));
    let value2 = call(&target, "Value2", DISPATCH_PROPERTYGET, lcid, Frame::empty());
    let value = call(&target, "Value", DISPATCH_PROPERTYGET, lcid, Frame::empty());
    Ok(json!({
        "schema_version":1,"id":"M-01","category":"mixed-array","activation_mode":mode.id(),"fresh_process":true,
        "range_address":address,"input_safearray":input_layout,"write":brief(&write),"value2_read":rectangular_observation(&value2.result),"value2_read_call":brief(&value2),"value_read":rectangular_observation(&value.result),"value_read_call":brief(&value),
        "clear_before":brief(&clear_before),"clear_after":brief(&call(&target,"ClearContents",DISPATCH_METHOD,lcid,Frame::empty())),"raw_pointer_values_recorded":false,
    }))
}

#[derive(Clone, Copy)]
enum MixedReplacement {
    R8,
    Empty,
    Null,
    I4,
    Error,
    Date,
    Currency,
}

impl MixedReplacement {
    fn semantic(self) -> &'static str {
        match self {
            Self::R8 => "VT_R8(9.5) control",
            Self::Empty => "VT_EMPTY",
            Self::Null => "VT_NULL",
            Self::I4 => "VT_I4(42)",
            Self::Error => "VT_ERROR(2042)",
            Self::Date => "VT_DATE(45292.5)",
            Self::Currency => "VT_CY(123.4500)",
        }
    }

    fn value(self) -> OwnedVariant {
        match self {
            Self::R8 => OwnedVariant::r8(9.5),
            Self::Empty => OwnedVariant::empty(),
            Self::Null => OwnedVariant::null(),
            Self::I4 => OwnedVariant::i4(42),
            Self::Error => OwnedVariant::error(2042),
            Self::Date => OwnedVariant::date(45_292.5),
            Self::Currency => OwnedVariant::currency(1_234_500),
        }
    }
}

/// Narrows Prompt 05H's mixed array failure to one fixed position at a time.
/// The payload is always a 3x3 `SAFEARRAY(VARIANT)` with the same R8/BSTR/BOOL
/// control cells.  This is deliberately raw evidence: every read retains the
/// physical owner and element VARTYPEs through the SDK inspector.
pub(super) fn mixed_array_differential(mode: Mode) -> Result<Value, String> {
    let _apartment = ComApartment::initialize()?;
    let app = activate(mode)?;
    let lcid = if matches!(mode, Mode::L) { 0x0400 } else { 0 };
    let owned = OwnedProcess::from_app(&app, lcid)?;
    let workbooks = get(&app, "Workbooks", lcid, vec![])?;
    let add = call(&workbooks, "Add", DISPATCH_METHOD, lcid, Frame::empty());
    let workbook = add
        .result
        .dispatch()
        .ok_or_else(|| format!("Workbooks.Add did not return VT_DISPATCH: {}", add.hr))?;
    let sheet = get(&app, "ActiveSheet", lcid, vec![])?;
    let mut observations = Vec::new();
    let mut first_failure = None;
    for (id, replacement) in [
        ("M-D-00", MixedReplacement::R8),
        ("M-D-01", MixedReplacement::Empty),
        ("M-D-02", MixedReplacement::Null),
        ("M-D-03", MixedReplacement::I4),
        ("M-D-04", MixedReplacement::Error),
        ("M-D-05", MixedReplacement::Date),
        ("M-D-06", MixedReplacement::Currency),
    ] {
        let row = mixed_differential_case(&sheet, lcid, mode, id, replacement)?;
        if first_failure.is_none()
            && row.pointer("/write/hresult").and_then(Value::as_i64) != Some(0)
        {
            first_failure = Some((id, replacement));
        }
        observations.push(row);
    }
    if let Some((failed_id, replacement)) = first_failure {
        for retry in 1..=3 {
            let mut row = mixed_differential_case(&sheet, lcid, mode, &format!("{failed_id}-rerun-{retry}"), replacement)?;
            row["reproduces"] = Value::String(failed_id.to_owned());
            row["rerun"] = Value::from(retry);
            observations.push(row);
        }
    }
    let close = call(&workbook, "Close", DISPATCH_METHOD, lcid, Frame::positional(vec![OwnedVariant::boolean(false)]));
    let quit = call(&app, "Quit", DISPATCH_METHOD, lcid, Frame::empty());
    drop(sheet);
    drop(workbook);
    drop(workbooks);
    drop(app);
    let exited = owned.wait();
    Ok(json!({
        "schema_version":1,"backend":"raw-windows-sys","activation_mode":mode.id(),"workbooks_add":brief(&add),"observations":observations,
        "cleanup":{"workbook_close":brief(&close),"excel_quit":brief(&quit),"owned_process_exit_verified":exited,"forced_termination":false},"success":exited,
    }))
}

fn mixed_differential_case(
    sheet: &super::com_ptr::ComPtr<super::com_ptr::Dispatch>,
    lcid: u32,
    mode: Mode,
    id: &str,
    replacement: MixedReplacement,
) -> Result<Value, String> {
    let target = get(sheet, "Range", lcid, vec![OwnedVariant::bstr("A1:C3")?])?;
    let clear_before = call(&target, "ClearContents", DISPATCH_METHOD, lcid, Frame::empty());
    let bounds = [
        SAFEARRAYBOUND { cElements: 3, lLbound: 1 },
        SAFEARRAYBOUND { cElements: 3, lLbound: 1 },
    ];
    let array = OwnedSafeArray::create_variant(&bounds)?;
    let values = vec![
        OwnedVariant::r8(1.25),
        OwnedVariant::bstr("text")?,
        OwnedVariant::boolean(true),
        OwnedVariant::r8(2.5),
        OwnedVariant::bstr("more")?,
        OwnedVariant::boolean(false),
        OwnedVariant::r8(3.5),
        replacement.value(),
        OwnedVariant::boolean(true),
    ];
    for (index, value) in values.iter().enumerate() {
        let row = i32::try_from(index / 3 + 1).map_err(|_| "mixed differential row overflow")?;
        let column = i32::try_from(index % 3 + 1).map_err(|_| "mixed differential column overflow")?;
        array.put_variant(&[row, column], value)?;
    }
    let input_safearray = unsafe { ObservedSafeArray::inspect(array.as_ptr()) }
        .ok_or_else(|| "cannot inspect differential mixed SAFEARRAY".to_owned())?;
    let write = call(&target, "Value2", DISPATCH_PROPERTYPUT, lcid, Frame::put(OwnedVariant::array(array)));
    let value2 = call(&target, "Value2", DISPATCH_PROPERTYGET, lcid, Frame::empty());
    let value = call(&target, "Value", DISPATCH_PROPERTYGET, lcid, Frame::empty());
    Ok(json!({
        "schema_version":1,"id":id,"category":"mixed-array-differential","activation_mode":mode.id(),"fresh_process":true,
        "range_address":"A1:C3","fixed_position":{"logical_row":2,"logical_column":1,"physical_indices":[3,2]},
        "control_elements":"R8/BSTR/BOOL in all non-target positions","replacement_semantic":replacement.semantic(),"input_safearray":input_safearray,
        "clear_before":brief(&clear_before),"write":brief(&write),"value2_read_call":brief(&value2),"value2_read":rectangular_observation(&value2.result),
        "value_read_call":brief(&value),"value_read":rectangular_observation(&value.result),"clear_after":brief(&call(&target,"ClearContents",DISPATCH_METHOD,lcid,Frame::empty())),"raw_pointer_values_recorded":false,
    }))
}

/// Separates `VT_DATE` writes from equal OA doubles through both members and
/// General/date number formats.  Negative serials remain observations, not a
/// policy decision about Excel's calendar boundary.
pub(super) fn date_differential(mode: Mode) -> Result<Value, String> {
    let _apartment = ComApartment::initialize()?;
    let app = activate(mode)?;
    let lcid = if matches!(mode, Mode::L) { 0x0400 } else { 0 };
    let owned = OwnedProcess::from_app(&app, lcid)?;
    let workbooks = get(&app, "Workbooks", lcid, vec![])?;
    let add = call(&workbooks, "Add", DISPATCH_METHOD, lcid, Frame::empty());
    let workbook = add.result.dispatch().ok_or_else(|| format!("Workbooks.Add did not return VT_DISPATCH: {}", add.hr))?;
    let sheet = get(&app, "ActiveSheet", lcid, vec![])?;
    let mut observations = Vec::new();
    for (serial_id, serial) in [("minus-one", -1.0), ("zero", 0.0), ("one", 1.0), ("half", 0.5), ("modern", 45_292.0), ("negative-half", -0.5)] {
        for (format_id, number_format) in [("general", "General"), ("date", "m/d/yyyy h:mm")] {
            observations.push(date_or_currency_case(&sheet, lcid, mode, &format!("D-DATE-{serial_id}-{format_id}"), "date-differential", "Value", OwnedVariant::date(serial), number_format)?);
            observations.push(date_or_currency_case(&sheet, lcid, mode, &format!("D-R8-{serial_id}-{format_id}"), "date-differential", "Value2", OwnedVariant::r8(serial), number_format)?);
        }
    }
    let close = call(&workbook, "Close", DISPATCH_METHOD, lcid, Frame::positional(vec![OwnedVariant::boolean(false)]));
    let quit = call(&app, "Quit", DISPATCH_METHOD, lcid, Frame::empty());
    drop(sheet);
    drop(workbook);
    drop(workbooks);
    drop(app);
    let exited = owned.wait();
    Ok(json!({"schema_version":1,"backend":"raw-windows-sys","activation_mode":mode.id(),"workbooks_add":brief(&add),"observations":observations,"cleanup":{"workbook_close":brief(&close),"excel_quit":brief(&quit),"owned_process_exit_verified":exited,"forced_termination":false},"success":exited}))
}

/// Covers the target/input shape mismatch and rank cases without making a
/// future SAFEARRAY coercion policy.  Inputs are constructed only with SDK
/// SAFEARRAY APIs and observed again through raw `Value2` reads.
pub(super) fn shape_differential(mode: Mode) -> Result<Value, String> {
    let _apartment = ComApartment::initialize()?;
    let app = activate(mode)?;
    let lcid = if matches!(mode, Mode::L) { 0x0400 } else { 0 };
    let owned = OwnedProcess::from_app(&app, lcid)?;
    let workbooks = get(&app, "Workbooks", lcid, vec![])?;
    let add = call(&workbooks, "Add", DISPATCH_METHOD, lcid, Frame::empty());
    let workbook = add.result.dispatch().ok_or_else(|| format!("Workbooks.Add did not return VT_DISPATCH: {}", add.hr))?;
    let sheet = get(&app, "ActiveSheet", lcid, vec![])?;
    let mut observations = Vec::new();
    for (id, address, dimensions) in [
        ("SH-01-1x2-to-1x3", "A1:C1", vec![1_u32, 2]),
        ("SH-02-1x3-to-1x2", "A1:B1", vec![1, 3]),
        ("SH-03-2x2-to-2x3", "A1:C2", vec![2, 2]),
        ("SH-04-2x3-to-2x2", "A1:B2", vec![2, 3]),
        ("SH-05-rank1-row", "A1:C1", vec![3]),
        ("SH-06-rank1-column", "A1:A3", vec![3]),
        ("SH-07-rank3", "A1", vec![1, 1, 1]),
    ] {
        observations.push(shape_differential_case(&sheet, lcid, mode, id, address, &dimensions)?);
    }
    let close = call(&workbook, "Close", DISPATCH_METHOD, lcid, Frame::positional(vec![OwnedVariant::boolean(false)]));
    let quit = call(&app, "Quit", DISPATCH_METHOD, lcid, Frame::empty());
    drop(sheet);
    drop(workbook);
    drop(workbooks);
    drop(app);
    let exited = owned.wait();
    Ok(json!({"schema_version":1,"backend":"raw-windows-sys","activation_mode":mode.id(),"workbooks_add":brief(&add),"observations":observations,"cleanup":{"workbook_close":brief(&close),"excel_quit":brief(&quit),"owned_process_exit_verified":exited,"forced_termination":false},"success":exited}))
}

fn shape_differential_case(
    sheet: &super::com_ptr::ComPtr<super::com_ptr::Dispatch>,
    lcid: u32,
    mode: Mode,
    id: &str,
    address: &str,
    dimensions: &[u32],
) -> Result<Value, String> {
    let target = get(sheet, "Range", lcid, vec![OwnedVariant::bstr(address)?])?;
    let clear_before = call(&target, "ClearContents", DISPATCH_METHOD, lcid, Frame::empty());
    let bounds = dimensions.iter().map(|elements| SAFEARRAYBOUND { cElements: *elements, lLbound: 1 }).collect::<Vec<_>>();
    let array = OwnedSafeArray::create_variant(&bounds)?;
    let mut marker = 1_i32;
    match dimensions {
        [one] => for first in 1..=*one { let value = OwnedVariant::i4(marker); marker += 1; array.put_variant(&[i32::try_from(first).map_err(|_| "rank1 index overflow")?], &value)?; },
        [rows, columns] => for row in 1..=*rows { for column in 1..=*columns { let value = OwnedVariant::i4(marker); marker += 1; array.put_variant(&[i32::try_from(row).map_err(|_| "rank2 row index overflow")?, i32::try_from(column).map_err(|_| "rank2 column index overflow")?], &value)?; } },
        [first, second, third] => for one in 1..=*first { for two in 1..=*second { for three in 1..=*third { let value = OwnedVariant::i4(marker); marker += 1; array.put_variant(&[i32::try_from(one).map_err(|_| "rank3 first index overflow")?, i32::try_from(two).map_err(|_| "rank3 second index overflow")?, i32::try_from(three).map_err(|_| "rank3 third index overflow")?], &value)?; } } },
        _ => return Err("unsupported differential SAFEARRAY rank".to_owned()),
    }
    let input_safearray = unsafe { ObservedSafeArray::inspect(array.as_ptr()) }.ok_or_else(|| "cannot inspect shape differential input".to_owned())?;
    let write = call(&target, "Value2", DISPATCH_PROPERTYPUT, lcid, Frame::put(OwnedVariant::array(array)));
    let read = call(&target, "Value2", DISPATCH_PROPERTYGET, lcid, Frame::empty());
    Ok(json!({"schema_version":1,"id":id,"category":"shape-mismatch","activation_mode":mode.id(),"fresh_process":true,"target_range":address,"input_safearray":input_safearray,"clear_before":brief(&clear_before),"write":brief(&write),"value2_read_call":brief(&read),"value2_read":rectangular_observation(&read.result),"clear_after":brief(&call(&target,"ClearContents",DISPATCH_METHOD,lcid,Frame::empty())),"raw_pointer_values_recorded":false}))
}

/// Exercises Formula2's normal spill, text spill, and blocked spill forms.
pub(super) fn dynamic_array_differential(mode: Mode) -> Result<Value, String> {
    let _apartment = ComApartment::initialize()?;
    let app = activate(mode)?;
    let lcid = if matches!(mode, Mode::L) { 0x0400 } else { 0 };
    let owned = OwnedProcess::from_app(&app, lcid)?;
    let workbooks = get(&app, "Workbooks", lcid, vec![])?;
    let add = call(&workbooks, "Add", DISPATCH_METHOD, lcid, Frame::empty());
    let workbook = add.result.dispatch().ok_or_else(|| format!("Workbooks.Add did not return VT_DISPATCH: {}", add.hr))?;
    let sheet = get(&app, "ActiveSheet", lcid, vec![])?;
    let mut observations = Vec::new();
    for (id, formula, blocked) in [("DA-D-01", "=SEQUENCE(2,3)", false), ("DA-D-02", "=SEQUENCE(2,3)&\"x\"", false), ("DA-D-03", "=SEQUENCE(2,3)", true)] {
        let owner = get(&sheet, "Range", lcid, vec![OwnedVariant::bstr("A1")?])?;
        let spill = get(&sheet, "Range", lcid, vec![OwnedVariant::bstr("A1:C2")?])?;
        let clear_before = call(&spill, "ClearContents", DISPATCH_METHOD, lcid, Frame::empty());
        let blocker = if blocked {
            let cell = get(&sheet, "Range", lcid, vec![OwnedVariant::bstr("B1")?])?;
            Some(call(
                &cell,
                "Value2",
                DISPATCH_PROPERTYPUT,
                lcid,
                Frame::put(OwnedVariant::bstr("blocked")?),
            ))
        } else {
            None
        };
        let write = call(&owner, "Formula2", DISPATCH_PROPERTYPUT, lcid, Frame::put(OwnedVariant::bstr(formula)?));
        let value2 = call(&spill, "Value2", DISPATCH_PROPERTYGET, lcid, Frame::empty());
        observations.push(json!({"schema_version":1,"id":id,"category":"dynamic-array","activation_mode":mode.id(),"fresh_process":true,"formula2_input":formula,"blocked":blocked,"clear_before":brief(&clear_before),"blocker":blocker.as_ref().map(brief),"formula2_write":brief(&write),"owner_reads":read_members(&owner,lcid),"spill_value2":{"call":brief(&value2),"value":rectangular_observation(&value2.result)},"clear_after":brief(&call(&spill,"ClearContents",DISPATCH_METHOD,lcid,Frame::empty())),"raw_pointer_values_recorded":false}));
    }
    let close = call(&workbook, "Close", DISPATCH_METHOD, lcid, Frame::positional(vec![OwnedVariant::boolean(false)]));
    let quit = call(&app, "Quit", DISPATCH_METHOD, lcid, Frame::empty());
    drop(sheet);
    drop(workbook);
    drop(workbooks);
    drop(app);
    let exited = owned.wait();
    Ok(json!({"schema_version":1,"backend":"raw-windows-sys","activation_mode":mode.id(),"workbooks_add":brief(&add),"observations":observations,"cleanup":{"workbook_close":brief(&close),"excel_quit":brief(&quit),"owned_process_exit_verified":exited,"forced_termination":false},"success":exited}))
}

fn dynamic_array_case(
    sheet: &super::com_ptr::ComPtr<super::com_ptr::Dispatch>,
    lcid: u32,
    mode: Mode,
) -> Result<Value, String> {
    let owner = get(sheet, "Range", lcid, vec![OwnedVariant::bstr("A1")?])?;
    let spill = get(sheet, "Range", lcid, vec![OwnedVariant::bstr("A1:C2")?])?;
    let clear_before = call(&spill, "ClearContents", DISPATCH_METHOD, lcid, Frame::empty());
    let formula2_write = call(&owner, "Formula2", DISPATCH_PROPERTYPUT, lcid, Frame::put(OwnedVariant::bstr("=SEQUENCE(2,3)")?));
    let formula_write = call(&owner, "Formula", DISPATCH_PROPERTYGET, lcid, Frame::empty());
    let formula2_read = call(&owner, "Formula2", DISPATCH_PROPERTYGET, lcid, Frame::empty());
    let value2 = call(&spill, "Value2", DISPATCH_PROPERTYGET, lcid, Frame::empty());
    Ok(json!({
        "schema_version":1,"id":"DA-01","category":"dynamic-array","activation_mode":mode.id(),"fresh_process":true,
        "formula2_input":"=SEQUENCE(2,3)","clear_before":brief(&clear_before),"formula2_write":brief(&formula2_write),
        "formula_read":{"call":brief(&formula_write),"value":ObservedVariant::from_variant(&formula_write.result)},
        "formula2_read":{"call":brief(&formula2_read),"value":ObservedVariant::from_variant(&formula2_read.result)},
        "spill_value2":{"call":brief(&value2),"value":rectangular_observation(&value2.result)},
        "clear_after":brief(&call(&spill,"ClearContents",DISPATCH_METHOD,lcid,Frame::empty())),"raw_pointer_values_recorded":false,
    }))
}
