#![cfg(windows)]

use std::thread;
use std::time::{Duration, Instant};

use excel_com::{
    Application, AutomationArray, AutomationValue, CalculationMode, CalculationState, ComApartment,
    FindLookIn, FindOptions, FormulaValue, MixedValue, Range, ReplaceOptions, SpecialCellType,
    SpecialCellValueMask,
};
use windows_sys::Win32::Foundation::{CloseHandle, INVALID_HANDLE_VALUE};
use windows_sys::Win32::System::Diagnostics::ToolHelp::{
    CreateToolhelp32Snapshot, PROCESSENTRY32W, Process32FirstW, Process32NextW, TH32CS_SNAPPROCESS,
};

fn excel_process_count() -> Result<u32, String> {
    // SAFETY: process snapshots accept the documented all-processes flag and zero process ID.
    let snapshot = unsafe { CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0) };
    if snapshot == INVALID_HANDLE_VALUE {
        return Err("CreateToolhelp32Snapshot failed".to_owned());
    }
    let mut entry = PROCESSENTRY32W {
        dwSize: std::mem::size_of::<PROCESSENTRY32W>() as u32,
        ..Default::default()
    };
    let mut count = 0;
    // SAFETY: snapshot is valid and entry has the SDK-required dwSize initialized.
    let mut present = unsafe { Process32FirstW(snapshot, &mut entry) } != 0;
    while present {
        let name = String::from_utf16_lossy(&entry.szExeFile);
        if name
            .trim_end_matches('\0')
            .eq_ignore_ascii_case("EXCEL.EXE")
        {
            count += 1;
        }
        // SAFETY: snapshot remains valid and entry storage is initialized for each enumeration step.
        present = unsafe { Process32NextW(snapshot, &mut entry) } != 0;
    }
    // SAFETY: this function owns the snapshot handle and closes it exactly once.
    unsafe { CloseHandle(snapshot) };
    Ok(count)
}

fn wait_for_excel_exit() -> Result<(), Box<dyn std::error::Error>> {
    let deadline = Instant::now() + Duration::from_secs(20);
    while excel_process_count()? != 0 && Instant::now() < deadline {
        thread::sleep(Duration::from_millis(100));
    }
    assert_eq!(excel_process_count()?, 0, "owned Excel must exit naturally");
    Ok(())
}

fn text_formula(value: FormulaValue) -> Result<String, Box<dyn std::error::Error>> {
    match value {
        FormulaValue::Text(value) => Ok(value),
        other => Err(format!("expected scalar formula text, got {other:?}").into()),
    }
}

fn uniform_bool(value: MixedValue<bool>) -> Result<bool, Box<dyn std::error::Error>> {
    match value {
        MixedValue::Uniform(value) => Ok(value),
        other => Err(format!("expected uniform Boolean, got {other:?}").into()),
    }
}

fn address(range: &Range) -> Result<String, Box<dyn std::error::Error>> {
    Ok(range.external_address(excel_com::ReferenceStyle::A1)?)
}

#[test]
#[ignore = "launches a fresh visible Excel process; run explicitly with one test thread"]
fn formulas_calculation_and_auditing_naturally_exit() -> Result<(), Box<dyn std::error::Error>> {
    assert_eq!(
        excel_process_count()?,
        0,
        "live test requires no pre-existing EXCEL.EXE"
    );
    let apartment = ComApartment::sta()?;
    let application = Application::new(&apartment)?;
    application.set_visible(true)?;
    let original_alerts = application.display_alerts()?;
    let original_reference_style = application.reference_style()?;
    let mut cleanup_workbook = None;

    let outcome = std::panic::catch_unwind(std::panic::AssertUnwindSafe(
        || -> Result<(), Box<dyn std::error::Error>> {
            let workbook = application.workbooks()?.add()?;
            cleanup_workbook = Some(workbook.clone());
            // These properties report a legitimate enum/Boolean only after a workbook
            // exists. Microsoft’s raw C++ Automation sample also creates its workbook
            // before exercising workbook-scoped application state.
            let original_mode = application.calculation_mode()?;
            let original_before_save = application.calculate_before_save()?;
            let worksheet = workbook.worksheets()?.item_by_index(1)?;

            worksheet
                .range("A1")?
                .set_value2(AutomationValue::Number(2.0))?;
            worksheet
                .range("B1")?
                .set_value2(AutomationValue::Number(3.0))?;
            let a1_formula = worksheet.range("C1")?;
            a1_formula.set_formula("=A1+B1")?;
            assert_eq!(text_formula(a1_formula.formula()?)?, "=A1+B1");
            assert_eq!(a1_formula.value2()?, AutomationValue::Number(5.0));

            let r1c1_formula = worksheet.range("D1")?;
            r1c1_formula.set_formula_r1c1("=RC[-3]*2")?;
            assert_eq!(text_formula(r1c1_formula.formula_r1c1()?)?, "=RC[-3]*2");
            assert_eq!(r1c1_formula.value2()?, AutomationValue::Number(4.0));

            let local = worksheet.range("E1")?;
            local.set_formula_local("=1+1")?;
            assert_eq!(text_formula(local.formula_local()?)?, "=1+1");
            let local_r1c1 = worksheet.range("E2")?;
            local_r1c1.set_formula_r1c1_local("=1+1")?;
            assert_eq!(text_formula(local_r1c1.formula_r1c1_local()?)?, "=1+1");

            let formula_matrix = AutomationArray::from_rows(vec![
                vec![
                    AutomationValue::Text("=A1".to_owned()),
                    AutomationValue::Text("=B1".to_owned()),
                ],
                vec![
                    AutomationValue::Text("=A1*2".to_owned()),
                    AutomationValue::Text("=B1*2".to_owned()),
                ],
            ])?;
            let formula_target = worksheet.range("A3:B4")?;
            formula_target.set_formula_array_values(&formula_matrix)?;
            assert!(matches!(formula_target.formula()?, FormulaValue::Array(_)));
            assert!(
                formula_target
                    .set_formula_array_values(&AutomationArray::row(vec![AutomationValue::Text(
                        "=1".to_owned(),
                    )])?)
                    .is_err()
            );
            assert!(formula_target.set_formula("=1").is_err());

            assert!(uniform_bool(worksheet.range("A3:B4")?.has_formula()?)?);
            assert!(!uniform_bool(worksheet.range("A1:B1")?.has_formula()?)?);
            assert!(matches!(
                worksheet.range("A1:C1")?.has_formula()?,
                MixedValue::Mixed
            ));
            assert!(!uniform_bool(worksheet.range("A20")?.has_formula()?)?);

            worksheet
                .range("A6:A8")?
                .set_value2(AutomationValue::Array(AutomationArray::column(vec![
                    AutomationValue::Number(1.0),
                    AutomationValue::Number(2.0),
                    AutomationValue::Number(3.0),
                ])?))?;
            worksheet
                .range("B6:B8")?
                .set_value2(AutomationValue::Array(AutomationArray::column(vec![
                    AutomationValue::Number(10.0),
                    AutomationValue::Number(20.0),
                    AutomationValue::Number(30.0),
                ])?))?;
            let legacy = worksheet.range("C6:C8")?;
            legacy.set_formula_array("=A6:A8*B6:B8")?;
            assert!(uniform_bool(legacy.has_array()?)?);
            assert_eq!(legacy.current_array()?.address_a1()?, "$C$6:$C$8");
            assert_eq!(text_formula(legacy.formula_array()?)?, "=A6:A8*B6:B8");
            assert_eq!(
                legacy.value2()?,
                AutomationValue::Array(AutomationArray::column(vec![
                    AutomationValue::Number(10.0),
                    AutomationValue::Number(40.0),
                    AutomationValue::Number(90.0),
                ])?)
            );
            // Excel normally displays “You cannot change part of an array” for this
            // deliberately invalid write. Suppressing that interactive alert keeps the
            // harness unattended; Excel may report either a COM error or an unchanged
            // formula, so the invariant is that the array formula survives intact.
            application.set_display_alerts(false)?;
            let partial_edit = legacy.cell(1, 1)?.set_formula("=1");
            application.set_display_alerts(original_alerts)?;
            assert!(
                partial_edit.is_err() || text_formula(legacy.formula_array()?)? == "=A6:A8*B6:B8"
            );
            assert_eq!(text_formula(legacy.formula_array()?)?, "=A6:A8*B6:B8");

            let spill_origin = worksheet.range("F1")?;
            spill_origin.set_formula2("=SEQUENCE(2,2)")?;
            assert_eq!(text_formula(spill_origin.formula2()?)?, "=SEQUENCE(2,2)");
            assert!(uniform_bool(spill_origin.has_spill()?)?);
            assert_eq!(spill_origin.spilling_to_range()?.address_a1()?, "$F$1:$G$2");
            assert_eq!(worksheet.range("F2")?.spill_parent()?.address_a1()?, "$F$1");
            assert!(uniform_bool(worksheet.range("F2")?.has_spill()?)?);

            worksheet.range("F1:G2")?.clear_contents()?;
            worksheet
                .range("G2")?
                .set_value2(AutomationValue::Number(99.0))?;
            spill_origin.set_formula2("=SEQUENCE(2,2)")?;
            let blocked_spill_value = spill_origin.value2()?;
            assert!(matches!(blocked_spill_value, AutomationValue::Error(_)));
            worksheet.range("G2")?.clear_contents()?;
            spill_origin.calculate()?;
            assert!(uniform_bool(spill_origin.has_spill()?)?);

            let formula2_r1c1 = worksheet.range("I1")?;
            formula2_r1c1.set_formula2_r1c1("=SEQUENCE(2,2)")?;
            assert_eq!(
                text_formula(formula2_r1c1.formula2_r1c1()?)?,
                "=SEQUENCE(2,2)"
            );
            assert_eq!(
                formula2_r1c1.spilling_to_range()?.address_a1()?,
                "$I$1:$J$2"
            );

            let dependent = worksheet.range("B10")?;
            worksheet
                .range("A10")?
                .set_value2(AutomationValue::Number(10.0))?;
            dependent.set_formula("=A10*2")?;
            assert_eq!(dependent.value2()?, AutomationValue::Number(20.0));
            {
                let guard = application.calculation_mode_guard(CalculationMode::MANUAL)?;
                worksheet
                    .range("A10")?
                    .set_value2(AutomationValue::Number(12.0))?;
                let stale_value = dependent.value2()?;
                assert_eq!(stale_value, AutomationValue::Number(20.0));
                dependent.calculate()?;
                assert_eq!(dependent.value2()?, AutomationValue::Number(24.0));
                worksheet
                    .range("A10")?
                    .set_value2(AutomationValue::Number(13.0))?;
                worksheet.calculate()?;
                assert_eq!(dependent.value2()?, AutomationValue::Number(26.0));
                worksheet
                    .range("A10")?
                    .set_value2(AutomationValue::Number(14.0))?;
                dependent.mark_dirty()?;
                application.calculate()?;
                assert_eq!(dependent.value2()?, AutomationValue::Number(28.0));
                assert_eq!(application.calculation_state()?, CalculationState::DONE);
                guard.restore()?;
            }
            assert_eq!(application.calculation_mode()?, original_mode);
            application.calculate_full()?;
            assert_eq!(application.calculation_state()?, CalculationState::DONE);
            application.calculate_full_rebuild()?;
            assert_eq!(application.calculation_state()?, CalculationState::DONE);

            let source = worksheet.range("A12")?;
            source.set_value2(AutomationValue::Number(7.0))?;
            let direct = worksheet.range("B12")?;
            direct.set_formula("=A12+1")?;
            let transitive = worksheet.range("C12")?;
            transitive.set_formula("=B12+1")?;
            let direct_precedents = address(&direct.direct_precedents()?)?;
            let transitive_precedents = address(&transitive.precedents()?)?;
            let direct_dependents = address(&source.direct_dependents()?)?;
            let dependents = address(&source.dependents()?)?;
            assert!(direct_precedents.contains("$A$12"));
            assert!(transitive_precedents.contains("$A$12:$B$12"));
            assert!(direct_dependents.contains("$B$12"));
            assert!(dependents.contains("$B$12:$C$12"));
            let cross_sheet = workbook.worksheets()?.add(Default::default())?;
            cross_sheet.range("A1")?.set_formula("=Sheet1!A12")?;
            let cross_sheet_precedents = cross_sheet.range("A1")?.direct_precedents();
            assert!(cross_sheet_precedents.is_err());

            let discovery = worksheet.range("A20:C22")?;
            worksheet
                .range("A20")?
                .set_value2(AutomationValue::Text("target".to_owned()))?;
            worksheet
                .range("A21")?
                .set_value2(AutomationValue::Text("target".to_owned()))?;
            worksheet
                .range("A22")?
                .set_value2(AutomationValue::Number(42.0))?;
            worksheet.range("B20")?.set_formula("=A22*2")?;
            let formulas = discovery.special_cells(SpecialCellType::FORMULAS, None)?;
            assert!(address(&formulas)?.contains("$B$20"));
            let constants = discovery.special_cells(
                SpecialCellType::CONSTANTS,
                Some(SpecialCellValueMask::TEXT | SpecialCellValueMask::NUMBERS),
            )?;
            assert!(address(&constants)?.contains("$A$20"));
            let blanks = discovery.special_cells(SpecialCellType::BLANKS, None)?;
            assert!(address(&blanks)?.contains("$C$20"));
            let visible = discovery.special_cells(SpecialCellType::VISIBLE, None)?;
            assert!(address(&visible)?.contains("$A$20"));
            let no_formula_cells = worksheet
                // Excel expands some one-cell SpecialCells searches beyond the
                // receiver; a blank multi-cell range preserves the no-match case.
                .range("Z100:AA101")?
                .special_cells(SpecialCellType::FORMULAS, None);
            assert!(no_formula_cells.is_err());

            let matches = worksheet.range("A20:A21")?;
            let text_options = FindOptions::default();
            let first = matches
                .find(&AutomationValue::Text("target".to_owned()), &text_options)?
                .expect("text must be found");
            assert!(matches.find_next(Some(&first))?.is_some());
            assert!(matches.find_previous(Some(&first))?.is_some());
            let all = matches
                .find_all(&AutomationValue::Text("target".to_owned()), &text_options)?
                .collect::<Result<Vec<_>, _>>()?;
            assert_eq!(all.len(), 2);
            assert!(
                discovery
                    .find(&AutomationValue::Number(42.0), &FindOptions::default())?
                    .is_some()
            );
            assert!(
                discovery
                    .find(
                        &AutomationValue::Text("A22".to_owned()),
                        &FindOptions {
                            look_in: Some(FindLookIn::FORMULAS),
                            ..FindOptions::default()
                        },
                    )?
                    .is_some()
            );
            assert!(
                matches
                    .find(
                        &AutomationValue::Text("missing".to_owned()),
                        &FindOptions::default(),
                    )?
                    .is_none()
            );
            assert!(matches.replace(
                &AutomationValue::Text("target".to_owned()),
                &AutomationValue::Text("updated".to_owned()),
                &ReplaceOptions::default(),
            )?);
            assert_eq!(
                matches.value2()?,
                AutomationValue::Array(AutomationArray::column(vec![
                    AutomationValue::Text("updated".to_owned()),
                    AutomationValue::Text("updated".to_owned()),
                ])?)
            );
            let formula_replace = worksheet.range("C20")?;
            formula_replace.set_formula("=SUM(1,2)")?;
            assert!(formula_replace.replace(
                &AutomationValue::Text("SUM".to_owned()),
                &AutomationValue::Text("AVERAGE".to_owned()),
                &ReplaceOptions::default(),
            )?);
            assert!(text_formula(formula_replace.formula()?)?.contains("AVERAGE"));

            application.set_calculate_before_save(!original_before_save)?;
            assert_eq!(application.calculate_before_save()?, !original_before_save);
            application.set_calculate_before_save(original_before_save)?;
            assert_eq!(application.calculate_before_save()?, original_before_save);
            assert_eq!(application.display_alerts()?, original_alerts);
            assert_eq!(application.reference_style()?, original_reference_style);

            cleanup_workbook
                .take()
                .expect("fresh workbook must be available for cleanup")
                .close_without_saving()?;
            Ok(())
        },
    ));
    if let Some(workbook) = cleanup_workbook.take() {
        let _ = workbook.close_without_saving();
    }
    let quit = application.quit();
    match outcome {
        Ok(result) => result?,
        Err(payload) => std::panic::resume_unwind(payload),
    }
    quit?;
    wait_for_excel_exit()?;
    Ok(())
}
