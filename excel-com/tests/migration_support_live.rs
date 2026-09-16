//! End-to-end real-Excel evidence for the reusable workbook-migration surface.
//!
//! This suite opens a copied repository `.xlsm` through a crate-owned private
//! Excel server with macros disabled, exercises formula/name/range discovery,
//! saves a separate `.xlsx`, reopens it, and waits for the exact owned process
//! to exit. It deliberately contains no application-specific migration logic.

#![cfg(windows)]

mod support;

use std::fs;
use std::time::Duration;

use excel_com::{
    AutomationValue, CalculationMode, ComApartment, FormulaValue, MixedValue, NameAddOptions,
    NameRefersTo, OwnedApplication, SafeWorkbookOpenOptions, SaveChanges, WorkbookCloseOptions,
    WorkbookOpenOptions, WorkbookSaveAsOptions, XlFileFormat, XlUpdateLinks,
};
use support::Fixture;

fn formula_text(value: FormulaValue) -> Result<String, Box<dyn std::error::Error>> {
    match value {
        FormulaValue::Text(value) => Ok(value),
        other => Err(format!("expected one formula string, got {other:?}").into()),
    }
}

fn uniform_bool(value: MixedValue<bool>) -> Result<bool, Box<dyn std::error::Error>> {
    match value {
        MixedValue::Uniform(value) => Ok(value),
        other => Err(format!("expected one Boolean state, got {other:?}").into()),
    }
}

#[test]
#[ignore = "requires Windows desktop Excel and launches a private Excel process"]
fn migration_surface_round_trips_and_owned_excel_exits() -> Result<(), Box<dyn std::error::Error>> {
    let input = Fixture::BlankXlsm.copy_for_test()?;
    let output = std::env::temp_dir().join(format!(
        "excel-com-migration-support-{}.xlsx",
        std::process::id()
    ));
    let _ = fs::remove_file(&output);

    let apartment = ComApartment::sta()?;
    let excel = OwnedApplication::new(&apartment)?;
    let diagnostics = excel.diagnostics()?;
    assert_eq!(diagnostics.ownership, excel_com::SessionOwnership::Owned);
    assert!(diagnostics.process_id.is_some());
    eprintln!(
        "migration-support Excel version={} bitness={:?} process_id_observed={}",
        diagnostics.version,
        diagnostics.bitness,
        diagnostics.process_id.is_some()
    );
    excel.set_visible(false)?;

    let initial_user_control = excel.user_control()?;
    excel.set_user_control(false)?;
    assert!(!excel.user_control()?);
    let initial_events = excel.enable_events()?;
    let initial_alerts = excel.display_alerts()?;
    let initial_links = excel.ask_to_update_links()?;
    let initial_security = excel.automation_security()?;

    let mut cleanup_workbook = None;
    let outcome = std::panic::catch_unwind(std::panic::AssertUnwindSafe(
        || -> Result<(), Box<dyn std::error::Error>> {
            let events = excel.enable_events_guard(false)?;
            let alerts = excel.display_alerts_guard(false)?;
            let links = excel.ask_to_update_links_guard(false)?;
            assert!(!excel.enable_events()?);
            assert!(!excel.display_alerts()?);
            assert!(!excel.ask_to_update_links()?);

            let workbooks = excel.workbooks()?;
            let workbook = workbooks.open_safely(
                input.path(),
                SafeWorkbookOpenOptions {
                    open: WorkbookOpenOptions {
                        update_links: Some(XlUpdateLinks::DO_NOT_UPDATE),
                        read_only: Some(true),
                        ..WorkbookOpenOptions::new()
                    },
                    link_prompt: Some(false),
                },
            )?;
            cleanup_workbook = Some(workbook.clone());
            assert!(workbook.read_only()?);
            assert_eq!(excel.automation_security()?, initial_security);

            // Some licensed Excel hosts reject Application.Calculation while a
            // copied fixture is the only open workbook (0x800A03EC). The
            // PowerShell reference also treats manual calculation as a
            // best-effort optimization rather than a migration precondition.
            let calculation = match excel.calculation_mode_guard(CalculationMode::MANUAL) {
                Ok(guard) => Some(guard),
                Err(error) => {
                    eprintln!("manual calculation mode unavailable on this Excel host: {error}");
                    None
                }
            };
            let worksheets = workbook.worksheets()?;
            let worksheet = worksheets.item_by_index(1)?;
            let original_sheet_name = worksheet.name()?;

            let dynamic = worksheet.range("A1")?;
            dynamic.set_formula2("=SEQUENCE(2,2)")?;
            assert_eq!(formula_text(dynamic.formula2()?)?, "=SEQUENCE(2,2)");
            assert!(dynamic.worksheet()?.is_same_object(&worksheet)?);

            worksheet
                .range("D1:D3")?
                .set_value2(AutomationValue::Array(excel_com::AutomationArray::column(
                    vec![
                        AutomationValue::Number(1.0),
                        AutomationValue::Number(2.0),
                        AutomationValue::Number(3.0),
                    ],
                )?))?;
            let legacy_array = worksheet.range("E1:E3")?;
            legacy_array.set_formula_array("=D1:D3*2")?;
            assert!(uniform_bool(legacy_array.has_array()?)?);
            assert_eq!(legacy_array.current_array()?.address_a1()?, "$E$1:$E$3");
            assert_eq!(formula_text(legacy_array.formula_array()?)?, "=D1:D3*2");

            let workbook_target = worksheet.range("G1")?;
            let local_target = worksheet.range("G2")?;
            let workbook_name = workbook.names()?.add(&NameAddOptions {
                name: "MigrationHeader",
                refers_to: NameRefersTo::Range(&workbook_target),
                visible: Some(true),
            })?;
            let local_name = worksheet.names()?.add(&NameAddOptions {
                name: "MigrationHeader",
                refers_to: NameRefersTo::Range(&local_target),
                visible: Some(true),
            })?;
            assert_eq!(workbook_name.refers_to_range()?.address_a1()?, "$G$1");
            assert_eq!(local_name.refers_to_range()?.address_a1()?, "$G$2");
            assert_ne!(workbook_name.name()?, local_name.name()?);

            let used = worksheet.used_range()?;
            let formulas = used
                .try_formula_cells(None)?
                .ok_or("migration fixture should contain formula cells")?;
            assert!(formulas.cell_count()? >= 4);

            workbook.save_as(
                &output,
                WorkbookSaveAsOptions {
                    file_format: Some(XlFileFormat::OPEN_XML_WORKBOOK),
                    ..WorkbookSaveAsOptions::new()
                },
            )?;
            drop((formulas, used, local_name, workbook_name));
            drop((local_target, workbook_target, legacy_array, dynamic));
            drop((worksheet, worksheets));
            cleanup_workbook
                .take()
                .expect("opened workbook must remain available for cleanup")
                .close(WorkbookCloseOptions {
                    save_changes: SaveChanges::Discard,
                    ..WorkbookCloseOptions::new()
                })?;

            let reopened = workbooks.open(
                &output,
                WorkbookOpenOptions {
                    update_links: Some(XlUpdateLinks::DO_NOT_UPDATE),
                    read_only: Some(true),
                    ..WorkbookOpenOptions::new()
                },
            )?;
            cleanup_workbook = Some(reopened.clone());
            assert_eq!(reopened.file_format()?, XlFileFormat::OPEN_XML_WORKBOOK);
            assert!(!reopened.has_vb_project()?);
            let reopened_sheet = reopened.worksheets()?.item_by_name(&original_sheet_name)?;
            assert_eq!(
                formula_text(reopened_sheet.range("A1")?.formula2()?)?,
                "=SEQUENCE(2,2)"
            );
            assert!(uniform_bool(reopened_sheet.range("E1")?.has_array()?)?);
            assert_eq!(
                reopened
                    .names()?
                    .item_by_name("MigrationHeader")?
                    .refers_to_range()?
                    .address_a1()?,
                "$G$1"
            );
            assert_eq!(
                reopened_sheet
                    .names()?
                    .item_by_name("MigrationHeader")?
                    .refers_to_range()?
                    .address_a1()?,
                "$G$2"
            );
            drop(reopened_sheet);
            cleanup_workbook
                .take()
                .expect("reopened workbook must remain available for cleanup")
                .close(WorkbookCloseOptions {
                    save_changes: SaveChanges::Discard,
                    ..WorkbookCloseOptions::new()
                })?;
            drop(workbooks);

            if let Some(calculation) = calculation {
                if let Err(error) = calculation.restore() {
                    eprintln!(
                        "calculation mode restoration unavailable on this Excel host: {error}"
                    );
                }
            }
            links.restore()?;
            alerts.restore()?;
            events.restore()?;
            assert_eq!(excel.enable_events()?, initial_events);
            assert_eq!(excel.display_alerts()?, initial_alerts);
            assert_eq!(excel.ask_to_update_links()?, initial_links);
            Ok(())
        },
    ));

    if let Some(workbook) = cleanup_workbook.take() {
        let _ = workbook.close_without_saving();
    }
    if initial_user_control {
        let _ = excel.set_user_control(true);
    }
    let exit = excel.quit_and_wait(Duration::from_secs(30));
    let _ = fs::remove_file(&output);

    match outcome {
        Ok(result) => result?,
        Err(payload) => std::panic::resume_unwind(payload),
    }
    assert!(exit?.exited);
    Ok(())
}

#[test]
#[ignore = "requires Windows desktop Excel and launches two private Excel processes"]
fn owned_application_does_not_reuse_or_close_an_existing_session()
-> Result<(), Box<dyn std::error::Error>> {
    let apartment = ComApartment::sta()?;
    let existing = OwnedApplication::new(&apartment)?;
    existing.set_visible(false)?;
    let existing_diagnostics = existing.diagnostics()?;
    let created = OwnedApplication::new(&apartment)?;
    created.set_visible(false)?;
    let created_diagnostics = created.diagnostics()?;

    let existing_pid = existing_diagnostics
        .process_id
        .ok_or("first owned Excel process id was not observable")?;
    let created_pid = created_diagnostics
        .process_id
        .ok_or("second owned Excel process id was not observable")?;
    if existing_pid == created_pid {
        let _ = created.quit();
        let _ = existing.quit();
        return Err("OwnedApplication::new reused an existing Excel process".into());
    }

    let created_exit = created.quit_and_wait(Duration::from_secs(30))?;
    let existing_version_after_created_exit = existing.version()?;
    let existing_exit = existing.quit_and_wait(Duration::from_secs(30))?;

    assert!(created_exit.exited);
    assert!(!existing_version_after_created_exit.is_empty());
    assert!(existing_exit.exited);
    Ok(())
}
