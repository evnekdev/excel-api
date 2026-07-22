#![cfg(windows)]

use std::thread;
use std::time::{Duration, Instant};

use excel_com::{
    Application, AutomationArray, AutomationValue, ComApartment, FormulaConversionOptions,
    NameAddOptions, NameRefersTo, RangeAddressOptions, ReferenceStyle,
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
        if String::from_utf16_lossy(&entry.szExeFile)
            .trim_end_matches('\0')
            .eq_ignore_ascii_case("EXCEL.EXE")
        {
            count += 1;
        }
        // SAFETY: snapshot and initialized entry remain valid for enumeration.
        present = unsafe { Process32NextW(snapshot, &mut entry) } != 0;
    }
    // SAFETY: this function owns the snapshot and closes it exactly once.
    unsafe { CloseHandle(snapshot) };
    Ok(count)
}

fn wait_for_zero_excel_processes() -> Result<(), String> {
    let deadline = Instant::now() + Duration::from_secs(15);
    while Instant::now() < deadline {
        if excel_process_count()? == 0 {
            return Ok(());
        }
        thread::sleep(Duration::from_millis(100));
    }
    Err(format!(
        "crate-owned Excel process did not exit naturally; count={}",
        excel_process_count()?
    ))
}

fn external_a1(range: &excel_com::Range) -> Result<String, excel_com::ExcelComError> {
    range.external_address(ReferenceStyle::A1)
}

#[test]
#[ignore = "launches a fresh visible Excel process; run explicitly"]
fn references_names_and_evaluation_naturally_exit() -> Result<(), Box<dyn std::error::Error>> {
    assert_eq!(
        excel_process_count()?,
        0,
        "live test requires no pre-existing EXCEL.EXE"
    );
    let apartment = ComApartment::sta()?;
    let application = Application::new(&apartment)?;
    application.set_visible(true)?;
    let mut cleanup_workbook = None;

    let outcome = std::panic::catch_unwind(std::panic::AssertUnwindSafe(
        || -> Result<(), Box<dyn std::error::Error>> {
            let alerts = application.display_alerts_guard(false)?;
            let original_style = application.reference_style()?;
            let workbooks = application.workbooks()?;
            let workbook = workbooks.add()?;
            cleanup_workbook = Some(workbook.clone());
            let worksheet = workbook.worksheets()?.item_by_index(1)?;
            worksheet.set_name("Sheet Name")?;

            let selected = worksheet.range("A1:X10")?;
            assert_eq!(selected.row_count()?, 10);
            assert_eq!(selected.column_count()?, 24);
            assert_eq!(selected.address_a1()?, "$A$1:$X$10");
            assert_eq!(selected.address_r1c1()?, "R1C1:R10C24");
            assert_eq!(
                worksheet.range_between("A1", "X10")?.address_a1()?,
                selected.address_a1()?
            );
            assert_eq!(
                worksheet.range_from_cells(1, 1, 10, 24)?.address_a1()?,
                selected.address_a1()?
            );
            assert_eq!(
                worksheet.range_r1c1("R1C1:R10C24")?.address_a1()?,
                selected.address_a1()?
            );

            assert_eq!(worksheet.range("A:A")?.address_a1()?, "$A:$A");
            assert_eq!(worksheet.range("1:1")?.address_a1()?, "$1:$1");
            assert_eq!(worksheet.range("'Sheet Name'!A1")?.address_a1()?, "$A$1");
            let b2 = worksheet.range("B2")?;
            let a1 = worksheet.range("A1")?;
            assert_eq!(
                b2.address_with_options(&RangeAddressOptions {
                    row_absolute: Some(false),
                    column_absolute: Some(false),
                    reference_style: ReferenceStyle::R1C1,
                    external: Some(false),
                    relative_to: Some(&a1),
                })?,
                "R[1]C[1]"
            );
            assert_eq!(
                a1.address_with_options(&RangeAddressOptions {
                    row_absolute: Some(false),
                    column_absolute: Some(false),
                    reference_style: ReferenceStyle::R1C1,
                    external: Some(false),
                    relative_to: Some(&b2),
                })?,
                "R[-1]C[-1]"
            );
            assert_eq!(
                a1.address_with_options(&RangeAddressOptions {
                    row_absolute: Some(false),
                    column_absolute: Some(false),
                    reference_style: ReferenceStyle::R1C1,
                    external: Some(false),
                    relative_to: Some(&a1),
                })?,
                "RC"
            );
            let external = selected.external_address(ReferenceStyle::A1)?;
            assert!(external.contains("Sheet Name"));
            eprintln!(
                "addresses a1={} r1c1={} external={external}",
                selected.address_a1()?,
                selected.address_r1c1()?,
            );

            assert_eq!(
                application.convert_formula(
                    "$A$1:$X$10",
                    ReferenceStyle::A1,
                    ReferenceStyle::R1C1,
                    &FormulaConversionOptions::default(),
                )?,
                "R1C1:R10C24"
            );
            assert_eq!(
                application.convert_formula(
                    "R1C1:R10C24",
                    ReferenceStyle::R1C1,
                    ReferenceStyle::A1,
                    &FormulaConversionOptions::default(),
                )?,
                "$A$1:$X$10"
            );
            assert!(
                application
                    .convert_formula(
                        "=SUM(A1:X10)",
                        ReferenceStyle::A1,
                        ReferenceStyle::R1C1,
                        &FormulaConversionOptions::default(),
                    )?
                    .starts_with("=SUM(")
            );
            assert_eq!(
                application.convert_formula(
                    "R[1]C[1]",
                    ReferenceStyle::R1C1,
                    ReferenceStyle::A1,
                    &FormulaConversionOptions {
                        to_absolute: None,
                        relative_to: Some(&a1),
                    },
                )?,
                "B2"
            );

            let direct_a1 = {
                let style = application.reference_style_guard(ReferenceStyle::A1)?;
                let result = worksheet.range("R1C1").and_then(|range| range.address_a1());
                style.restore()?;
                result
            };
            let direct_r1c1 = {
                let style = application.reference_style_guard(ReferenceStyle::R1C1)?;
                let result = worksheet.range("R1C1").and_then(|range| range.address_a1());
                style.restore()?;
                result
            };
            eprintln!(
                "direct Worksheet.Range R1C1 under A1={direct_a1:?}; under R1C1={direct_r1c1:?}"
            );
            assert_eq!(application.reference_style()?, original_style);

            let input = worksheet.range("B2:B3")?;
            input.set_value(AutomationValue::Array(AutomationArray::new(
                2,
                1,
                vec![AutomationValue::Number(2.0), AutomationValue::Number(3.0)],
            )?))?;
            let names = workbook.names()?;
            let initial_count = names.count()?;
            let input_name = names.add(&NameAddOptions {
                name: "InputRange",
                refers_to: NameRefersTo::Range(&input),
                visible: Some(true),
            })?;
            assert_eq!(names.count()?, initial_count + 1);
            assert_eq!(input_name.name()?, "InputRange");
            let refers_to = input_name.refers_to()?;
            let refers_to_r1c1 = input_name.refers_to_r1c1()?;
            assert!(refers_to.starts_with('='));
            assert!(refers_to_r1c1.starts_with('='));
            eprintln!("InputRange refers_to={refers_to}; refers_to_r1c1={refers_to_r1c1}");
            assert!(input_name.visible()?);
            let resolved = input_name.range()?;
            assert_eq!(external_a1(&resolved)?, external_a1(&input)?);
            eprintln!(
                "InputRange canonical identity={}",
                input.is_same_object(&resolved)?
            );
            assert_eq!(names.item_by_name("InputRange")?.name()?, "InputRange");
            assert_eq!(
                names.item_by_index(initial_count + 1)?.name()?,
                "InputRange"
            );
            assert!(
                names
                    .iter()?
                    .collect::<Result<Vec<_>, _>>()?
                    .iter()
                    .any(|name| name.name().as_deref() == Ok("InputRange"))
            );

            let evaluated = application.evaluate_range("InputRange")?;
            assert_eq!(external_a1(&evaluated)?, external_a1(&input)?);
            assert_eq!(
                application.evaluate_value("SUM(InputRange)")?,
                AutomationValue::Number(5.0)
            );
            assert!(application.evaluate_value("InputRange").is_err());
            assert!(application.evaluate_range("SUM(InputRange)").is_err());

            let a1_name = names.add(&NameAddOptions {
                name: "A1Target",
                refers_to: NameRefersTo::A1("='Sheet Name'!$B$2:$B$3"),
                visible: None,
            })?;
            assert_eq!(external_a1(&a1_name.range()?)?, external_a1(&input)?);
            let r1c1_name = names.add(&NameAddOptions {
                name: "ReferenceTarget",
                refers_to: NameRefersTo::R1C1("='Sheet Name'!R2C2:R3C2"),
                visible: None,
            })?;
            assert_eq!(external_a1(&r1c1_name.range()?)?, external_a1(&input)?);
            let constant_name = names.add(&NameAddOptions {
                name: "ConstantValue",
                refers_to: NameRefersTo::Formula("=42"),
                visible: None,
            })?;
            assert!(constant_name.refers_to().is_ok());
            assert!(constant_name.range().is_err());
            assert_eq!(
                application.evaluate_value("ConstantValue")?,
                AutomationValue::Number(42.0)
            );

            let local_names = worksheet.names()?;
            let local_name = local_names.add(&NameAddOptions {
                name: "LocalInput",
                refers_to: NameRefersTo::Range(&input),
                visible: Some(true),
            })?;
            assert!(local_name.name()?.contains("LocalInput"));
            eprintln!("local name={}", local_name.name()?);
            assert_eq!(
                external_a1(&worksheet.evaluate_range("LocalInput")?)?,
                external_a1(&input)?
            );
            let local_collision = local_names.add(&NameAddOptions {
                name: "InputRange",
                refers_to: NameRefersTo::A1("='Sheet Name'!$A$1"),
                visible: None,
            })?;
            eprintln!(
                "scope collision workbook={} worksheet={} application={:?} worksheet={:?}",
                input_name.name()?,
                local_collision.name()?,
                application
                    .evaluate_range("InputRange")
                    .and_then(|range| external_a1(&range)),
                worksheet
                    .evaluate_range("InputRange")
                    .and_then(|range| external_a1(&range)),
            );

            worksheet.set_name("O'Brien")?;
            assert!(
                input
                    .external_address(ReferenceStyle::A1)?
                    .contains("O''Brien")
            );
            assert_eq!(
                worksheet.range("'O''Brien'!B2:B3")?.address_a1()?,
                "$B$2:$B$3"
            );

            assert!(worksheet.range("A\0").is_err());
            assert!(names.item_by_name("Input\0Range").is_err());
            assert!(worksheet.cell(0, 1).is_err());

            local_collision.delete()?;
            local_name.delete()?;
            constant_name.delete()?;
            r1c1_name.delete()?;
            a1_name.delete()?;
            input_name.delete()?;
            assert!(names.item_by_name("InputRange").is_err());
            cleanup_workbook
                .take()
                .expect("fresh workbook must be available for cleanup")
                .close_without_saving()?;
            alerts.restore()?;
            assert_eq!(application.reference_style()?, original_style);
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
    wait_for_zero_excel_processes()?;
    Ok(())
}
