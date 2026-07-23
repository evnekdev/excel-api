#![cfg(windows)]

use std::thread;
use std::time::{Duration, Instant};

use excel_com::{
    AutoFilterOptions, AutomationArray, AutomationValue, ComApartment, ListObjectAddOptions,
    MixedValue, OwnedApplication, PasteOperation, PasteSpecialOptions, PasteType,
    RangeInsertOptions, RemoveDuplicatesOptions, SortDataOption, SortOrder, TableHeaderMode,
    TotalsCalculation, ValidationAddOptions, ValidationType,
};
use windows_sys::Win32::Foundation::{CloseHandle, INVALID_HANDLE_VALUE};
use windows_sys::Win32::System::Diagnostics::ToolHelp::{
    CreateToolhelp32Snapshot, PROCESSENTRY32W, Process32FirstW, Process32NextW, TH32CS_SNAPPROCESS,
};

fn excel_process_count() -> Result<u32, String> {
    // SAFETY: the documented process snapshot API accepts these arguments.
    let snapshot = unsafe { CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0) };
    if snapshot == INVALID_HANDLE_VALUE {
        return Err("CreateToolhelp32Snapshot failed".to_owned());
    }
    let mut entry = PROCESSENTRY32W {
        dwSize: std::mem::size_of::<PROCESSENTRY32W>() as u32,
        ..Default::default()
    };
    let mut count = 0;
    // SAFETY: the snapshot is valid and the entry size is initialized as required.
    let mut present = unsafe { Process32FirstW(snapshot, &mut entry) } != 0;
    while present {
        if String::from_utf16_lossy(&entry.szExeFile)
            .trim_end_matches('\0')
            .eq_ignore_ascii_case("EXCEL.EXE")
        {
            count += 1;
        }
        // SAFETY: the snapshot and initialized entry remain valid during enumeration.
        present = unsafe { Process32NextW(snapshot, &mut entry) } != 0;
    }
    // SAFETY: this function owns and closes the snapshot exactly once.
    unsafe { CloseHandle(snapshot) };
    Ok(count)
}

fn wait_for_excel_exit() -> Result<(), Box<dyn std::error::Error>> {
    let deadline = Instant::now() + Duration::from_secs(20);
    while excel_process_count()? != 0 && Instant::now() < deadline {
        thread::sleep(Duration::from_millis(100));
    }
    assert_eq!(
        excel_process_count()?,
        0,
        "crate-owned Excel must exit naturally"
    );
    Ok(())
}

#[test]
#[ignore = "launches a fresh visible Excel process; run explicitly with one test thread"]
fn structured_data_operations_naturally_exit() -> Result<(), Box<dyn std::error::Error>> {
    assert_eq!(
        excel_process_count()?,
        0,
        "live test requires no pre-existing EXCEL.EXE"
    );
    let apartment = ComApartment::sta()?;
    let application = OwnedApplication::new(&apartment)?;
    application.set_visible(true)?;
    let mut cleanup_workbook = None;
    let outcome = std::panic::catch_unwind(std::panic::AssertUnwindSafe(
        || -> Result<(), Box<dyn std::error::Error>> {
            let alerts = application.display_alerts_guard(false)?;
            let workbook = application.workbooks()?.add()?;
            cleanup_workbook = Some(workbook.clone());
            let worksheet = workbook.worksheets()?.item_by_index(1)?;
            worksheet
                .range("A1:C4")?
                .set_value2(AutomationValue::Array(AutomationArray::from_rows(vec![
                    vec![
                        AutomationValue::Text("Item".to_owned()),
                        AutomationValue::Text("Quantity".to_owned()),
                        AutomationValue::Text("Price".to_owned()),
                    ],
                    vec![
                        AutomationValue::Text("Pear".to_owned()),
                        AutomationValue::Number(3.0),
                        AutomationValue::Number(2.0),
                    ],
                    vec![
                        AutomationValue::Text("Apple".to_owned()),
                        AutomationValue::Number(1.0),
                        AutomationValue::Number(4.0),
                    ],
                    vec![
                        AutomationValue::Text("Pear".to_owned()),
                        AutomationValue::Number(3.0),
                        AutomationValue::Number(2.0),
                    ],
                ])?))?;

            let source = worksheet.range("A1:C4")?;
            let tables = worksheet.list_objects()?;
            let table = tables.add_from_range(&ListObjectAddOptions {
                source: &source,
                has_headers: TableHeaderMode::YES,
                destination: None,
                table_style_name: None,
            })?;
            table.set_name("SalesData")?;
            assert_eq!(tables.count()?, 1);
            assert_eq!(tables.item_by_name("SalesData")?.name()?, "SalesData");
            assert_eq!(table.list_columns()?.count()?, 3);
            assert_eq!(table.list_rows()?.count()?, 3);
            assert!(table.header_row_range()?.is_some());
            assert!(table.data_body_range()?.is_some());
            // This Excel build reports no InsertRowRange until the UI's insert row is active.
            assert!(table.insert_row_range()?.is_none());

            let amount = table.list_columns()?.add(None)?;
            amount.set_name("Amount")?;
            amount.set_calculated_column_formula("=[@Quantity]*[@Price]")?;
            assert_eq!(table.list_columns()?.count()?, 4);
            assert_eq!(table.list_rows()?.item(1)?.range()?.column_count()?, 4);
            table.set_show_totals(true)?;
            amount.set_totals_calculation(TotalsCalculation::SUM)?;
            assert!(table.totals_row_range()?.is_some());

            let filter_range = table.range()?;
            filter_range.apply_auto_filter(&AutoFilterOptions {
                field: 1,
                criterion1: Some(excel_com::FilterCriterion::Value(AutomationValue::Text(
                    "Pear".to_owned(),
                ))),
                operator: None,
                criterion2: None,
                visible_dropdown: Some(true),
            })?;
            // Table filters have an AutoFilter object without necessarily setting
            // worksheet AutoFilterMode on this Excel build.
            assert!(table.auto_filter()?.filters()?.item(1)?.is_on()?);
            filter_range.clear_auto_filter()?;
            assert!(!worksheet.filter_mode()?);

            let sort = table.sort()?;
            let fields = sort.sort_fields()?;
            fields.clear()?;
            let quantity = table
                .list_columns()?
                .item_by_name("Quantity")?
                .data_body_range()?
                .expect("data rows");
            let _field = fields.add(
                &quantity,
                SortOrder::ASCENDING,
                Some(SortDataOption::NORMAL),
            )?;
            // A ListObject-owned Sort is already bound to its table; this Excel
            // build rejects SetRange on that object.
            sort.set_header(TableHeaderMode::YES)?;
            sort.apply()?;
            assert_eq!(
                table.list_rows()?.item(1)?.range()?.cell(1, 2)?.value2()?,
                AutomationValue::Number(1.0)
            );

            let validation = worksheet.range("F2")?.validation()?;
            validation.add(&ValidationAddOptions {
                validation_type: ValidationType::LIST,
                alert_style: None,
                operator: None,
                formula1: Some("yes,no"),
                formula2: None,
            })?;
            validation.set_input_title("Choice")?;
            assert_eq!(validation.input_title()?, "Choice");
            assert_eq!(validation.validation_type()?, ValidationType::LIST);

            worksheet
                .range("H1:I4")?
                .set_value2(AutomationValue::Array(AutomationArray::from_rows(vec![
                    vec![
                        AutomationValue::Text("A".to_owned()),
                        AutomationValue::Text("B".to_owned()),
                    ],
                    vec![AutomationValue::Number(1.0), AutomationValue::Number(2.0)],
                    vec![AutomationValue::Number(1.0), AutomationValue::Number(2.0)],
                    vec![AutomationValue::Number(3.0), AutomationValue::Number(4.0)],
                ])?))?;
            worksheet
                .range("H1:I4")?
                .remove_duplicates(&RemoveDuplicatesOptions {
                    columns: vec![1, 2],
                    header: TableHeaderMode::YES,
                })?;
            assert_eq!(worksheet.range("H1:I3")?.row_count()?, 3);
            assert!(worksheet.range("A2")?.current_region()?.row_count()? >= 4);
            assert!(worksheet.used_range()?.row_count()? >= 4);

            worksheet
                .range("K1")?
                .set_value2(AutomationValue::Number(5.0))?;
            worksheet.range("K1")?.copy(Some(&worksheet.range("L1")?))?;
            assert_eq!(
                worksheet.range("L1")?.value2()?,
                AutomationValue::Number(5.0)
            );
            worksheet.range("L1")?.copy(None)?;
            worksheet.range("M1")?.paste_special(&PasteSpecialOptions {
                paste: PasteType::VALUES,
                operation: PasteOperation::NONE,
                skip_blanks: false,
                transpose: false,
            })?;
            assert_eq!(
                worksheet.range("M1")?.value2()?,
                AutomationValue::Number(5.0)
            );
            worksheet
                .range("N1")?
                .insert(&RangeInsertOptions::default())?;
            worksheet.range("N1")?.clear()?;
            worksheet.range("A1")?.entire_column()?.set_hidden(true)?;
            assert!(matches!(
                worksheet.range("A1")?.entire_column()?.hidden()?,
                MixedValue::Uniform(true)
            ));
            worksheet.range("A1")?.entire_column()?.set_hidden(false)?;

            let ordinary = table.range()?;
            table.unlist()?;
            assert!(tables.count()? == 0);
            assert_eq!(
                ordinary.cell(1, 1)?.value2()?,
                AutomationValue::Text("Item".to_owned())
            );
            cleanup_workbook
                .take()
                .expect("workbook")
                .close_without_saving()?;
            alerts.restore()?;
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
