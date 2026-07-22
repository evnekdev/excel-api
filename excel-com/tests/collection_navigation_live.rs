#![cfg(windows)]

use std::thread;
use std::time::{Duration, Instant};

use excel_com::{
    Application, AutomationArray, AutomationValue, ComApartment, SaveChanges, WorkbookCloseOptions,
};
use windows_sys::Win32::Foundation::{CloseHandle, INVALID_HANDLE_VALUE};
use windows_sys::Win32::System::Diagnostics::ToolHelp::{
    CreateToolhelp32Snapshot, PROCESSENTRY32W, Process32FirstW, Process32NextW, TH32CS_SNAPPROCESS,
};

fn excel_process_count() -> Result<u32, String> {
    // SAFETY: the documented all-processes flag and a zero process ID are valid.
    let snapshot = unsafe { CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0) };
    if snapshot == INVALID_HANDLE_VALUE {
        return Err("CreateToolhelp32Snapshot failed".to_owned());
    }
    let mut entry = PROCESSENTRY32W {
        dwSize: std::mem::size_of::<PROCESSENTRY32W>() as u32,
        ..Default::default()
    };
    let mut count = 0;
    // SAFETY: snapshot is valid and entry has the required initialized size.
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

fn range(
    worksheet: &excel_com::Worksheet,
    address: &str,
) -> Result<excel_com::Range, excel_com::ExcelComError> {
    worksheet.range(address)
}

#[test]
#[ignore = "launches a fresh visible Excel process; run explicitly"]
fn collection_identity_and_navigation_naturally_exits() -> Result<(), Box<dyn std::error::Error>> {
    assert_eq!(
        excel_process_count()?,
        0,
        "live test requires no pre-existing EXCEL.EXE"
    );
    let apartment = ComApartment::sta()?;
    let application = Application::new(&apartment)?;
    application.set_visible(true)?;
    let outcome = std::panic::catch_unwind(std::panic::AssertUnwindSafe(
        || -> Result<(), Box<dyn std::error::Error>> {
            let alerts = application.display_alerts_guard(false)?;
            let workbooks = application.workbooks()?;
            let workbook_one = workbooks.add()?;
            let workbook_two = workbooks.add()?;
            let workbook_one_by_index = workbooks.item_by_index(1)?;
            assert!(workbook_one.is_same_object(&workbook_one_by_index)?);
            let workbook_one_by_name = workbooks.item_by_name(&workbook_one.name()?)?;
            assert!(workbook_one.is_same_object(&workbook_one_by_name)?);
            let workbook_names = workbooks
                .iter()?
                .map(|item| item?.name())
                .collect::<Result<Vec<_>, _>>()?;
            assert!(
                workbook_names.len() >= 2
                    && workbook_names.windows(2).all(|pair| pair[0] != pair[1])
            );

            let sheets = workbook_one.worksheets()?;
            let sheet_one = sheets.item_by_index(1)?;
            sheet_one.set_name("Prompt10One")?;
            let sheet_two = sheets.add(&Default::default())?;
            sheet_two.set_name("Prompt10Two")?;
            let sheet_three = sheets.add(&Default::default())?;
            sheet_three.set_name("Prompt10Three")?;
            assert!(sheet_two.is_same_object(&sheets.item_by_index(2)?)?);
            assert!(sheet_three.is_same_object(&sheets.item_by_name("Prompt10Three")?)?);
            let indexed_sheet_names = (1..=3)
                .map(|index| sheets.item_by_index(index)?.name())
                .collect::<Result<Vec<_>, _>>()?;
            let sheet_names = sheets
                .iter()?
                .map(|item| item?.name())
                .collect::<Result<Vec<_>, _>>()?;
            assert_eq!(sheet_names, indexed_sheet_names);
            assert_eq!(sheet_names.len(), 3);
            assert!(sheet_names.windows(2).all(|pair| pair[0] != pair[1]));
            let first_iterated_sheet = sheets.iter()?.next().transpose()?.expect("first worksheet");
            assert!(
                sheets
                    .item_by_index(1)?
                    .is_same_object(&first_iterated_sheet)?
            );
            assert!(
                sheets
                    .item_by_name(&first_iterated_sheet.name()?)?
                    .is_same_object(&first_iterated_sheet)?
            );
            let mut early = sheets.iter()?;
            assert!(early.next().transpose()?.is_some());
            drop(early);

            let block = range(&sheet_one, "B2:D4")?;
            assert!(block.cell(0, 1).is_err());
            assert!(block.item(1, Some(0)).is_err());
            let cells = block.cells()?;
            let b2 = cells.cell(1, 1)?;
            assert_eq!(b2.address()?, "$B$2");
            assert!(b2.is_same_object(&b2.clone())?);
            assert_eq!(block.offset(1, 1)?.address()?, "$C$3:$E$5");
            assert_eq!(block.offset(-1, -1)?.address()?, "$A$1:$C$3");
            let resized = block.resize(2, 2)?;
            assert_eq!(resized.address()?, "$B$2:$C$3");
            resized.set_value2(AutomationValue::Array(AutomationArray::from_rows(vec![
                vec![AutomationValue::Number(10.0), AutomationValue::Number(11.0)],
                vec![AutomationValue::Number(12.0), AutomationValue::Number(13.0)],
            ])?))?;
            assert_eq!(block.cell(2, 2)?.value2()?, AutomationValue::Number(13.0));
            assert_eq!(block.rows()?.address()?, "$B$2:$D$4");
            assert_eq!(block.columns()?.address()?, "$B$2:$D$4");
            assert_eq!(block.entire_row()?.address()?, "$2:$4");
            assert_eq!(block.entire_column()?.address()?, "$B:$D");

            let union = application.union2(&b2, &block.cell(3, 3)?)?;
            let areas = union.areas()?;
            assert_eq!(areas.count()?, 2);
            let first = areas.item(1)?;
            let iterated = areas.iter()?.collect::<Result<Vec<_>, _>>()?;
            assert_eq!(iterated.len(), 2);
            let area_item_iterator_same = first.is_same_object(&iterated[0])?;
            eprintln!(
                "Areas.Item(1) versus Areas iterator first identity: {area_item_iterator_same}"
            );
            assert_eq!(iterated[0].address()?, "$B$2");
            assert_eq!(iterated[1].address()?, "$D$4");

            drop((
                iterated,
                first,
                areas,
                union,
                resized,
                b2,
                cells,
                block,
                sheet_three,
                sheet_two,
                sheet_one,
                sheets,
                workbook_one_by_name,
                workbook_one_by_index,
            ));
            workbook_two.close(WorkbookCloseOptions {
                save_changes: SaveChanges::Discard,
                ..WorkbookCloseOptions::new()
            })?;
            workbook_one.close(WorkbookCloseOptions {
                save_changes: SaveChanges::Discard,
                ..WorkbookCloseOptions::new()
            })?;
            drop(workbooks);
            alerts.restore()?;
            Ok(())
        },
    ));
    let quit = application.quit();
    let deadline = Instant::now() + Duration::from_secs(20);
    while Instant::now() < deadline && excel_process_count()? != 0 {
        thread::sleep(Duration::from_millis(100));
    }
    let exited = excel_process_count()? == 0;
    quit?;
    assert!(exited, "owned Excel process did not exit naturally");
    match outcome {
        Ok(result) => result,
        Err(payload) => std::panic::resume_unwind(payload),
    }
}
