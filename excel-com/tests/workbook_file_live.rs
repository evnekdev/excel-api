#![cfg(windows)]

use std::fs;
use std::path::Path;
use std::thread;
use std::time::{Duration, Instant, SystemTime, UNIX_EPOCH};

use excel_com::{
    Application, AutomationArgument, AutomationArray, AutomationValue, ComApartment, ExcelComError,
    SaveChanges, WorkbookCloseOptions, WorkbookOpenOptions, WorkbookSaveAsOptions, XlFileFormat,
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

fn range(
    worksheet: &excel_com::Worksheet,
    address: &str,
) -> Result<excel_com::Range, ExcelComError> {
    worksheet.range(
        AutomationArgument::Value(AutomationValue::Text(address.to_owned())),
        None,
    )
}

fn unique_test_directory() -> std::path::PathBuf {
    let stamp = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("system clock before Unix epoch")
        .as_nanos();
    std::env::temp_dir().join(format!("excel-com-workbook-file-live-{stamp}"))
}

#[test]
#[ignore = "launches a fresh visible Excel process and writes only a unique temporary directory"]
fn workbook_file_lifecycle_naturally_exits() -> Result<(), Box<dyn std::error::Error>> {
    assert_eq!(
        excel_process_count()?,
        0,
        "live test requires no pre-existing EXCEL.EXE"
    );
    let directory = unique_test_directory();
    fs::create_dir_all(&directory)?;
    let primary = directory.join("saved workbook.xlsx");
    let copy = directory.join("saved copy.xlsx");

    let apartment = ComApartment::sta()?;
    let application = Application::new(&apartment)?;
    application.set_visible(true)?;
    let outcome = (|| -> Result<(), Box<dyn std::error::Error>> {
        let original_alerts = application.display_alerts()?;
        {
            let guard = application.display_alerts_guard(false)?;
            let invalid = application
                .workbooks()?
                .open_default(Path::new("invalid\0workbook.xlsx"));
            assert!(matches!(invalid, Err(ExcelComError::InvalidPath { .. })));
            guard.restore()?;
        }
        assert_eq!(application.display_alerts()?, original_alerts);

        let workbooks = application.workbooks()?;
        let workbook = workbooks.add()?;
        let worksheets = workbook.worksheets()?;
        let worksheet = worksheets.item_by_index(1)?;
        range(&worksheet, "A1")?.set_value2(AutomationValue::Text("saved value".to_owned()))?;
        range(&worksheet, "B1")?
            .set_formula2(AutomationValue::Text("=SEQUENCE(2,2)".to_owned()))?;
        workbook.save_as(
            &primary,
            WorkbookSaveAsOptions {
                file_format: Some(XlFileFormat::OPEN_XML_WORKBOOK),
                ..WorkbookSaveAsOptions::new()
            },
        )?;
        assert_eq!(workbook.file_format()?, XlFileFormat::OPEN_XML_WORKBOOK);
        assert!(!workbook.read_only()?);
        assert!(workbook.full_name()?.ends_with("saved workbook.xlsx"));
        assert!(workbook.path()?.contains("excel-com-workbook-file-live-"));
        range(&worksheet, "A2")?.set_value2(AutomationValue::Number(42.0))?;
        workbook.save()?;
        let identity_before_copy = workbook.full_name()?;
        workbook.save_copy_as(&copy)?;
        assert!(copy.exists());
        assert_eq!(workbook.full_name()?, identity_before_copy);

        drop((worksheet, worksheets));
        workbook.close(WorkbookCloseOptions {
            save_changes: SaveChanges::Discard,
            ..WorkbookCloseOptions::new()
        })?;

        let reopened = workbooks.open(
            &primary,
            WorkbookOpenOptions {
                read_only: Some(true),
                ..WorkbookOpenOptions::new()
            },
        )?;
        assert!(reopened.read_only()?);
        let reopened_sheets = reopened.worksheets()?;
        let reopened_sheet = reopened_sheets.item_by_index(1)?;
        assert_eq!(
            range(&reopened_sheet, "A1")?.value2()?,
            AutomationValue::Text("saved value".to_owned())
        );
        assert_eq!(
            range(&reopened_sheet, "B1:C2")?.value2()?,
            AutomationValue::Array(AutomationArray::from_rows(vec![
                vec![AutomationValue::Number(1.0), AutomationValue::Number(2.0)],
                vec![AutomationValue::Number(3.0), AutomationValue::Number(4.0)],
            ])?)
        );
        match reopened.save() {
            Ok(()) => eprintln!("read-only Workbook.Save completed without a reported error"),
            Err(error) => eprintln!("read-only Workbook.Save reported: {error}"),
        }
        drop((reopened_sheet, reopened_sheets));
        reopened.close(WorkbookCloseOptions {
            save_changes: SaveChanges::Discard,
            ..WorkbookCloseOptions::new()
        })?;
        drop(workbooks);
        Ok(())
    })();
    let quit = application.quit();
    let deadline = Instant::now() + Duration::from_secs(20);
    while Instant::now() < deadline && excel_process_count()? != 0 {
        thread::sleep(Duration::from_millis(100));
    }
    let exited = excel_process_count()? == 0;
    fs::remove_dir_all(&directory)?;
    outcome?;
    quit?;
    assert!(exited, "owned Excel process did not exit naturally");
    Ok(())
}
