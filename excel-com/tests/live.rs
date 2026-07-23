#![cfg(windows)]

use std::thread;
use std::time::{Duration, Instant};

use excel_com::{ComApartment, OwnedApplication};
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

#[test]
#[ignore = "launches a fresh visible Excel process; run explicitly"]
fn application_workbooks_workbook_live_slice_naturally_exits()
-> Result<(), Box<dyn std::error::Error>> {
    if excel_process_count()? != 0 {
        return Err("live test requires zero pre-existing EXCEL.EXE processes".into());
    }
    let apartment = ComApartment::sta()?;
    let application = OwnedApplication::new(&apartment)?;
    assert!(!application.version()?.is_empty());
    application.set_visible(true)?;
    assert!(application.visible()?);
    let workbooks = application.workbooks()?;
    let before = workbooks.count()?;
    let workbook = workbooks.add()?;
    assert!(!workbook.name()?.is_empty());
    workbook.set_saved(true)?;
    assert!(workbook.saved()?);
    workbook.close_without_saving()?;
    drop(workbooks);
    application.quit()?;
    let deadline = Instant::now() + Duration::from_secs(15);
    while Instant::now() < deadline && excel_process_count()? != 0 {
        thread::sleep(Duration::from_millis(100));
    }
    assert_eq!(
        excel_process_count()?,
        0,
        "owned Excel process did not exit naturally; initial Workbooks.Count was {before}"
    );
    Ok(())
}
