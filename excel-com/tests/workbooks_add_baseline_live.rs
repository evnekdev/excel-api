#![cfg(windows)]

use std::thread;
use std::time::{Duration, Instant};

use excel_com::{Application, ComApartment};
use windows_sys::Win32::Foundation::{CloseHandle, INVALID_HANDLE_VALUE};
use windows_sys::Win32::System::Diagnostics::ToolHelp::{
    CreateToolhelp32Snapshot, PROCESSENTRY32W, Process32FirstW, Process32NextW, TH32CS_SNAPPROCESS,
};

fn excel_process_count() -> Result<u32, String> {
    // SAFETY: the documented process-snapshot API accepts these arguments.
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
    let deadline = Instant::now() + Duration::from_secs(30);
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

/// A minimal, independently runnable `Workbooks.Add` baseline.
///
/// When Excel rejects `Add`, the failure contains only normalized Application
/// state needed to distinguish the environmental `0x800A03EC` condition from
/// a failure to activate Excel. It does not change any Application setting.
#[test]
#[ignore = "launches a fresh Excel process; run explicitly with one test thread"]
fn workbooks_add_baseline_naturally_exits() -> Result<(), Box<dyn std::error::Error>> {
    assert_eq!(
        excel_process_count()?,
        0,
        "live test requires no pre-existing EXCEL.EXE"
    );
    let apartment = ComApartment::sta()?;
    let application = Application::new(&apartment)?;
    let state = format!(
        "application_created=true visible={:?} display_alerts={:?} calculation_mode={:?} reference_style={:?}",
        application.visible(),
        application.display_alerts(),
        application.calculation_mode(),
        application.reference_style(),
    );

    let outcome = (|| -> Result<(), Box<dyn std::error::Error>> {
        let workbook = application.workbooks()?.add()?;
        workbook.close_without_saving()?;
        Ok(())
    })();
    let quit = application.quit();
    wait_for_excel_exit()?;
    quit?;
    outcome.map_err(|error| format!("Workbooks.Add baseline failed; {state}; {error}").into())
}
