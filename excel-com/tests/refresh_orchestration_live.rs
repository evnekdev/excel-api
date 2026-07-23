#![cfg(windows)]

//! Controlled refresh orchestration coverage for a functioning local Excel host.

use std::time::Duration;

use excel_com::{ComApartment, OwnedApplication, RefreshWaitOptions};

#[test]
#[ignore = "runtime-blocked: Prompt 19 baseline Workbooks.Add returned 0x800A03EC"]
fn workbook_refresh_operations_are_bounded_and_best_effort()
-> Result<(), Box<dyn std::error::Error>> {
    let apartment = ComApartment::sta()?;
    let application = OwnedApplication::new(&apartment)?;
    let mut workbook = None;
    let outcome = (|| -> Result<(), Box<dyn std::error::Error>> {
        let value = application.workbooks()?.add()?;
        workbook = Some(value.clone());
        value.refresh_all()?;
        let _ = value.is_refreshing()?;
        let report = value.wait_for_refresh(RefreshWaitOptions {
            timeout: Duration::from_secs(5),
            poll_interval: Duration::from_millis(50),
        })?;
        assert!(report.completed || report.remaining_queries > 0);
        let _ = value.cancel_all_refreshes()?;
        application.calculate_until_async_queries_done()?;
        value.close_without_saving()?;
        workbook = None;
        Ok(())
    })();
    if let Some(value) = workbook.take() {
        let _ = value.close_without_saving();
    }
    application.quit()?;
    outcome
}
