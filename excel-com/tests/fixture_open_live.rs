//! Controlled fixture-open fallback. This test never calls `Workbooks.Add`.

mod support;

use excel_com::{
    ComApartment, ComRetryPolicy, OwnedApplication, SaveChanges, WorkbookCloseOptions,
    WorkbookOpenOptions,
};
use std::time::Duration;
use support::Fixture;

#[test]
#[ignore = "controlled live test: opens only a copied repository fixture"]
fn owned_session_opens_a_copied_blank_fixture_and_exits_naturally()
-> Result<(), Box<dyn std::error::Error>> {
    let apartment = ComApartment::sta()?;
    let application =
        OwnedApplication::new_with_retry_policy(&apartment, ComRetryPolicy::default())?;
    application.set_visible(false)?;
    let temporary = Fixture::BlankXlsx.copy_for_test()?;
    let workbooks = application.workbooks()?;
    let workbook = workbooks.open(temporary.path(), WorkbookOpenOptions::new())?;
    workbook.close(WorkbookCloseOptions {
        save_changes: SaveChanges::Discard,
        ..WorkbookCloseOptions::new()
    })?;
    drop(workbooks);
    drop(temporary);
    let report = application.quit_and_wait(Duration::from_secs(30))?;
    assert!(report.exited);
    Ok(())
}
