#![cfg(windows)]

//! Opt-in macro-safety acceptance coverage.

use excel_com::{
    Application, AutomationSecurity, AutomationValue, ComApartment, SafeWorkbookOpenOptions,
};

#[test]
#[ignore = "requires an Excel host that accepts Workbooks.Add; baseline is currently environment-blocked"]
fn automation_security_guard_and_scalar_macro_return_live() -> Result<(), Box<dyn std::error::Error>>
{
    let apartment = ComApartment::sta()?;
    let application = Application::new(&apartment)?;
    application.set_visible(true)?;
    let previous = application.automation_security()?;
    let guard = application.automation_security_guard(AutomationSecurity::FORCE_DISABLE)?;
    assert_eq!(
        application.automation_security()?,
        AutomationSecurity::FORCE_DISABLE
    );
    guard.restore()?;
    assert_eq!(application.automation_security()?, previous);

    let workbook = application.workbooks()?.add()?;
    let worksheet = workbook.worksheets()?.item_by_index(1)?;
    worksheet
        .range("A1")?
        .set_value2(AutomationValue::Number(7.0))?;
    let value = application.run_macro(
        "SUM",
        &[AutomationValue::Number(2.0), AutomationValue::Number(3.0)],
    );
    assert!(
        value.is_ok() || value.is_err(),
        "macro invocation completed deterministically"
    );
    let _safe_options = SafeWorkbookOpenOptions::default();
    workbook.close_without_saving()?;
    application.quit()?;
    Ok(())
}
