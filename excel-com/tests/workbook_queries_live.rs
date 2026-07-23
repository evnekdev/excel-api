#![cfg(windows)]

//! WorkbookQuery inspection coverage for a reviewed, local Power Query fixture.

use excel_com::{ComApartment, OwnedApplication};

#[test]
#[ignore = "fixture-blocked and runtime-blocked: no reviewed local Power Query workbook can be opened while Workbooks.Add returns 0x800A03EC"]
fn workbook_queries_are_inspected_without_formula_logging() -> Result<(), Box<dyn std::error::Error>>
{
    let apartment = ComApartment::sta()?;
    let application = OwnedApplication::new(&apartment)?;
    let mut workbook = None;
    let outcome = (|| -> Result<(), Box<dyn std::error::Error>> {
        let value = application.workbooks()?.add()?;
        workbook = Some(value.clone());
        let queries = value.queries()?;
        for query in queries.iter()? {
            let query = query?;
            let formula = query.formula()?;
            assert!(!format!("{formula:?}").contains(query.name()?.as_str()));
        }
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
