#![cfg(windows)]

//! Controlled local-file QueryTable coverage for a host that accepts `Workbooks.Add`.

use std::fs;
use std::time::{SystemTime, UNIX_EPOCH};

use excel_com::{
    ComApartment, OwnedApplication, TextDelimiter, TextParsingType, TextQualifier,
    TextQueryAddOptions,
};

#[test]
#[ignore = "runtime-blocked: Prompt 19 baseline Workbooks.Add returned 0x800A03EC"]
fn local_text_querytable_is_owned_refreshable_and_removable()
-> Result<(), Box<dyn std::error::Error>> {
    let csv_path = std::env::temp_dir().join(format!(
        "excel-com-prompt19-{}-{}.csv",
        std::process::id(),
        SystemTime::now().duration_since(UNIX_EPOCH)?.as_nanos()
    ));
    fs::write(&csv_path, "Region,Amount\nNorth,10\nSouth,20\n")?;

    let apartment = ComApartment::sta()?;
    let application = OwnedApplication::new(&apartment)?;
    let mut workbook = None;
    let outcome = (|| -> Result<(), Box<dyn std::error::Error>> {
        let value = application.workbooks()?.add()?;
        workbook = Some(value.clone());
        let worksheet = value.worksheets()?.item_by_index(1)?;
        let query = worksheet
            .query_tables()?
            .add_from_local_text(&TextQueryAddOptions {
                path: &csv_path,
                destination: &worksheet.range("A1")?,
                parsing_type: TextParsingType::DELIMITED,
                delimiter: Some(TextDelimiter::Comma),
                text_qualifier: Some(TextQualifier::DOUBLE_QUOTE),
                columns: vec![],
                refresh_on_file_open: false,
                background_query: false,
            })?;
        assert_eq!(worksheet.query_tables()?.count()?, 1);
        let _refresh_result = query.refresh(Some(false))?;
        assert!(!query.refreshing()?);
        assert!(query.result_range()?.is_some());
        query.delete()?;
        value.close_without_saving()?;
        workbook = None;
        Ok(())
    })();
    if let Some(value) = workbook.take() {
        let _ = value.close_without_saving();
    }
    let quit = application.quit();
    let _ = fs::remove_file(&csv_path);
    quit?;
    outcome
}
