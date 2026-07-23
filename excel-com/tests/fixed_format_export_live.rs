#![cfg(windows)]

//! Opt-in fixed-format export acceptance coverage.

use std::path::PathBuf;

use excel_com::{ComApartment, FixedFormatOptions, FixedFormatType, OwnedApplication};

#[test]
#[ignore = "requires an Excel host that accepts Workbooks.Add; baseline is currently environment-blocked"]
fn fixed_format_pdf_export_live() -> Result<(), Box<dyn std::error::Error>> {
    let output: PathBuf = std::env::temp_dir().join("excel-com-prompt15-fixed-format.pdf");
    let apartment = ComApartment::sta()?;
    let application = OwnedApplication::new(&apartment)?;
    application.set_visible(true)?;
    let workbook = application.workbooks()?.add()?;
    let worksheet = workbook.worksheets()?.item_by_index(1)?;
    worksheet
        .range("A1")?
        .set_value2(excel_com::AutomationValue::Text("Prompt 15".to_owned()))?;
    worksheet.export_as_fixed_format(
        FixedFormatType::PDF,
        &FixedFormatOptions {
            output: Some(&output),
            ..Default::default()
        },
    )?;
    assert!(output.is_file(), "Excel did not create the requested PDF");
    workbook.close_without_saving()?;
    application.quit()?;
    let _ = std::fs::remove_file(output);
    Ok(())
}
