#![cfg(windows)]

//! Opt-in presentation and worksheet-lifecycle acceptance coverage.

use excel_com::{
    ComApartment, OwnedApplication, PageFit, PageOrientation, PageZoom, SheetDestination,
    SheetVisibility, SummaryRow, WorkbookProtectOptions, WorksheetAddOptions,
    WorksheetProtectOptions,
};

#[test]
#[ignore = "requires an Excel host that accepts Workbooks.Add; baseline is currently environment-blocked"]
fn workbook_presentation_lifecycle_and_layout_live() -> Result<(), Box<dyn std::error::Error>> {
    let apartment = ComApartment::sta()?;
    let application = OwnedApplication::new(&apartment)?;
    application.set_visible(true)?;
    let workbook = application.workbooks()?.add()?;
    let worksheets = workbook.worksheets()?;
    let first = worksheets.item_by_index(1)?;
    first.set_name("Presentation")?;
    let second = worksheets.add(&WorksheetAddOptions {
        after: Some(&first),
        ..Default::default()
    })?;
    second.set_name("Layout")?;
    second.move_to(SheetDestination::After(&first))?;
    second.copy_to(SheetDestination::After(&first))?;
    first.set_visible(SheetVisibility::VISIBLE)?;
    first
        .tab()?
        .set_color(excel_com::ExcelColor::from_rgb(32, 96, 160))?;
    let setup = first.page_setup()?;
    setup.set_orientation(PageOrientation::LANDSCAPE)?;
    setup.set_zoom(PageZoom::Automatic)?;
    setup.set_fit_to_pages(PageFit {
        wide: Some(1),
        tall: Some(1),
    })?;
    setup.set_center_header("&BPresentation")?;
    setup.set_center_footer("Page &P of &N")?;
    first.outline()?.set_summary_row(SummaryRow::ABOVE)?;
    first.protect(&WorksheetProtectOptions {
        contents: Some(true),
        ..Default::default()
    })?;
    first.unprotect(None)?;
    workbook.protect(&WorkbookProtectOptions {
        structure: Some(true),
        ..Default::default()
    })?;
    workbook.unprotect(None)?;
    workbook.close_without_saving()?;
    application.quit()?;
    Ok(())
}
