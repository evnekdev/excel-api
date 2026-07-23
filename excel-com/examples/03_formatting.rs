//! Apply explicit, non-mixed formatting to a Range.

use excel_com::{
    ComApartment, ExcelColor, HorizontalAlignment, OwnedApplication, WorkbookOpenOptions,
};
use std::env;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let path = env::args().nth(1).expect("pass a copied .xlsx path");
    let apartment = ComApartment::sta()?;
    let excel = OwnedApplication::new(&apartment)?;
    let workbook = excel.workbooks()?.open(path, WorkbookOpenOptions::new())?;
    let sheet = workbook.worksheets()?.item_by_index(1)?;
    let range = sheet.range("A1:C1")?;
    range.font()?.set_bold(true)?;
    range.font()?.set_color(ExcelColor::from_rgb(20, 40, 180))?;
    range.set_horizontal_alignment(HorizontalAlignment::CENTER)?;
    drop(range);
    drop(sheet);
    workbook.close_without_saving()?;
    excel.quit()?;
    Ok(())
}
