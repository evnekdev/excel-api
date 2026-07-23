//! Read and write values and formulas in a copied workbook.

use excel_com::{AutomationValue, ComApartment, OwnedApplication, WorkbookOpenOptions};
use std::env;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let path = env::args().nth(1).expect("pass a copied .xlsx path");
    let apartment = ComApartment::sta()?;
    let excel = OwnedApplication::new(&apartment)?;
    let workbook = excel.workbooks()?.open(path, WorkbookOpenOptions::new())?;
    let worksheet = workbook.worksheets()?.item_by_index(1)?;
    worksheet
        .range("A1")?
        .set_value(AutomationValue::Number(21.0))?;
    worksheet.range("B1")?.set_formula("=A1*2")?;
    worksheet.calculate()?;
    println!("B1 = {:?}", worksheet.range("B1")?.value2()?);
    drop(worksheet);
    workbook.close_without_saving()?;
    excel.quit()?;
    Ok(())
}
