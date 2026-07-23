//! Create a table over an existing Range in a copied workbook.

use excel_com::tables::{ListObjectAddOptions, TableHeaderMode};
use excel_com::{ComApartment, OwnedApplication, WorkbookOpenOptions};
use std::env;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let path = env::args().nth(1).expect("pass a copied .xlsx path");
    let apartment = ComApartment::sta()?;
    let excel = OwnedApplication::new(&apartment)?;
    let workbook = excel.workbooks()?.open(path, WorkbookOpenOptions::new())?;
    let sheet = workbook.worksheets()?.item_by_index(1)?;
    let source = sheet.range("A1:C10")?;
    let table = sheet
        .list_objects()?
        .add_from_range(&ListObjectAddOptions {
            source: &source,
            has_headers: TableHeaderMode::YES,
            destination: None,
            table_style_name: Some("TableStyleMedium2"),
        })?;
    println!("created table {}", table.name()?);
    drop(table);
    drop(source);
    drop(sheet);
    workbook.close_without_saving()?;
    excel.quit()?;
    Ok(())
}
