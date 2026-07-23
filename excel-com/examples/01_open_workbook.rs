//! Open a caller-supplied copied workbook and shut down only the owned session.

use excel_com::{
    ComApartment, OwnedApplication, SaveChanges, WorkbookCloseOptions, WorkbookOpenOptions,
};
use std::{env, time::Duration};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let path = env::args().nth(1).expect("pass a copied .xlsx path");
    let apartment = ComApartment::sta()?;
    let excel = OwnedApplication::new(&apartment)?;
    let workbook = excel.workbooks()?.open(path, WorkbookOpenOptions::new())?;
    println!("opened {}", workbook.name()?);
    workbook.close(WorkbookCloseOptions {
        save_changes: SaveChanges::Discard,
        ..WorkbookCloseOptions::new()
    })?;
    excel.quit_and_wait(Duration::from_secs(30))?;
    Ok(())
}
