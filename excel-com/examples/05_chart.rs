//! Create a small embedded chart from an existing Range.

use excel_com::drawing::{ChartBounds, ChartCreateOptions, ChartType};
use excel_com::{ComApartment, OwnedApplication, WorkbookOpenOptions};
use std::env;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let path = env::args().nth(1).expect("pass a copied .xlsx path");
    let apartment = ComApartment::sta()?;
    let excel = OwnedApplication::new(&apartment)?;
    let workbook = excel.workbooks()?.open(path, WorkbookOpenOptions::new())?;
    let sheet = workbook.worksheets()?.item_by_index(1)?;
    let source = sheet.range("A1:B10")?;
    let chart = sheet.add_chart(&ChartCreateOptions {
        source: &source,
        chart_type: ChartType::COLUMN_CLUSTERED,
        bounds: ChartBounds {
            left: 300.0,
            top: 20.0,
            width: 360.0,
            height: 220.0,
        },
        plot_by: None,
        title: Some("Example chart"),
        has_legend: Some(true),
    })?;
    println!("created chart {}", chart.name()?);
    drop(chart);
    drop(source);
    drop(sheet);
    workbook.close_without_saving()?;
    excel.quit()?;
    Ok(())
}
