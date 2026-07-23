//! Tier 2 release smoke test using only a copied repository fixture.
//!
//! This deliberately never calls `Workbooks.Add`. It is ignored because a
//! licensed Windows desktop Excel installation is required.

mod support;

use excel_com::drawing::{ChartBounds, ChartCreateOptions, ChartType};
use excel_com::tables::{ListObjectAddOptions, TableHeaderMode};
use excel_com::{
    AutomationValue, ComApartment, ExcelColor, FixedFormatOptions, FixedFormatType,
    HorizontalAlignment, OwnedApplication, SaveChanges, WorkbookCloseOptions, WorkbookOpenOptions,
    WorkbookSaveAsOptions, XlFileFormat,
};
use std::time::Duration;
use support::Fixture;

#[test]
#[ignore = "Tier 2 controlled release smoke test: requires Windows desktop Excel"]
fn copied_fixture_exercises_the_release_smoke_path() -> Result<(), Box<dyn std::error::Error>> {
    let apartment = ComApartment::sta()?;
    let excel = OwnedApplication::new(&apartment)?;
    excel.set_visible(false)?;
    let input = Fixture::BlankXlsx.copy_for_test()?;
    let output = std::env::temp_dir().join(format!(
        "excel-com-release-smoke-{}.xlsx",
        std::process::id()
    ));
    let pdf = output.with_extension("pdf");

    let result = (|| -> Result<(), Box<dyn std::error::Error>> {
        let workbook = excel
            .workbooks()?
            .open(input.path(), WorkbookOpenOptions::new())?;
        let sheet = workbook.worksheets()?.item_by_index(1)?;
        sheet
            .range("A1")?
            .set_value(AutomationValue::Text("Category".into()))?;
        sheet
            .range("B1")?
            .set_value(AutomationValue::Text("Amount".into()))?;
        sheet
            .range("A2")?
            .set_value(AutomationValue::Text("One".into()))?;
        sheet
            .range("B2")?
            .set_value(AutomationValue::Number(21.0))?;
        sheet
            .range("A3")?
            .set_value(AutomationValue::Text("Two".into()))?;
        sheet
            .range("B3")?
            .set_value(AutomationValue::Number(34.0))?;
        sheet.range("B4")?.set_formula("=SUM(B2:B3)")?;
        sheet.calculate()?;

        let header = sheet.range("A1:B1")?;
        header.font()?.set_bold(true)?;
        header
            .interior()?
            .set_color(ExcelColor::from_rgb(220, 230, 241))?;
        header.set_horizontal_alignment(HorizontalAlignment::CENTER)?;
        let source = sheet.range("A1:B3")?;
        let table = sheet
            .list_objects()?
            .add_from_range(&ListObjectAddOptions {
                source: &source,
                has_headers: TableHeaderMode::YES,
                destination: None,
                table_style_name: Some("TableStyleMedium2"),
            })?;
        let chart = sheet.add_chart(&ChartCreateOptions {
            source: &source,
            chart_type: ChartType::COLUMN_CLUSTERED,
            bounds: ChartBounds {
                left: 240.0,
                top: 20.0,
                width: 360.0,
                height: 220.0,
            },
            plot_by: None,
            title: Some("Release smoke"),
            has_legend: Some(false),
        })?;
        workbook.save_as(
            &output,
            WorkbookSaveAsOptions {
                file_format: Some(XlFileFormat::OPEN_XML_WORKBOOK),
                ..WorkbookSaveAsOptions::new()
            },
        )?;
        // Excel may not have a PDF export filter. Absence is documented rather
        // than treated as a fixture or object-model failure.
        let _ = workbook.export_as_fixed_format(
            FixedFormatType::PDF,
            &FixedFormatOptions {
                output: Some(&pdf),
                ..FixedFormatOptions::default()
            },
        );
        drop(chart);
        drop(table);
        drop(source);
        drop(header);
        drop(sheet);
        workbook.close(WorkbookCloseOptions {
            save_changes: SaveChanges::Discard,
            ..WorkbookCloseOptions::new()
        })?;
        Ok(())
    })();

    drop(input);
    let exit = excel.quit_and_wait(Duration::from_secs(30));
    let _ = std::fs::remove_file(&pdf);
    let _ = std::fs::remove_file(&output);
    result?;
    assert!(exit?.exited);
    Ok(())
}
