#![cfg(windows)]

//! Opt-in embedded-chart and chart-sheet coverage using a fresh owned workbook.

use std::fs;

use excel_com::{
    AutomationArray, AutomationValue, AxisGroup, AxisType, ChartBounds, ChartCreateOptions,
    ChartExportOptions, ChartType, ComApartment, CopyPictureFormat, CopyPictureOptions,
    LegendPosition, MarkerStyle, OwnedApplication, PictureAppearance, PlotBy,
};

#[test]
#[ignore = "launches a fresh visible Excel process; run explicitly with one test thread"]
fn charts_live() -> Result<(), Box<dyn std::error::Error>> {
    let apartment = ComApartment::sta()?;
    let application = OwnedApplication::new(&apartment)?;
    application.set_visible(true)?;
    let mut exported = None;
    let result = (|| -> Result<(), Box<dyn std::error::Error>> {
        let workbook = application.workbooks()?.add()?;
        (|| -> Result<(), Box<dyn std::error::Error>> {
            let worksheet = workbook.worksheets()?.item_by_index(1)?;
            worksheet
                .range("A1:B5")?
                .set_value2(AutomationValue::Array(AutomationArray::from_rows(vec![
                    vec![
                        AutomationValue::Text("Category".to_owned()),
                        AutomationValue::Text("Value".to_owned()),
                    ],
                    vec![
                        AutomationValue::Text("A".to_owned()),
                        AutomationValue::Number(1.0),
                    ],
                    vec![
                        AutomationValue::Text("B".to_owned()),
                        AutomationValue::Number(2.0),
                    ],
                    vec![
                        AutomationValue::Text("C".to_owned()),
                        AutomationValue::Number(3.0),
                    ],
                    vec![
                        AutomationValue::Text("D".to_owned()),
                        AutomationValue::Number(4.0),
                    ],
                ])?))?;
            let source = worksheet.range("A1:B5")?;
            let chart_object = worksheet.add_chart(&ChartCreateOptions {
                source: &source,
                chart_type: ChartType::COLUMN_CLUSTERED,
                bounds: ChartBounds {
                    left: 240.0,
                    top: 20.0,
                    width: 360.0,
                    height: 220.0,
                },
                plot_by: Some(PlotBy::COLUMNS),
                title: Some("Prompt 17 chart"),
                has_legend: Some(true),
            })?;
            chart_object.set_name("Prompt17Chart")?;
            assert_eq!(worksheet.chart_objects()?.count()?, 1);
            assert!(chart_object.width()? > 0.0);
            let chart = chart_object.chart()?;
            assert_eq!(chart.chart_type()?, ChartType::COLUMN_CLUSTERED);
            chart
                .legend()?
                .expect("legend")
                .set_position(LegendPosition::BOTTOM)?;
            let series = chart.series_collection()?.item(1)?;
            series.set_marker_style(MarkerStyle::CIRCLE)?;
            series.set_marker_size(6)?;
            assert!(
                chart
                    .axes()?
                    .item(AxisType::VALUE, Some(AxisGroup::PRIMARY))?
                    .is_some()
            );
            let output = std::env::temp_dir().join("excel-com-prompt17-chart.png");
            let exported_ok = chart.export(&ChartExportOptions {
                path: &output,
                filter_name: Some("PNG"),
                interactive: Some(false),
            })?;
            assert!(exported_ok);
            assert!(fs::metadata(&output)?.len() > 0);
            exported = Some(output);
            chart.copy_picture(&CopyPictureOptions {
                appearance: PictureAppearance::SCREEN,
                format: CopyPictureFormat::BITMAP,
            })?;
            application.clear_cut_copy_mode()?;
            chart_object.delete()?;
            workbook.close_without_saving()?;
            Ok(())
        })()
    })();
    if let Some(path) = exported.take() {
        let _ = fs::remove_file(path);
    }
    let quit = application.quit();
    if let Err(error) = quit {
        if result.is_ok() {
            return Err(Box::new(error));
        }
    }
    result
}
