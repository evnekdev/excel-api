#![cfg(windows)]

//! Range-backed PivotCache and PivotTable coverage for a functioning local Excel host.

use excel_com::{
    AggregationFunction, Application, AutomationArray, AutomationValue, ComApartment,
    PivotDataField, PivotFieldOrientation, PivotFieldPlacement, PivotLayoutOptions,
    PivotTableCreateOptions,
};

#[test]
#[ignore = "runtime-blocked: Prompt 19 baseline Workbooks.Add returned 0x800A03EC"]
fn range_cache_pivot_layout_refreshes_and_cleans_up() -> Result<(), Box<dyn std::error::Error>> {
    let apartment = ComApartment::sta()?;
    let application = Application::new(&apartment)?;
    let mut workbook = None;
    let outcome = (|| -> Result<(), Box<dyn std::error::Error>> {
        let value = application.workbooks()?.add()?;
        workbook = Some(value.clone());
        let worksheet = value.worksheets()?.item_by_index(1)?;
        worksheet
            .range("A1:B3")?
            .set_value2(AutomationValue::Array(AutomationArray::from_rows(vec![
                vec![
                    AutomationValue::Text("Region".to_owned()),
                    AutomationValue::Text("Amount".to_owned()),
                ],
                vec![
                    AutomationValue::Text("North".to_owned()),
                    AutomationValue::Number(10.0),
                ],
                vec![
                    AutomationValue::Text("South".to_owned()),
                    AutomationValue::Number(20.0),
                ],
            ])?))?;
        let cache = value
            .pivot_caches()?
            .create_from_range(&worksheet.range("A1:B3")?, None)?;
        let pivot = cache.create_pivot_table(&PivotTableCreateOptions {
            destination: &worksheet.range("D1")?,
            name: "Prompt19Pivot",
            version: None,
            read_data: Some(true),
            default_version: None,
        })?;
        pivot.apply_layout(&PivotLayoutOptions {
            fields: vec![PivotFieldPlacement {
                field_name: "Region",
                orientation: PivotFieldOrientation::ROW,
                position: Some(1),
            }],
            data_fields: vec![PivotDataField {
                field_name: "Amount",
                caption: Some("Total Amount"),
                function: AggregationFunction::SUM,
                number_format: Some("#,##0.00"),
            }],
        })?;
        let _refresh_result = pivot.refresh_table()?;
        pivot.clear_table()?;
        value.close_without_saving()?;
        workbook = None;
        Ok(())
    })();
    if let Some(value) = workbook.take() {
        let _ = value.close_without_saving();
    }
    application.quit()?;
    outcome
}
