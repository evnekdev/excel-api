#![cfg(windows)]

//! Opt-in cell-bound Sparkline group coverage.

use excel_com::{Application, AutomationArray, AutomationValue, ComApartment, SparklineType};

#[test]
#[ignore = "launches a fresh visible Excel process; run explicitly with one test thread"]
fn sparklines_live() -> Result<(), Box<dyn std::error::Error>> {
    let apartment = ComApartment::sta()?;
    let application = Application::new(&apartment)?;
    application.set_visible(true)?;
    let result = (|| -> Result<(), Box<dyn std::error::Error>> {
        let workbook = application.workbooks()?.add()?;
        (|| -> Result<(), Box<dyn std::error::Error>> {
            let worksheet = workbook.worksheets()?.item_by_index(1)?;
            worksheet
                .range("A1:D2")?
                .set_value2(AutomationValue::Array(AutomationArray::from_rows(vec![
                    vec![
                        AutomationValue::Number(1.0),
                        AutomationValue::Number(2.0),
                        AutomationValue::Number(3.0),
                        AutomationValue::Number(4.0),
                    ],
                    vec![
                        AutomationValue::Number(4.0),
                        AutomationValue::Number(3.0),
                        AutomationValue::Number(2.0),
                        AutomationValue::Number(1.0),
                    ],
                ])?))?;
            let source = worksheet.range("A1:D2")?;
            let location = worksheet.range("E1:E2")?;
            let group =
                location
                    .sparkline_groups()?
                    .add(SparklineType::LINE, &source, &location)?;
            assert_eq!(group.sparkline_type()?, SparklineType::LINE);
            group.delete()?;
            workbook.close_without_saving()?;
            Ok(())
        })()
    })();
    let quit = application.quit();
    if let Err(error) = quit {
        if result.is_ok() {
            return Err(Box::new(error));
        }
    }
    result
}
