#![cfg(windows)]

//! Opt-in, visible conditional-formatting acceptance coverage.

use excel_com::{
    Application, CellValueRuleOptions, ComApartment, ConditionalFormat, ConditionalOperator,
    IconSetKind,
};

#[test]
#[ignore = "launches a fresh visible Excel process; run explicitly"]
fn conditional_formatting_live() -> Result<(), Box<dyn std::error::Error>> {
    let apartment = ComApartment::sta()?;
    let application = Application::new(&apartment)?;
    application.set_visible(true)?;
    let outcome = match application.workbooks()?.add() {
        Ok(workbook) => {
            let work = (|| -> Result<(), Box<dyn std::error::Error>> {
                let worksheet = workbook.worksheets()?.item_by_index(1)?;
                let range = worksheet.range("A1:A5")?;
                range.set_value2(excel_com::AutomationValue::Array(
                    excel_com::AutomationArray::column(vec![
                        excel_com::AutomationValue::Number(1.0),
                        excel_com::AutomationValue::Number(2.0),
                        excel_com::AutomationValue::Number(3.0),
                        excel_com::AutomationValue::Number(4.0),
                        excel_com::AutomationValue::Number(5.0),
                    ])?,
                ))?;
                let conditions = range.format_conditions()?;
                let cell_rule = CellValueRuleOptions {
                    operator: ConditionalOperator::GREATER,
                    formula1: "2",
                    formula2: None,
                };
                let rule = conditions.add_cell_value(&cell_rule)?;
                assert_eq!(rule.condition_type()?.raw(), 1);
                assert_eq!(conditions.count()?, 1);
                let _scale = conditions.add_color_scale(3)?;
                let _bar = conditions.add_data_bar()?;
                let _icons = conditions.add_icon_set_with_kind(IconSetKind::THREE_ARROWS)?;
                assert!(conditions.count()? >= 4);
                match conditions.item_by_index(1)? {
                    ConditionalFormat::CellValue(value) => assert_eq!(value.priority()?, 1),
                    other => {
                        return Err(
                            format!("unexpected first conditional format: {other:?}").into()
                        );
                    }
                }
                Ok(())
            })();
            let close = workbook.close_without_saving();
            match (work, close) {
                (Err(error), _) => Err(error),
                (Ok(()), Err(error)) => Err(Box::new(error) as Box<dyn std::error::Error>),
                (Ok(()), Ok(())) => Ok(()),
            }
        }
        Err(error) => Err(Box::new(error) as Box<dyn std::error::Error>),
    };
    let quit = application.quit();
    outcome?;
    Ok(quit?)
}
