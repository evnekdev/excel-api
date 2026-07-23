#![cfg(windows)]

//! Opt-in, visible Styles, legacy Notes, and internal-hyperlink acceptance coverage.

use excel_com::{Application, ComApartment, HyperlinkAddOptions};

#[test]
#[ignore = "launches a fresh visible Excel process; run explicitly"]
fn styles_comments_and_links_live() -> Result<(), Box<dyn std::error::Error>> {
    let apartment = ComApartment::sta()?;
    let application = Application::new(&apartment)?;
    application.set_visible(true)?;
    let outcome = match application.workbooks()?.add() {
        Ok(workbook) => {
            let work = (|| -> Result<(), Box<dyn std::error::Error>> {
                let worksheet = workbook.worksheets()?.item_by_index(1)?;
                let cell = worksheet.range("A1")?;
                let styles = workbook.styles()?;
                let style = styles.add("CodexPrompt16Style", None)?;
                style.set_number_format("0.00")?;
                cell.set_style_by_name("CodexPrompt16Style")?;
                assert_eq!(
                    cell.style_name()?,
                    excel_com::MixedValue::Uniform("CodexPrompt16Style".to_owned())
                );
                let note = cell.add_comment(Some("legacy Note"))?;
                assert_eq!(note.text()?, "legacy Note");
                note.set_visible(true)?;
                assert!(note.visible()?);
                let link_options = HyperlinkAddOptions {
                    anchor: &cell,
                    address: None,
                    sub_address: Some("A1"),
                    screen_tip: None,
                    text_to_display: Some("jump"),
                };
                let link = cell.hyperlinks()?.add(&link_options)?;
                assert_eq!(link.sub_address()?.as_deref(), Some("A1"));
                link.delete()?;
                cell.comment()?.ok_or("expected legacy Note")?.delete()?;
                style.delete()?;
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
