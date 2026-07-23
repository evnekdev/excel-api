#![cfg(windows)]

//! Opt-in Shape, picture, text-box, and Excel clipboard-state coverage.

use std::fs;

use excel_com::{
    Application, AutoShapeType, ComApartment, CopyPictureFormat, CopyPictureOptions,
    PictureAddOptions, PictureAppearance, ShapeBounds, ShapePoint, ShapeType, TextBoxAddOptions,
    TextOrientation, ZOrderCommand,
};

/// A reviewed, repository-owned one-pixel PNG used only to exercise local insertion.
const OWNED_PNG_FIXTURE: &[u8] = &[
    137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 1, 0, 0, 0, 1, 8, 6, 0,
    0, 0, 31, 21, 196, 137, 0, 0, 0, 13, 73, 68, 65, 84, 8, 215, 99, 248, 207, 192, 240, 31, 0, 5,
    0, 1, 255, 137, 153, 61, 29, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130,
];

#[test]
#[ignore = "launches a fresh visible Excel process; run explicitly with one test thread"]
fn shapes_images_live() -> Result<(), Box<dyn std::error::Error>> {
    let apartment = ComApartment::sta()?;
    let application = Application::new(&apartment)?;
    application.set_visible(true)?;
    let fixture = std::env::temp_dir().join("excel-com-prompt17-owned-fixture.png");
    fs::write(&fixture, OWNED_PNG_FIXTURE)?;
    let result = (|| -> Result<(), Box<dyn std::error::Error>> {
        let workbook = application.workbooks()?.add()?;
        (|| -> Result<(), Box<dyn std::error::Error>> {
            let worksheet = workbook.worksheets()?.item_by_index(1)?;
            let shapes = worksheet.shapes()?;
            let rectangle = shapes.add_shape(
                AutoShapeType::RECTANGLE,
                ShapeBounds {
                    left: 20.0,
                    top: 20.0,
                    width: 160.0,
                    height: 60.0,
                },
            )?;
            rectangle.set_name("Prompt17Rectangle")?;
            rectangle.set_rotation(10.0)?;
            rectangle.z_order(ZOrderCommand::BRING_TO_FRONT)?;
            let line = shapes.add_line(
                ShapePoint { x: 20.0, y: 100.0 },
                ShapePoint { x: 180.0, y: 100.0 },
            )?;
            let text = shapes.add_text_box(&TextBoxAddOptions {
                orientation: TextOrientation::HORIZONTAL,
                bounds: ShapeBounds {
                    left: 20.0,
                    top: 120.0,
                    width: 180.0,
                    height: 40.0,
                },
                text: "Prompt 17",
            })?;
            let picture = shapes.add_picture(&PictureAddOptions {
                path: &fixture,
                link_to_file: false,
                save_with_document: true,
                bounds: ShapeBounds {
                    left: 220.0,
                    top: 20.0,
                    width: 32.0,
                    height: 32.0,
                },
            })?;
            assert_eq!(picture.shape_type()?, ShapeType::PICTURE);
            assert_eq!(shapes.count()?, 4);
            worksheet
                .range("A1:B2")?
                .copy_picture(&CopyPictureOptions {
                    appearance: PictureAppearance::SCREEN,
                    format: CopyPictureFormat::BITMAP,
                })?;
            application.clear_cut_copy_mode()?;
            picture.delete()?;
            text.delete()?;
            line.delete()?;
            rectangle.delete()?;
            workbook.close_without_saving()?;
            Ok(())
        })()
    })();
    let _ = fs::remove_file(&fixture);
    let quit = application.quit();
    if let Err(error) = quit {
        if result.is_ok() {
            return Err(Box::new(error));
        }
    }
    result
}
