#![cfg(windows)]

use std::thread;
use std::time::{Duration, Instant};

use excel_com::{
    AutomationArray, AutomationValue, BorderIndex, BorderLineStyle, BorderWeight, ComApartment,
    ExcelColor, FillPattern, HorizontalAlignment, MixedValue, OwnedApplication, UnderlineStyle,
    VerticalAlignment,
};
use windows_sys::Win32::Foundation::{CloseHandle, INVALID_HANDLE_VALUE};
use windows_sys::Win32::System::Diagnostics::ToolHelp::{
    CreateToolhelp32Snapshot, PROCESSENTRY32W, Process32FirstW, Process32NextW, TH32CS_SNAPPROCESS,
};

fn excel_process_count() -> Result<u32, String> {
    // SAFETY: process snapshots accept the documented all-processes flag and zero process ID.
    let snapshot = unsafe { CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0) };
    if snapshot == INVALID_HANDLE_VALUE {
        return Err("CreateToolhelp32Snapshot failed".to_owned());
    }
    let mut entry = PROCESSENTRY32W {
        dwSize: std::mem::size_of::<PROCESSENTRY32W>() as u32,
        ..Default::default()
    };
    let mut count = 0;
    // SAFETY: snapshot is valid and entry has the SDK-required dwSize initialized.
    let mut present = unsafe { Process32FirstW(snapshot, &mut entry) } != 0;
    while present {
        if String::from_utf16_lossy(&entry.szExeFile)
            .trim_end_matches('\0')
            .eq_ignore_ascii_case("EXCEL.EXE")
        {
            count += 1;
        }
        // SAFETY: snapshot and initialized entry remain valid for enumeration.
        present = unsafe { Process32NextW(snapshot, &mut entry) } != 0;
    }
    // SAFETY: this function owns the snapshot and closes it exactly once.
    unsafe { CloseHandle(snapshot) };
    Ok(count)
}

fn wait_for_zero_excel_processes() -> Result<(), String> {
    let deadline = Instant::now() + Duration::from_secs(15);
    while Instant::now() < deadline {
        if excel_process_count()? == 0 {
            return Ok(());
        }
        thread::sleep(Duration::from_millis(100));
    }
    Err(format!(
        "crate-owned Excel process did not exit naturally; count={}",
        excel_process_count()?
    ))
}

fn uniform<T: std::fmt::Debug>(value: MixedValue<T>) -> Result<T, Box<dyn std::error::Error>> {
    match value {
        MixedValue::Uniform(value) => Ok(value),
        value => Err(format!("expected uniform formatting value, received {value:?}").into()),
    }
}

#[test]
#[ignore = "launches a fresh visible Excel process; run explicitly"]
fn range_formatting_naturally_exits() -> Result<(), Box<dyn std::error::Error>> {
    assert_eq!(
        excel_process_count()?,
        0,
        "live test requires no pre-existing EXCEL.EXE"
    );
    let apartment = ComApartment::sta()?;
    let application = OwnedApplication::new(&apartment)?;
    application.set_visible(true)?;
    let mut cleanup_workbook = None;

    let outcome = std::panic::catch_unwind(std::panic::AssertUnwindSafe(
        || -> Result<(), Box<dyn std::error::Error>> {
            let alerts = application.display_alerts_guard(false)?;
            let workbook = application.workbooks()?.add()?;
            cleanup_workbook = Some(workbook.clone());
            let worksheet = workbook.worksheets()?.item_by_index(1)?;
            let range = worksheet.range("A1:C4")?;
            range.set_value2(AutomationValue::Array(AutomationArray::new(
                4,
                3,
                vec![
                    AutomationValue::Text("text".to_owned()),
                    AutomationValue::Number(12.5),
                    AutomationValue::Number(45_000.0),
                    AutomationValue::Number(1.0),
                    AutomationValue::Number(2.0),
                    AutomationValue::Number(3.0),
                    AutomationValue::Number(4.0),
                    AutomationValue::Number(5.0),
                    AutomationValue::Number(6.0),
                    AutomationValue::Number(7.0),
                    AutomationValue::Number(8.0),
                    AutomationValue::Number(9.0),
                ],
            )?))?;
            worksheet.range("C1")?.set_formula("=SUM(A2:B2)")?;
            assert_eq!(uniform(range.number_format()?)?, "General");

            let font = range.font()?;
            font.set_name("Arial")?;
            font.set_size(12.0)?;
            font.set_bold(true)?;
            font.set_italic(true)?;
            font.set_underline(UnderlineStyle::SINGLE)?;
            font.set_strikethrough(false)?;
            for color in [
                ExcelColor::from_rgb(255, 0, 0),
                ExcelColor::from_rgb(0, 255, 0),
                ExcelColor::from_rgb(0, 0, 255),
                ExcelColor::from_rgb(0, 0, 0),
                ExcelColor::from_rgb(255, 255, 255),
            ] {
                font.set_color(color)?;
                assert_eq!(uniform(font.color()?)?, color);
            }
            let color = ExcelColor::from_rgb(12, 34, 56);
            font.set_color(color)?;
            assert_eq!(uniform(font.name()?)?, "Arial");
            assert_eq!(uniform(font.size()?)?, 12.0);
            assert!(uniform(font.bold()?)?);
            assert!(uniform(font.italic()?)?);
            assert_eq!(uniform(font.underline()?)?, UnderlineStyle::SINGLE);
            assert!(!uniform(font.strikethrough()?)?);
            let round_trip = uniform(font.color()?)?;
            assert_eq!(
                (round_trip.red(), round_trip.green(), round_trip.blue()),
                (12, 34, 56)
            );
            eprintln!(
                "font name={:?} size={:?} bold={:?} color_raw={}",
                font.name()?,
                font.size()?,
                font.bold()?,
                round_trip.raw()
            );

            let interior = range.interior()?;
            interior.set_color(ExcelColor::from_rgb(240, 200, 120))?;
            interior.set_pattern(FillPattern::SOLID)?;
            interior.set_pattern_color(ExcelColor::from_rgb(11, 22, 33))?;
            assert_eq!(uniform(interior.pattern()?)?, FillPattern::SOLID);
            assert_eq!(
                uniform(interior.color()?)?,
                ExcelColor::from_rgb(240, 200, 120)
            );
            eprintln!(
                "interior color={:?} pattern={:?} pattern_color={:?}",
                interior.color()?,
                interior.pattern()?,
                interior.pattern_color()?
            );

            range.set_number_format("0.00")?;
            assert_eq!(uniform(range.number_format()?)?, "0.00");
            worksheet.range("B1")?.set_number_format("0%")?;
            worksheet.range("C1")?.set_number_format("yyyy-mm-dd")?;
            worksheet
                .range("A2")?
                .set_number_format("[Green]0.00;[Red]-0.00;0.00;@")?;
            eprintln!(
                "number formats A1={:?} B1={:?} C1={:?}",
                worksheet.range("A1")?.number_format()?,
                worksheet.range("B1")?.number_format()?,
                worksheet.range("C1")?.number_format()?
            );

            range.set_horizontal_alignment(HorizontalAlignment::CENTER)?;
            range.set_vertical_alignment(VerticalAlignment::CENTER)?;
            range.set_wrap_text(true)?;
            assert_eq!(
                uniform(range.horizontal_alignment()?)?,
                HorizontalAlignment::CENTER
            );
            assert_eq!(
                uniform(range.vertical_alignment()?)?,
                VerticalAlignment::CENTER
            );
            assert!(uniform(range.wrap_text()?)?);

            let borders = range.borders()?;
            assert!(borders.count()? >= 4);
            for index in [
                BorderIndex::EDGE_BOTTOM,
                BorderIndex::EDGE_TOP,
                BorderIndex::EDGE_LEFT,
                BorderIndex::EDGE_RIGHT,
            ] {
                let border = borders.item(index)?;
                border.set_line_style(BorderLineStyle::CONTINUOUS)?;
                border.set_weight(BorderWeight::THIN)?;
                border.set_color(ExcelColor::from_rgb(90, 80, 70))?;
                assert_eq!(uniform(border.line_style()?)?, BorderLineStyle::CONTINUOUS);
                assert_eq!(uniform(border.weight()?)?, BorderWeight::THIN);
            }
            let enumerated = borders
                .iter()?
                .map(|border| border.and_then(|border| border.line_style()))
                .collect::<Result<Vec<_>, _>>()?;
            assert!(!enumerated.is_empty());
            eprintln!(
                "Borders count={} enumeration={enumerated:?}",
                borders.count()?
            );
            let bottom = borders.item(BorderIndex::EDGE_BOTTOM)?;
            bottom.set_line_style(BorderLineStyle::NONE)?;
            assert_eq!(uniform(bottom.line_style()?)?, BorderLineStyle::NONE);
            borders.set_line_style(BorderLineStyle::CONTINUOUS)?;
            borders.set_weight(BorderWeight::THIN)?;
            borders.set_color(ExcelColor::from_rgb(50, 60, 70))?;
            assert_eq!(uniform(borders.line_style()?)?, BorderLineStyle::CONTINUOUS);
            for index in [
                BorderIndex::EDGE_BOTTOM,
                BorderIndex::EDGE_TOP,
                BorderIndex::EDGE_LEFT,
                BorderIndex::EDGE_RIGHT,
            ] {
                assert_eq!(
                    uniform(borders.item(index)?.line_style()?)?,
                    BorderLineStyle::CONTINUOUS
                );
            }

            worksheet.range("A1:C2")?.set_row_height(20.0)?;
            worksheet.range("A3:C4")?.set_row_height(30.0)?;
            assert!(matches!(range.row_height()?, MixedValue::Mixed));
            worksheet.range("A1:C4")?.set_row_height(22.0)?;
            assert_eq!(uniform(range.row_height()?)?, 22.0);
            worksheet.range("A1:A4")?.set_column_width(14.0)?;
            worksheet.range("B1:B4")?.set_column_width(18.0)?;
            assert!(matches!(
                worksheet.range("A1:B4")?.column_width()?,
                MixedValue::Mixed
            ));
            worksheet.range("A1:C4")?.set_column_width(16.0)?;
            assert_eq!(uniform(range.column_width()?)?, 16.0);
            assert!(range.auto_fit().is_err());
            range.entire_column()?.auto_fit()?;
            range.entire_row()?.auto_fit()?;
            for cell in ["A1", "A2", "A3", "A4"] {
                assert!(
                    matches!(worksheet.range(cell)?.row_height()?, MixedValue::Uniform(value) if value.is_finite() && value > 0.0)
                );
            }
            for column in ["A1:A4", "B1:B4", "C1:C4"] {
                assert!(
                    matches!(worksheet.range(column)?.column_width()?, MixedValue::Uniform(value) if value.is_finite() && value > 0.0)
                );
            }

            let left = worksheet.range("A1")?.font()?;
            left.set_bold(true)?;
            left.set_size(11.0)?;
            left.set_color(ExcelColor::from_rgb(1, 2, 3))?;
            let right = worksheet.range("B1")?.font()?;
            right.set_bold(false)?;
            right.set_size(13.0)?;
            right.set_color(ExcelColor::from_rgb(3, 2, 1))?;
            let mixed_font = worksheet.range("A1:B1")?.font()?;
            assert!(matches!(mixed_font.bold()?, MixedValue::Mixed));
            assert!(matches!(mixed_font.size()?, MixedValue::Mixed));
            assert!(matches!(mixed_font.color()?, MixedValue::Mixed));
            let mixed_number = worksheet.range("A1:B1")?.number_format()?;
            assert!(matches!(mixed_number, MixedValue::Mixed));
            mixed_font.set_bold(true)?;
            assert!(uniform(mixed_font.bold()?)?);
            eprintln!(
                "mixed font bold={:?} size={:?} color={:?}; mixed number={mixed_number:?}",
                worksheet.range("A1:B1")?.font()?.bold()?,
                mixed_font.size()?,
                mixed_font.color()?
            );

            assert!(range.set_row_height(f64::NAN).is_err());
            assert!(range.set_column_width(f64::INFINITY).is_err());
            assert!(range.set_number_format("0\0.00").is_err());
            assert!(range.font()?.set_name("A\0rial").is_err());

            cleanup_workbook
                .take()
                .expect("fresh workbook must be available for cleanup")
                .close_without_saving()?;
            alerts.restore()?;
            Ok(())
        },
    ));
    if let Some(workbook) = cleanup_workbook.take() {
        let _ = workbook.close_without_saving();
    }
    let quit = application.quit();
    match outcome {
        Ok(result) => result?,
        Err(payload) => std::panic::resume_unwind(payload),
    }
    quit?;
    wait_for_zero_excel_processes()?;
    Ok(())
}
