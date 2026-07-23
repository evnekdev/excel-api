use super::*;
use windows_sys::Win32::Foundation::DISP_E_PARAMNOTFOUND;

#[test]
fn transparent_values_preserve_unknown_excel_constants() {
    assert_eq!(WindowView::from_raw(99).raw(), 99);
    assert_eq!(AutomationSecurity::FORCE_DISABLE.raw(), 3);
    assert_eq!(PageBreakType::MANUAL.raw(), -4135);
}

#[test]
fn print_out_preserves_all_optional_positions() {
    let options = PrintOutOptions {
        copies: Some(2),
        collate: Some(true),
        ..Default::default()
    };
    let mut values = PositionalArguments::new();
    values.push_optional(
        options
            .from
            .map(|v| one_based(v, "PrintOut.From"))
            .transpose()
            .expect("valid"),
    );
    values.push_optional(
        options
            .to
            .map(|v| one_based(v, "PrintOut.To"))
            .transpose()
            .expect("valid"),
    );
    values.push_optional(
        options
            .copies
            .map(|v| one_based(v, "PrintOut.Copies"))
            .transpose()
            .expect("valid"),
    );
    values.push_optional(options.preview.map(OwnedVariant::bool));
    push_optional_text(&mut values, options.active_printer).expect("text");
    values.push_optional(options.print_to_file.map(OwnedVariant::bool));
    values.push_optional(options.collate.map(OwnedVariant::bool));
    push_optional_text(&mut values, options.pr_to_file_name).expect("text");
    values.push_optional(options.ignore_print_areas.map(OwnedVariant::bool));
    let values = values.into_inner();
    assert_eq!(values.len(), 9);
    assert_eq!(values[0].as_scode(), Some(DISP_E_PARAMNOTFOUND));
    assert_eq!(values[2].as_i32(), Some(2));
    assert_eq!(values[6].as_bool(), Some(true));
    assert_eq!(values[8].as_scode(), Some(DISP_E_PARAMNOTFOUND));
}

#[test]
fn secret_debug_fields_are_redacted() {
    let worksheet = format!(
        "{:?}",
        WorksheetProtectOptions {
            password: Some("secret"),
            ..Default::default()
        }
    );
    let workbook = format!(
        "{:?}",
        WorkbookProtectOptions {
            password: Some("secret"),
            ..Default::default()
        }
    );
    let print = format!(
        "{:?}",
        PrintOutOptions {
            active_printer: Some("private printer"),
            pr_to_file_name: Some("private path"),
            ..Default::default()
        }
    );
    for text in [worksheet, workbook, print] {
        assert!(text.contains("REDACTED"));
        assert!(!text.contains("secret"));
        assert!(!text.contains("private"));
    }
}
