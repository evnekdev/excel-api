#![cfg(windows)]

use std::thread;
use std::time::{Duration, Instant};

use excel_com::{
    Application, AutomationArray, AutomationValue, ComApartment, ConversionError, Currency,
    ExcelComError, ExcelError, FormulaValue, OaDate, Worksheet, WorksheetsAddOptions,
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
        let name = String::from_utf16_lossy(&entry.szExeFile);
        if name
            .trim_end_matches('\0')
            .eq_ignore_ascii_case("EXCEL.EXE")
        {
            count += 1;
        }
        // SAFETY: snapshot remains valid and entry storage is initialized for each enumeration step.
        present = unsafe { Process32NextW(snapshot, &mut entry) } != 0;
    }
    // SAFETY: this function owns the snapshot handle and closes it exactly once.
    unsafe { CloseHandle(snapshot) };
    Ok(count)
}

fn range(worksheet: &Worksheet, address: &str) -> Result<excel_com::Range, ExcelComError> {
    worksheet.range(address)
}

#[test]
#[ignore = "launches a fresh visible Excel process; run explicitly"]
fn worksheet_range_core_live_slice_naturally_exits() -> Result<(), Box<dyn std::error::Error>> {
    assert_eq!(
        excel_process_count()?,
        0,
        "live test requires no pre-existing EXCEL.EXE"
    );
    let apartment = ComApartment::sta()?;
    let application = Application::new(&apartment)?;
    application.set_visible(true)?;

    let workbooks = application.workbooks()?;
    let workbook = workbooks.add()?;
    let worksheets = workbook.worksheets()?;
    assert!(worksheets.count()? >= 1);
    let worksheet = worksheets.item_by_index(1)?;
    let original_name = worksheet.name()?;
    assert_eq!(
        worksheets.item_by_name(&original_name)?.index()?,
        worksheet.index()?
    );
    let added = worksheets.add(WorksheetsAddOptions::new())?;
    assert!(added.index()? >= 1);
    assert!(worksheets.count()? >= 2);

    let scalar = range(&worksheet, "A1")?;
    scalar.set_value2(AutomationValue::Number(42.5))?;
    assert_eq!(scalar.value2()?, AutomationValue::Number(42.5));
    let text = range(&worksheet, "B1")?;
    text.set_value2(AutomationValue::Text("excel-com".to_owned()))?;
    assert_eq!(
        text.value2()?,
        AutomationValue::Text("excel-com".to_owned())
    );
    let error = range(&worksheet, "C1")?;
    error.set_value2(AutomationValue::Error(ExcelError::NOT_AVAILABLE))?;
    assert_eq!(
        error.value2()?,
        AutomationValue::Error(ExcelError::NOT_AVAILABLE)
    );

    let date = OaDate::new(45_000.25)?;
    let date_cell = range(&worksheet, "D1")?;
    date_cell.set_value(AutomationValue::Date(date))?;
    assert_eq!(date_cell.value()?, AutomationValue::Date(date));
    let negative_serial = range(&worksheet, "E1")?;
    negative_serial.set_value2(AutomationValue::Number(-1.25))?;
    assert_eq!(negative_serial.value2()?, AutomationValue::Number(-1.25));
    let currency = range(&worksheet, "F1")?;
    currency.set_value(AutomationValue::Currency(Currency::from_scaled(123_456)))?;

    let matrix = AutomationArray::from_rows(vec![
        vec![
            AutomationValue::Number(1.0),
            AutomationValue::Number(2.0),
            AutomationValue::Number(3.0),
        ],
        vec![
            AutomationValue::Number(4.0),
            AutomationValue::Number(5.0),
            AutomationValue::Number(6.0),
        ],
    ])?;
    let matrix_range = range(&worksheet, "A3:C4")?;
    matrix_range.set_value2(AutomationValue::Array(matrix.clone()))?;
    assert_eq!(matrix_range.value2()?, AutomationValue::Array(matrix));
    let errors = AutomationArray::from_rows(vec![
        vec![
            AutomationValue::Error(ExcelError::NOT_AVAILABLE),
            AutomationValue::Error(ExcelError::NOT_AVAILABLE),
        ],
        vec![
            AutomationValue::Error(ExcelError::NOT_AVAILABLE),
            AutomationValue::Error(ExcelError::NOT_AVAILABLE),
        ],
    ])?;
    let error_matrix = range(&worksheet, "E3:F4")?;
    error_matrix.set_value2(AutomationValue::Array(errors.clone()))?;
    assert_eq!(error_matrix.value2()?, AutomationValue::Array(errors));

    let formula = range(&worksheet, "H1")?;
    formula.set_formula("=1+1")?;
    assert_eq!(formula.formula()?, FormulaValue::Text("=1+1".to_owned()));
    let sequence = range(&worksheet, "H3")?;
    sequence.set_formula2("=SEQUENCE(2,3)")?;
    assert_eq!(
        sequence.formula2()?,
        FormulaValue::Text("=SEQUENCE(2,3)".to_owned())
    );
    let spill = range(&worksheet, "H3:J4")?;
    assert_eq!(
        spill.value2()?,
        AutomationValue::Array(AutomationArray::from_rows(vec![
            vec![
                AutomationValue::Number(1.0),
                AutomationValue::Number(2.0),
                AutomationValue::Number(3.0)
            ],
            vec![
                AutomationValue::Number(4.0),
                AutomationValue::Number(5.0),
                AutomationValue::Number(6.0)
            ],
        ])?)
    );

    assert!(matches!(
        matrix_range.set_value2(AutomationValue::Number(9.0)),
        Err(ExcelComError::Conversion(
            ConversionError::ShapeMismatch { .. }
        ))
    ));
    matrix_range.clear_contents()?;
    assert_eq!(
        matrix_range.value2()?,
        AutomationValue::Array(AutomationArray::new(2, 3, vec![AutomationValue::Empty; 6])?)
    );
    assert_eq!(matrix_range.address()?, "$A$3:$C$4");
    assert_eq!(matrix_range.row()?, 3);
    assert_eq!(matrix_range.column()?, 1);
    assert_eq!(matrix_range.row_count()?, 2);
    assert_eq!(matrix_range.column_count()?, 3);
    assert_eq!(matrix_range.cell_count()?, 6);

    drop((
        scalar,
        text,
        error,
        date_cell,
        negative_serial,
        currency,
        matrix_range,
        error_matrix,
        formula,
        sequence,
        spill,
        added,
        worksheet,
        worksheets,
    ));
    workbook.close_without_saving()?;
    drop(workbooks);
    application.quit()?;
    let deadline = Instant::now() + Duration::from_secs(15);
    while Instant::now() < deadline && excel_process_count()? != 0 {
        thread::sleep(Duration::from_millis(100));
    }
    assert_eq!(
        excel_process_count()?,
        0,
        "owned Excel process did not exit naturally"
    );
    Ok(())
}
