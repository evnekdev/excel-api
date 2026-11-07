//! Scratch-workbook probes, compiled only by explicit acceptance builds (#126).
#[excel_api::excel_function(name = "RUST.TEST.EVENTS", thunk = "acceptance_events")]
pub fn events() -> f64 {
    excel_api::async_udf::diagnostics().map_or(-1.0, |d| d.calculation_ended as f64)
}
use excel_api::{
    ConversionLimits, DocumentInfo, ExcelError, ExcelReferenceArg, ExcelValue, ExcelValueRef,
    MacroContext, OwnedExcelReference, ReferenceArea, WorksheetContext, excel_command,
    excel_function,
};

#[excel_function(
    name = "RUST.TEST.READ",
    thunk = "acceptance_read",
    arguments(reference = "Scratch reference")
)]
pub fn read(
    reference: ExcelReferenceArg<'_>,
    context: &WorksheetContext<'_>,
) -> Result<ExcelValue, ExcelError> {
    let ExcelValueRef::Reference(reference) = reference.into_inner() else {
        return Err(ExcelError::Value);
    };
    let owned = OwnedExcelReference::from_borrowed(&reference).map_err(|_| ExcelError::Value)?;
    context
        .read_reference(&owned, &ConversionLimits::default())
        .map_err(|_| ExcelError::Value)
}

#[excel_function(name = "RUST.TEST.DOCUMENT", thunk = "acceptance_document", macro_type)]
pub fn document(context: &MacroContext<'_>) -> Result<ExcelValue, ExcelError> {
    context
        .document_info(DocumentInfo::ActiveWorkbookName, None)
        .map_err(|_| ExcelError::Value)
}

#[excel_function(
    name = "RUST.TEST.HELP",
    thunk = "acceptance_help",
    description = "Acceptance help metadata",
    help_topic = "https://example.com/help!0",
    arguments(value = "Acceptance argument")
)]
pub fn help(value: f64) -> f64 {
    value
}

fn destination() -> OwnedExcelReference {
    OwnedExcelReference::current_sheet(ReferenceArea::cell(0, 9).expect("fixed J1 coordinate"))
}

static CAPTURED: std::sync::Mutex<Option<OwnedExcelReference>> = std::sync::Mutex::new(None);

#[excel_command(name = "RUST.TEST.CAPTURE", thunk = "acceptance_capture")]
pub fn capture(context: &MacroContext<'_>) -> Result<(), ExcelError> {
    let id = context.active_sheet_id().map_err(|_| ExcelError::Value)?;
    let reference = OwnedExcelReference::on_sheet(
        usize::try_from(id).map_err(|_| ExcelError::Value)?,
        vec![ReferenceArea::cell(0, 9).map_err(|_| ExcelError::Value)?],
    )
    .map_err(|_| ExcelError::Value)?;
    *CAPTURED.lock().map_err(|_| ExcelError::Value)? = Some(reference);
    Ok(())
}

#[excel_command(name = "RUST.TEST.WRITE.CAPTURED", thunk = "acceptance_write_captured")]
pub fn write_captured(context: &MacroContext<'_>) -> Result<(), ExcelError> {
    let reference = CAPTURED
        .lock()
        .map_err(|_| ExcelError::Value)?
        .take()
        .ok_or(ExcelError::Value)?;
    context
        .set_reference(&reference, ExcelValue::Number(42.0))
        .map_err(|_| ExcelError::Value)
}

#[excel_command(name = "RUST.TEST.WRITE", thunk = "acceptance_write")]
pub fn write(context: &MacroContext<'_>) -> Result<(), ExcelError> {
    context
        .set_reference(&destination(), ExcelValue::Number(42.0))
        .map_err(|_| ExcelError::Value)
}
#[excel_command(name = "RUST.TEST.CLEAR", thunk = "acceptance_clear")]
pub fn clear(context: &MacroContext<'_>) -> Result<(), ExcelError> {
    context
        .clear_reference(&destination())
        .map_err(|_| ExcelError::Value)
}
#[excel_command(name = "RUST.TEST.CALCULATE", thunk = "acceptance_calculate")]
pub fn calculate(context: &MacroContext<'_>) -> Result<(), ExcelError> {
    context.calculate_now().map_err(|_| ExcelError::Value)
}

// Finite cooperative work: the runner can observe actual worker retirement,
// independently of the runtime's pending-request count after cancellation.
static ACTIVE_WORKERS: std::sync::atomic::AtomicU64 = std::sync::atomic::AtomicU64::new(0);
static CANCELED_WORKERS: std::sync::atomic::AtomicU64 = std::sync::atomic::AtomicU64::new(0);

#[excel_function(
    name = "RUST.TEST.DELAY",
    thunk = "acceptance_delay",
    asynchronous,
    thread_safe,
    arguments(milliseconds = "Bounded delay, at most 10000 milliseconds")
)]
pub fn delay(
    milliseconds: f64,
    cancel: excel_api::AsyncCancellationToken,
) -> Result<f64, ExcelError> {
    use std::sync::atomic::Ordering;
    if !milliseconds.is_finite() || !(0.0..=10000.0).contains(&milliseconds) {
        return Err(ExcelError::Value);
    }
    ACTIVE_WORKERS.fetch_add(1, Ordering::SeqCst);
    struct ActiveGuard;
    impl Drop for ActiveGuard {
        fn drop(&mut self) {
            ACTIVE_WORKERS.fetch_sub(1, Ordering::SeqCst);
        }
    }
    let _guard = ActiveGuard;
    let deadline =
        std::time::Instant::now() + std::time::Duration::from_millis(milliseconds as u64);
    loop {
        if cancel.is_cancellation_requested() {
            CANCELED_WORKERS.fetch_add(1, Ordering::SeqCst);
            return Err(ExcelError::Na);
        }
        if std::time::Instant::now() >= deadline {
            return Ok(42.0);
        }
        std::thread::sleep(std::time::Duration::from_millis(10));
    }
}

#[excel_function(name = "RUST.TEST.ASYNC.STATUS", thunk = "acceptance_async_status")]
pub fn async_status() -> excel_api::ExcelString {
    use std::sync::atomic::Ordering;
    let status = excel_api::async_udf::diagnostics().map_or_else(
        || "inactive".to_string(),
        |d| {
            format!(
                "{},{},{},{},{},{}",
                d.pending,
                d.calculation_canceled,
                d.calculation_ended,
                ACTIVE_WORKERS.load(Ordering::SeqCst),
                CANCELED_WORKERS.load(Ordering::SeqCst),
                d.completed
            )
        },
    );
    excel_api::ExcelString::from_utf8(&status)
}
