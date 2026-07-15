//! Finite cooperative-dispatch demonstration for manual Excel validation.

use std::sync::{
    Mutex, OnceLock,
    atomic::{AtomicUsize, Ordering},
};

use excel_api::{
    DispatchCompletionError, DispatchOperation, DispatchResult, DispatchTicket, ExcelError,
    ExcelString, ExcelValue, enqueue_dispatch,
};

static LAST_PROCESSED: AtomicUsize = AtomicUsize::new(0);

fn ticket_slot() -> &'static Mutex<Option<DispatchTicket>> {
    static SLOT: OnceLock<Mutex<Option<DispatchTicket>>> = OnceLock::new();
    SLOT.get_or_init(|| Mutex::new(None))
}

fn lock_ticket() -> std::sync::MutexGuard<'static, Option<DispatchTicket>> {
    ticket_slot()
        .lock()
        .unwrap_or_else(std::sync::PoisonError::into_inner)
}

pub fn enqueue_from_finished_background_thread() -> Result<(), ExcelError> {
    if lock_ticket()
        .as_ref()
        .is_some_and(|ticket| ticket.try_result().is_none())
    {
        return Err(ExcelError::Na);
    }
    let producer = std::thread::Builder::new()
        .name("excel-api-dispatch-demo".into())
        .spawn(|| {
            enqueue_dispatch(DispatchOperation::EchoOwned(ExcelValue::Text(
                ExcelString::from("cooperative-dispatch"),
            )))
        })
        .map_err(|_| ExcelError::Value)?;
    let ticket = producer
        .join()
        .map_err(|_| ExcelError::Value)?
        .map_err(|_| ExcelError::Na)?;
    if ticket.try_result().is_some() {
        return Err(ExcelError::Value);
    }
    *lock_ticket() = Some(ticket);
    Ok(())
}

pub fn record_pump(processed: usize) {
    LAST_PROCESSED.store(processed, Ordering::Release);
}

pub fn status() -> ExcelString {
    let ticket = lock_ticket();
    let processed = LAST_PROCESSED.load(Ordering::Acquire);
    let text = match ticket.as_ref() {
        None => format!("idle;last_pump_processed={processed}"),
        Some(ticket) => match ticket.try_result() {
            None => format!(
                "queued;generation={};request={};last_pump_processed={processed}",
                ticket.generation(),
                ticket.id()
            ),
            Some(Ok(DispatchResult::OwnedValue(value))) => format!(
                "completed;kind={};generation={};request={};last_pump_processed={processed}",
                value.kind_name(),
                ticket.generation(),
                ticket.id()
            ),
            Some(Ok(DispatchResult::CancellationRequested(value))) => format!(
                "completed;cancellation={value};generation={};request={};last_pump_processed={processed}",
                ticket.generation(),
                ticket.id()
            ),
            Some(Err(error)) => format!(
                "failed={};generation={};request={};last_pump_processed={processed}",
                completion_label(&error),
                ticket.generation(),
                ticket.id()
            ),
        },
    };
    ExcelString::from(text)
}

fn completion_label(error: &DispatchCompletionError) -> &'static str {
    match error {
        DispatchCompletionError::Canceled => "canceled",
        DispatchCompletionError::Expired => "expired",
        DispatchCompletionError::DispatcherShutdown => "shutdown",
        DispatchCompletionError::Operation(_) => "operation",
        DispatchCompletionError::WaitFromCallback => "callback-wait",
        DispatchCompletionError::WaitTimeout => "wait-timeout",
    }
}

pub fn reset() {
    if let Some(ticket) = lock_ticket().take() {
        let _ = ticket.cancel();
    }
    LAST_PROCESSED.store(0, Ordering::Release);
}
