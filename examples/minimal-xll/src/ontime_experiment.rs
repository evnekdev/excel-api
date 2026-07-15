//! Internal compatibility probe for the historical `xlcOnTime` XLM command.
//!
//! Nothing in this module is connected to the proposed M17 dispatcher queue.

use std::collections::VecDeque;
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::{Mutex, OnceLock};
use std::time::{SystemTime, UNIX_EPOCH};

use excel_api::{
    ExcelError, ExcelString, ExperimentalOnTimeOutcome, ExperimentalOnTimeValue, LifecycleContext,
    MacroContext,
};

pub const CALLBACK_NAME: &str = "RUST.ONTIME.CALLBACK";
const ENABLE_MARKER: &str = "excel-api-xlcontime-enable.marker";
const STATUS_FILE: &str = "excel-api-xlcontime-status.json";
const MAX_PENDING: usize = 32;
const MAX_EVENTS: usize = 128;
const DELAY_SECONDS: f64 = 3.0;
const LATEST_SLACK_SECONDS: f64 = 10.0;

static NEXT_GENERATION: AtomicU64 = AtomicU64::new(0);
static CALLBACK_COUNT: AtomicU64 = AtomicU64::new(0);
static STALE_CALLBACK_COUNT: AtomicU64 = AtomicU64::new(0);
static NEXT_ORDER: AtomicU64 = AtomicU64::new(0);

#[derive(Clone, Copy, Debug)]
enum ScheduleForm {
    TwoArguments,
    WithLatestTime,
}

impl ScheduleForm {
    const fn label(self) -> &'static str {
        match self {
            Self::TwoArguments => "two",
            Self::WithLatestTime => "latest",
        }
    }
}

#[derive(Clone, Copy, Debug)]
struct Pending {
    generation: u64,
    earliest: f64,
    form: ScheduleForm,
}

#[derive(Clone, Debug)]
struct Event {
    order: u64,
    kind: &'static str,
    generation: u64,
    thread_id: u64,
    timestamp_ms: u128,
    serial: Option<f64>,
    form: Option<ScheduleForm>,
    raw_code: Option<i32>,
    result: String,
}

#[derive(Default)]
struct State {
    active: bool,
    generation: u64,
    main_thread_id: u64,
    schedule_attempts: u64,
    pending: Vec<Pending>,
    events: VecDeque<Event>,
}

fn state() -> &'static Mutex<State> {
    static STATE: OnceLock<Mutex<State>> = OnceLock::new();
    STATE.get_or_init(|| Mutex::new(State::default()))
}

fn lock_state() -> std::sync::MutexGuard<'static, State> {
    state()
        .lock()
        .unwrap_or_else(std::sync::PoisonError::into_inner)
}

fn bounded_increment(counter: &AtomicU64) -> u64 {
    counter
        .fetch_update(Ordering::AcqRel, Ordering::Acquire, |value| {
            Some(value.saturating_add(1))
        })
        .unwrap_or(u64::MAX)
        .saturating_add(1)
}

#[cfg(windows)]
fn current_thread_id() -> u64 {
    #[link(name = "kernel32")]
    unsafe extern "system" {
        fn GetCurrentThreadId() -> u32;
    }
    // SAFETY: this Windows API takes no arguments and has no validity preconditions.
    u64::from(unsafe { GetCurrentThreadId() })
}

#[cfg(not(windows))]
fn current_thread_id() -> u64 {
    use std::hash::{Hash, Hasher};
    let mut hasher = std::collections::hash_map::DefaultHasher::new();
    std::thread::current().id().hash(&mut hasher);
    hasher.finish()
}

fn timestamp_ms() -> u128 {
    SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap_or_default()
        .as_millis()
}

fn enable_marker_path() -> std::path::PathBuf {
    harness_directory().join(ENABLE_MARKER)
}

fn harness_enabled_for_this_process() -> bool {
    std::fs::read_to_string(enable_marker_path())
        .ok()
        .and_then(|value| value.trim().parse::<u32>().ok())
        == Some(std::process::id())
}

fn status_path() -> std::path::PathBuf {
    std::env::var_os("EXCEL_API_ONTIME_STATUS_PATH")
        .map(std::path::PathBuf::from)
        .unwrap_or_else(|| harness_directory().join(STATUS_FILE))
}

fn harness_directory() -> std::path::PathBuf {
    std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("..")
        .join("..")
        .join("target")
        .join("ontime-validation")
}

fn outcome_text(outcome: ExperimentalOnTimeOutcome) -> String {
    match outcome.value {
        ExperimentalOnTimeValue::Boolean(value) => format!("bool:{value}"),
        ExperimentalOnTimeValue::ExcelError(value) => format!("error:{value}"),
        ExperimentalOnTimeValue::Other(xltype) => format!("xltype:{xltype}"),
    }
}

fn push_event(mut state: std::sync::MutexGuard<'static, State>, event: Event) {
    if state.events.len() == MAX_EVENTS {
        state.events.pop_front();
    }
    state.events.push_back(event);
}

fn record_call(
    kind: &'static str,
    generation: u64,
    serial: Option<f64>,
    form: Option<ScheduleForm>,
    outcome: Result<ExperimentalOnTimeOutcome, impl std::fmt::Display>,
) -> bool {
    let (raw_code, result, accepted) = match outcome {
        Ok(outcome) => (
            Some(outcome.return_code.0),
            outcome_text(outcome),
            outcome.accepted(),
        ),
        Err(error) => (None, format!("call-error:{error}"), false),
    };
    push_event(
        lock_state(),
        Event {
            order: bounded_increment(&NEXT_ORDER),
            kind,
            generation,
            thread_id: current_thread_id(),
            timestamp_ms: timestamp_ms(),
            serial,
            form,
            raw_code,
            result,
        },
    );
    if harness_enabled_for_this_process() {
        let _ = dump();
    }
    accepted
}

pub fn activate() {
    let generation = bounded_increment(&NEXT_GENERATION);
    let mut state = lock_state();
    state.active = true;
    state.generation = generation;
    state.main_thread_id = current_thread_id();
    state.schedule_attempts = 0;
    state.pending.clear();
    state.events.clear();
    CALLBACK_COUNT.store(0, Ordering::Release);
    STALE_CALLBACK_COUNT.store(0, Ordering::Release);
    NEXT_ORDER.store(0, Ordering::Release);
    drop(state);
    push_event(
        lock_state(),
        Event {
            order: bounded_increment(&NEXT_ORDER),
            kind: "activate",
            generation,
            thread_id: current_thread_id(),
            timestamp_ms: timestamp_ms(),
            serial: None,
            form: None,
            raw_code: None,
            result: "active".into(),
        },
    );
}

pub fn bootstrap(context: &LifecycleContext<'_>) -> bool {
    if !harness_enabled_for_this_process() {
        return true;
    }
    let generation = lock_state().generation;
    let now = match context.experimental_excel_serial_now() {
        Ok(value) => value,
        Err(error) => {
            let _ = record_call(
                "bootstrap-now",
                generation,
                None,
                None,
                Err::<ExperimentalOnTimeOutcome, _>(error),
            );
            let _ = dump();
            return false;
        }
    };
    let first = now + DELAY_SECONDS / 86_400.0;
    let second = now + (DELAY_SECONDS + 2.0) / 86_400.0;
    let latest = second + LATEST_SLACK_SECONDS / 86_400.0;

    let first_outcome = context.experimental_schedule_on_time(first, CALLBACK_NAME);
    let first_accepted = record_call(
        "schedule",
        generation,
        Some(first),
        Some(ScheduleForm::TwoArguments),
        first_outcome,
    );
    if first_accepted {
        lock_state().pending.push(Pending {
            generation,
            earliest: first,
            form: ScheduleForm::TwoArguments,
        });
    }

    let second_outcome =
        context.experimental_schedule_on_time_with_latest(second, CALLBACK_NAME, latest);
    let second_accepted = record_call(
        "schedule",
        generation,
        Some(second),
        Some(ScheduleForm::WithLatestTime),
        second_outcome,
    );
    if second_accepted {
        lock_state().pending.push(Pending {
            generation,
            earliest: second,
            form: ScheduleForm::WithLatestTime,
        });
    }

    let cancel = context.experimental_cancel_on_time(second, CALLBACK_NAME);
    let cancel_accepted = record_call(
        "cancel",
        generation,
        Some(second),
        Some(ScheduleForm::WithLatestTime),
        cancel,
    );
    if cancel_accepted {
        lock_state()
            .pending
            .retain(|pending| pending.earliest != second);
    }
    let repeated = context.experimental_cancel_on_time(second, CALLBACK_NAME);
    let repeated_rejected = matches!(
        &repeated,
        Ok(outcome)
            if outcome.return_code.is_success()
                && matches!(outcome.value, ExperimentalOnTimeValue::ExcelError(value) if value == excel_api_sys::xlerrValue)
    );
    let _ = record_call(
        "cancel-missing",
        generation,
        Some(second),
        Some(ScheduleForm::WithLatestTime),
        repeated,
    );

    let reschedule =
        context.experimental_schedule_on_time_with_latest(second, CALLBACK_NAME, latest);
    let rescheduled = record_call(
        "schedule",
        generation,
        Some(second),
        Some(ScheduleForm::WithLatestTime),
        reschedule,
    );
    if rescheduled {
        lock_state().pending.push(Pending {
            generation,
            earliest: second,
            form: ScheduleForm::WithLatestTime,
        });
    }
    lock_state().schedule_attempts = 2;
    if harness_enabled_for_this_process() {
        let _ = dump();
    }
    first_accepted && second_accepted && cancel_accepted && repeated_rejected && rescheduled
}

pub fn schedule(context: &MacroContext<'_>) -> Result<(), ExcelError> {
    let (generation, form) = {
        let mut state = lock_state();
        if !state.active || state.pending.len() >= MAX_PENDING {
            return Err(ExcelError::Na);
        }
        let form = if state.schedule_attempts % 2 == 0 {
            ScheduleForm::TwoArguments
        } else {
            ScheduleForm::WithLatestTime
        };
        state.schedule_attempts = state.schedule_attempts.saturating_add(1);
        (state.generation, form)
    };
    let now = context
        .experimental_excel_serial_now()
        .map_err(|_| ExcelError::Value)?;
    let earliest = now + DELAY_SECONDS / 86_400.0;
    let latest = matches!(form, ScheduleForm::WithLatestTime)
        .then_some(earliest + LATEST_SLACK_SECONDS / 86_400.0);
    let outcome = match latest {
        Some(latest) => {
            context.experimental_schedule_on_time_with_latest(earliest, CALLBACK_NAME, latest)
        }
        None => context.experimental_schedule_on_time(earliest, CALLBACK_NAME),
    };
    let accepted = record_call("schedule", generation, Some(earliest), Some(form), outcome);
    if !accepted {
        return Err(ExcelError::Value);
    }
    let mut state = lock_state();
    if !state.active || state.generation != generation {
        drop(state);
        let cancellation = context.experimental_cancel_on_time(earliest, CALLBACK_NAME);
        let _ = record_call(
            "schedule-race-cancel",
            generation,
            Some(earliest),
            Some(form),
            cancellation,
        );
        return Err(ExcelError::Na);
    }
    state.pending.push(Pending {
        generation,
        earliest,
        form,
    });
    Ok(())
}

pub fn callback(context: &MacroContext<'_>) -> Result<(), ExcelError> {
    let pending = {
        let mut state = lock_state();
        if !state.active {
            bounded_increment(&STALE_CALLBACK_COUNT);
            return Ok(());
        }
        let generation = state.generation;
        state
            .pending
            .iter()
            .position(|pending| pending.generation == generation)
            .map(|index| state.pending.remove(index))
    };
    let Some(pending) = pending else {
        bounded_increment(&STALE_CALLBACK_COUNT);
        return Ok(());
    };
    let callback_count = bounded_increment(&CALLBACK_COUNT);
    let cancellation_poll = context.is_cancellation_requested();
    let result = match cancellation_poll {
        Ok(value) => format!("macro-context-xlAbort:{value}"),
        Err(error) => format!("macro-context-error:{error}"),
    };
    push_event(
        lock_state(),
        Event {
            order: bounded_increment(&NEXT_ORDER),
            kind: "callback",
            generation: pending.generation,
            thread_id: current_thread_id(),
            timestamp_ms: timestamp_ms(),
            serial: Some(pending.earliest),
            form: Some(pending.form),
            raw_code: None,
            result,
        },
    );
    if callback_count == 2 {
        let _ = schedule(context);
    }
    let _ = dump();
    Ok(())
}

pub fn cancel_one(context: &MacroContext<'_>) -> Result<(), ExcelError> {
    let pending = {
        let mut state = lock_state();
        if !state.active {
            return Err(ExcelError::Na);
        }
        state.pending.pop()
    }
    .ok_or(ExcelError::Na)?;
    let outcome = context.experimental_cancel_on_time(pending.earliest, CALLBACK_NAME);
    let accepted = record_call(
        "cancel",
        pending.generation,
        Some(pending.earliest),
        Some(pending.form),
        outcome,
    );
    if accepted {
        Ok(())
    } else {
        lock_state().pending.push(pending);
        Err(ExcelError::Value)
    }
}

pub fn shutdown(context: &LifecycleContext<'_>) -> bool {
    let pending = {
        let mut state = lock_state();
        state.active = false;
        std::mem::take(&mut state.pending)
    };
    let mut failed = Vec::new();
    for pending in pending {
        let outcome = context.experimental_cancel_on_time(pending.earliest, CALLBACK_NAME);
        if !record_call(
            "close-cancel",
            pending.generation,
            Some(pending.earliest),
            Some(pending.form),
            outcome,
        ) {
            failed.push(pending);
        }
    }
    let success = failed.is_empty();
    lock_state().pending.extend(failed);
    if harness_enabled_for_this_process() {
        let _ = dump();
    }
    success
}

pub fn status() -> ExcelString {
    let state = lock_state();
    let events = state
        .events
        .iter()
        .map(|event| {
            format!(
                "{{\"order\":{},\"kind\":\"{}\",\"generation\":{},\"thread_id\":{},\"timestamp_ms\":{},\"serial\":{},\"form\":{},\"raw_code\":{},\"result\":\"{}\"}}",
                event.order,
                event.kind,
                event.generation,
                event.thread_id,
                event.timestamp_ms,
                event.serial.map_or_else(|| "null".into(), |value| value.to_string()),
                event.form.map_or_else(|| "null".into(), |value| format!("\"{}\"", value.label())),
                event.raw_code.map_or_else(|| "null".into(), |value| value.to_string()),
                event.result.replace('"', "'")
            )
        })
        .collect::<Vec<_>>()
        .join(",");
    ExcelString::from(format!(
        "{{\"active\":{},\"generation\":{},\"process_id\":{},\"main_thread_id\":{},\"callback_count\":{},\"stale_callback_count\":{},\"pending_count\":{},\"events\":[{}]}}",
        state.active,
        state.generation,
        std::process::id(),
        state.main_thread_id,
        CALLBACK_COUNT.load(Ordering::Acquire),
        STALE_CALLBACK_COUNT.load(Ordering::Acquire),
        state.pending.len(),
        events
    ))
}

pub fn dump() -> Result<(), ExcelError> {
    let path = status_path();
    let temporary = path.with_extension("tmp");
    let contents = status().to_string().map_err(|_| ExcelError::Value)?;
    std::fs::write(&temporary, contents).map_err(|_| ExcelError::Value)?;
    if path.exists() {
        std::fs::remove_file(&path).map_err(|_| ExcelError::Value)?;
    }
    std::fs::rename(temporary, path).map_err(|_| ExcelError::Value)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn diagnostics_are_bounded_and_status_is_json_shaped() {
        activate();
        for _ in 0..(MAX_EVENTS + 20) {
            push_event(
                lock_state(),
                Event {
                    order: bounded_increment(&NEXT_ORDER),
                    kind: "test",
                    generation: 1,
                    thread_id: 1,
                    timestamp_ms: 1,
                    serial: None,
                    form: None,
                    raw_code: Some(0),
                    result: "ok".into(),
                },
            );
        }
        assert_eq!(lock_state().events.len(), MAX_EVENTS);
        let status = status().to_string().unwrap();
        assert!(status.starts_with("{\"active\":true"));
        assert!(status.contains("\"events\":["));
    }
}
