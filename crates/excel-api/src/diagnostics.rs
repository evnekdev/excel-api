//! Bounded, non-panicking diagnostics for supportability-critical paths.
//!
//! Emission never calls Excel, allocates, formats, or invokes user code while
//! another diagnostic is being emitted. Drop and `xlAutoFree12` deliberately
//! do not emit events.

use core::sync::atomic::{AtomicBool, AtomicU64, AtomicUsize, Ordering};

/// Stable machine-readable diagnostic code. Formatting belongs to consumers.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
#[repr(u16)]
pub enum DiagnosticCode {
    /// Runtime initialization, closing, or lifecycle state.
    Runtime = 1,
    /// Function or command registration work.
    Registration = 2,
    /// A typed Excel C API invocation.
    ExcelCall = 3,
    /// A generated callback thunk contained a failure or panic.
    ThunkFailure = 4,
    /// Return planning or materialization failed.
    ReturnFailure = 5,
    /// An owned resource could not be released as required.
    ReleaseFailure = 6,
    /// An internal ownership invariant was detected without panicking.
    OwnershipInvariant = 7,
    /// A consumer-provided diagnostic sink panicked.
    SinkPanic = 8,
    /// Cooperative dispatcher admission, drain, or retirement activity.
    Dispatcher = 9,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
#[repr(u8)]
/// Severity assigned by the producer; no formatting or policy is implied.
pub enum DiagnosticSeverity {
    /// Verbose diagnostic information.
    Debug,
    /// Normal informational evidence.
    Info,
    /// A recoverable condition needing attention.
    Warn,
    /// A failed operation or invariant.
    Error,
}

/// Pointer-free fixed-size event. `excel_code` preserves raw Excel status.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct DiagnosticEvent {
    /// Monotonic process-local sequence assigned when the event is emitted.
    pub sequence: u64,
    /// Producer-provided correlation identifier, defaulting to [`Self::sequence`].
    pub correlation_id: u64,
    /// Stable machine-readable event category.
    pub code: DiagnosticCode,
    /// Producer-selected severity.
    pub severity: DiagnosticSeverity,
    /// Exact raw Excel return code when the event relates to an Excel call.
    pub excel_code: i32,
}

impl DiagnosticEvent {
    /// Creates an event before emission assigns its sequence number.
    pub const fn new(code: DiagnosticCode, severity: DiagnosticSeverity, excel_code: i32) -> Self {
        Self {
            sequence: 0,
            correlation_id: 0,
            code,
            severity,
            excel_code,
        }
    }

    /// Replaces the default correlation identifier for related event streams.
    pub const fn with_correlation(mut self, correlation_id: u64) -> Self {
        self.correlation_id = correlation_id;
        self
    }
}

const CAPACITY: usize = 64;
static NEXT: AtomicU64 = AtomicU64::new(1);
static WRITTEN: AtomicUsize = AtomicUsize::new(0);
static EMITTING: AtomicBool = AtomicBool::new(false);
static RING: std::sync::Mutex<[Option<DiagnosticEvent>; CAPACITY]> =
    std::sync::Mutex::new([None; CAPACITY]);
static USER_SINK: std::sync::Mutex<Option<&'static dyn DiagnosticSink>> =
    std::sync::Mutex::new(None);

/// Optional consumer for ordinary process diagnostics.
///
/// Implementations must not call Excel, block indefinitely, allocate in a
/// guarded callback path, or call [`emit`] recursively. Panics are contained
/// and the triggering event remains in the ring.
pub trait DiagnosticSink: Send + Sync {
    /// Receives one copied event outside Excel and ownership cleanup paths.
    fn record(&self, event: DiagnosticEvent);
}

/// Install or remove the optional process-wide user sink before callbacks.
/// A concurrently emitting callback is never blocked: it simply skips a sink
/// whose configuration lock is unavailable.
pub fn set_user_sink(sink: Option<&'static dyn DiagnosticSink>) {
    if let Ok(mut slot) = USER_SINK.try_lock() {
        *slot = sink;
    }
}

/// Emit into the bounded process-local ring. Reentrant and poisoned paths drop
/// the event rather than block, recurse, panic, allocate, or call Excel.
pub fn emit(mut event: DiagnosticEvent) {
    if EMITTING.swap(true, Ordering::Acquire) {
        return;
    }
    event.sequence = NEXT.fetch_add(1, Ordering::Relaxed);
    if event.correlation_id == 0 {
        event.correlation_id = event.sequence;
    }
    if let Ok(mut ring) = RING.try_lock() {
        let index = WRITTEN.fetch_add(1, Ordering::Relaxed) % CAPACITY;
        ring[index] = Some(event);
    }
    if let Ok(sink) = USER_SINK.try_lock() {
        if let Some(sink) = *sink {
            let _ = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| sink.record(event)));
        }
    }
    EMITTING.store(false, Ordering::Release);
}

/// Snapshot is intentionally for ordinary support tooling, never callbacks.
pub fn snapshot() -> Vec<DiagnosticEvent> {
    RING.lock()
        .ok()
        .map(|ring| ring.iter().flatten().copied().collect())
        .unwrap_or_default()
}

#[cfg(test)]
mod tests {
    use super::*;
    static TEST_LOCK: std::sync::Mutex<()> = std::sync::Mutex::new(());

    #[test]
    fn saturation_is_bounded_and_preserves_excel_codes() {
        let _serial = TEST_LOCK.lock().unwrap();
        for code in 0..(CAPACITY as i32 + 4) {
            emit(DiagnosticEvent::new(
                DiagnosticCode::ExcelCall,
                DiagnosticSeverity::Error,
                code,
            ));
        }
        let events = snapshot();
        assert!(events.len() <= CAPACITY);
        assert!(
            events
                .iter()
                .any(|event| event.excel_code == CAPACITY as i32 + 3)
        );
    }

    struct PanickingSink;
    impl DiagnosticSink for PanickingSink {
        fn record(&self, _: DiagnosticEvent) {
            panic!("test sink panic");
        }
    }
    static PANICKING: PanickingSink = PanickingSink;

    #[test]
    fn sink_panic_and_reentrancy_do_not_escape() {
        let _serial = TEST_LOCK.lock().unwrap();
        set_user_sink(Some(&PANICKING));
        emit(DiagnosticEvent::new(
            DiagnosticCode::ThunkFailure,
            DiagnosticSeverity::Error,
            17,
        ));
        set_user_sink(None);
        assert!(snapshot().iter().any(|event| event.excel_code == 17));
    }
}
