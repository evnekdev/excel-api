//! Bounded cooperative dispatch under genuine Excel callback capabilities.
//!
//! Enqueueing a request does not wake Excel. Work executes only when an XLL
//! callback explicitly drains through one of the typed context entry points.
//!
//! There is intentionally no context-free drain API:
//!
//! ```compile_fail
//! excel_api::dispatcher::drain();
//! ```

use core::{fmt, marker::PhantomData, panic::AssertUnwindSafe};
use std::{
    cell::Cell,
    collections::{HashMap, VecDeque},
    sync::{
        Arc, Condvar, Mutex, OnceLock, Weak,
        atomic::{AtomicBool, AtomicU8, AtomicU64, Ordering},
    },
    time::{Duration, Instant},
};

use crate::{
    DiagnosticCode, DiagnosticEvent, DiagnosticSeverity, ExcelCallError, ExcelValue,
    LifecycleContext, MacroContext, ThreadSafeContext, WorksheetContext,
};

const QUEUED: u8 = 0;
const SELECTED: u8 = 1;
const RUNNING: u8 = 2;
const COMPLETED: u8 = 3;
const FAILED: u8 = 4;
const CANCELED: u8 = 5;
const EXPIRED: u8 = 6;
const SHUTDOWN: u8 = 7;

const EVENT_ENQUEUE_ACCEPTED: i32 = 1;
const EVENT_ENQUEUE_REJECTED: i32 = 2;
const EVENT_SELECTED: i32 = 3;
const EVENT_EXECUTED: i32 = 4;
const EVENT_CANCELED: i32 = 5;
const EVENT_EXPIRED: i32 = 6;
const EVENT_COMPLETED: i32 = 7;
const EVENT_FAILED: i32 = 8;
const EVENT_SHUTDOWN: i32 = 9;
const EVENT_INCOMPATIBLE_SKIPPED: i32 = 10;
const EVENT_NESTED_SUPPRESSED: i32 = 11;
const EVENT_INTERNAL_INVARIANT: i32 = 12;
const MAXIMUM_CONDVAR_WAIT_SLICE: Duration = Duration::from_secs(24 * 60 * 60);

/// Bounds for one cooperative dispatcher generation.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct DispatchConfig {
    pub maximum_pending: usize,
    pub maximum_batch_per_drain: usize,
    pub default_timeout: Option<Duration>,
    pub maximum_drain_duration: Option<Duration>,
}

impl Default for DispatchConfig {
    fn default() -> Self {
        Self {
            maximum_pending: 64,
            maximum_batch_per_drain: 16,
            default_timeout: Some(Duration::from_secs(300)),
            maximum_drain_duration: Some(Duration::from_millis(50)),
        }
    }
}

impl DispatchConfig {
    fn normalized(self) -> Option<Self> {
        (self.maximum_pending > 0 && self.maximum_batch_per_drain > 0).then_some(self)
    }
}

/// Closed capability requirement recorded by every operation.
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
pub enum DispatchRequirement {
    ContextNeutral,
    ThreadSafeWorksheet,
    Worksheet,
    Macro,
    Lifecycle,
}

/// Closed owned operation catalogue. No arbitrary context-taking closure is accepted.
#[derive(Clone, Debug, PartialEq)]
pub enum DispatchOperation {
    /// Returns the same fully owned semantic value without calling Excel.
    EchoOwned(ExcelValue),
    /// Calls verified `xlAbort` with its preserving zero-argument form.
    PollCancellationPreservingBreak,
    #[cfg(test)]
    TestEnqueueNested(ExcelValue),
    #[cfg(test)]
    TestNestedMacroDrain,
    #[cfg(test)]
    TestBlock,
    #[cfg(test)]
    TestPanic,
}

impl DispatchOperation {
    pub const fn requirement(&self) -> DispatchRequirement {
        match self {
            Self::EchoOwned(_) => DispatchRequirement::ContextNeutral,
            Self::PollCancellationPreservingBreak => DispatchRequirement::Macro,
            #[cfg(test)]
            Self::TestEnqueueNested(_) | Self::TestBlock | Self::TestPanic => {
                DispatchRequirement::ContextNeutral
            }
            #[cfg(test)]
            Self::TestNestedMacroDrain => DispatchRequirement::Macro,
        }
    }

    pub const fn calls_excel(&self) -> bool {
        match self {
            Self::PollCancellationPreservingBreak => true,
            #[cfg(test)]
            Self::TestNestedMacroDrain => true,
            _ => false,
        }
    }
}

/// Owned completion values for the initial operation catalogue.
#[derive(Clone, Debug, PartialEq)]
pub enum DispatchResult {
    OwnedValue(ExcelValue),
    CancellationRequested(bool),
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum DispatchEnqueueError {
    NoActiveGeneration,
    QueueFull,
    RuntimeClosing,
    StaleGeneration,
}

impl fmt::Display for DispatchEnqueueError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::NoActiveGeneration => "no cooperative dispatcher generation is active",
            Self::QueueFull => "the cooperative dispatcher queue is full",
            Self::RuntimeClosing => "the cooperative dispatcher generation is closing",
            Self::StaleGeneration => "the cooperative dispatcher generation is stale",
        })
    }
}

impl std::error::Error for DispatchEnqueueError {}

#[derive(Clone, Debug, Eq, PartialEq)]
pub enum DispatchExecutionError {
    IncompatibleContext {
        required: DispatchRequirement,
        actual: DispatchCallbackKind,
    },
    Excel(ExcelCallError),
    Panicked,
    InternalInvariant,
}

impl fmt::Display for DispatchExecutionError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::IncompatibleContext { required, actual } => {
                write!(
                    formatter,
                    "{actual:?} cannot execute {required:?} dispatch work"
                )
            }
            Self::Excel(error) => write!(formatter, "dispatcher Excel call failed: {error}"),
            Self::Panicked => formatter.write_str("dispatcher operation panicked"),
            Self::InternalInvariant => {
                formatter.write_str("dispatcher internal execution invariant failed")
            }
        }
    }
}

impl std::error::Error for DispatchExecutionError {}

#[derive(Clone, Debug, Eq, PartialEq)]
pub enum DispatchCompletionError {
    Canceled,
    Expired,
    DispatcherShutdown,
    Operation(DispatchExecutionError),
    WaitFromCallback,
    WaitTimeout,
}

impl fmt::Display for DispatchCompletionError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Canceled => formatter.write_str("dispatch request was canceled"),
            Self::Expired => formatter.write_str("dispatch request expired before execution"),
            Self::DispatcherShutdown => {
                formatter.write_str("dispatcher shut down before completion")
            }
            Self::Operation(error) => write!(formatter, "{error}"),
            Self::WaitFromCallback => {
                formatter.write_str("waiting is forbidden from an Excel callback")
            }
            Self::WaitTimeout => formatter.write_str("timed out waiting for dispatch completion"),
        }
    }
}

impl std::error::Error for DispatchCompletionError {}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum DispatchCancelOutcome {
    Canceled,
    TooLate,
    AlreadyTerminal,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum DispatchCallbackKind {
    ThreadSafe,
    Worksheet,
    Macro,
    Lifecycle,
}

impl DispatchCallbackKind {
    pub const fn permits(self, requirement: DispatchRequirement) -> bool {
        matches!(
            (self, requirement),
            (_, DispatchRequirement::ContextNeutral)
                | (Self::ThreadSafe, DispatchRequirement::ThreadSafeWorksheet)
                | (Self::Worksheet, DispatchRequirement::Worksheet)
                | (Self::Macro, DispatchRequirement::Macro)
                | (Self::Lifecycle, DispatchRequirement::Lifecycle)
        )
    }
}

#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct DispatchDrainReport {
    pub selected: usize,
    pub processed: usize,
    pub failed: usize,
    pub expired: usize,
    pub incompatible_skipped: usize,
    pub nested_suppressed: bool,
    pub duration_limit_reached: bool,
}

#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct DispatchDiagnostics {
    pub generation: u64,
    pub pending: usize,
    pub running: usize,
    pub accepted: u64,
    pub rejected: u64,
    pub selected: u64,
    pub executed: u64,
    pub canceled: u64,
    pub expired: u64,
    pub completed: u64,
    pub failed: u64,
    pub shutdown_retired: u64,
    pub incompatible_skipped: u64,
    pub nested_suppressed: u64,
}

struct Completion {
    result: Option<Result<DispatchResult, DispatchCompletionError>>,
}

struct Request {
    id: u64,
    generation: u64,
    requirement: DispatchRequirement,
    operation: Mutex<Option<DispatchOperation>>,
    state: AtomicU8,
    retired: AtomicBool,
    completion: Mutex<Completion>,
    ready: Condvar,
    deadline: Option<Instant>,
    owner: Weak<DispatchController>,
    #[cfg(test)]
    wait_returns: AtomicU64,
}

impl Request {
    fn lock_completion(&self) -> std::sync::MutexGuard<'_, Completion> {
        self.completion
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner)
    }

    fn publish(&self, state: u8, result: Result<DispatchResult, DispatchCompletionError>) -> bool {
        let mut completion = self.lock_completion();
        if completion.result.is_some() {
            return false;
        }
        self.state.store(state, Ordering::Release);
        completion.result = Some(result);
        self.ready.notify_all();
        true
    }

    fn retire(&self) {
        if self.retired.swap(true, Ordering::AcqRel) {
            return;
        }
        if let Some(owner) = self.owner.upgrade() {
            owner.retire(self.id, self as *const Self);
        }
    }

    fn expire_if_due(&self) {
        if self
            .deadline
            .is_some_and(|deadline| Instant::now() >= deadline)
            && let Some(owner) = self.owner.upgrade()
        {
            owner.expire(self);
        }
    }
}

/// An owned request handle containing no Excel pointer or callback capability.
///
/// Dropping a ticket detaches it: queued work remains owned by its generation
/// and is eventually executed, canceled, expired, or retired by shutdown.
pub struct DispatchTicket {
    request: Arc<Request>,
}

impl fmt::Debug for DispatchTicket {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("DispatchTicket")
            .field("id", &self.id())
            .field("generation", &self.generation())
            .finish_non_exhaustive()
    }
}

impl DispatchTicket {
    pub fn id(&self) -> u64 {
        self.request.id
    }

    pub fn generation(&self) -> u64 {
        self.request.generation
    }

    pub fn try_result(&self) -> Option<Result<DispatchResult, DispatchCompletionError>> {
        self.request.expire_if_due();
        self.request.lock_completion().result.clone()
    }

    pub fn wait_timeout(
        &self,
        timeout: Duration,
    ) -> Result<DispatchResult, DispatchCompletionError> {
        if callback_depth() > 0 {
            return Err(DispatchCompletionError::WaitFromCallback);
        }
        // An already-due queued request wins before this wait establishes its
        // caller deadline.
        self.request.expire_if_due();
        let wait_started = Instant::now();
        let caller_deadline = wait_started.checked_add(timeout);
        loop {
            let completion = self.request.lock_completion();
            if let Some(result) = completion.result.clone() {
                return result;
            }

            let now = Instant::now();
            let queued_deadline = (self.request.state.load(Ordering::Acquire) == QUEUED)
                .then_some(self.request.deadline)
                .flatten();
            let caller_due = caller_deadline.is_some_and(|deadline| now >= deadline);
            let request_due = queued_deadline.is_some_and(|deadline| now >= deadline);

            // If an oversleep crosses both deadlines, preserve which deadline
            // was earlier rather than allowing the later request expiry to
            // overwrite a caller timeout that had already won.
            if caller_due
                && queued_deadline.is_none_or(|request_deadline| {
                    caller_deadline.is_some_and(|caller| caller < request_deadline)
                })
            {
                return Err(DispatchCompletionError::WaitTimeout);
            }
            if request_due {
                drop(completion);
                self.request.expire_if_due();
                continue;
            }
            let caller_remaining = caller_deadline.map_or_else(
                || timeout.saturating_sub(wait_started.elapsed()),
                |deadline| deadline.saturating_duration_since(now),
            );
            if caller_remaining.is_zero() {
                return Err(DispatchCompletionError::WaitTimeout);
            }

            // Expiry can win only while Queued. Once selection commits, the
            // original queue deadline no longer limits an active wait.
            let request_remaining =
                queued_deadline.map(|deadline| deadline.saturating_duration_since(now));
            let wait_for = request_remaining
                .map_or(caller_remaining, |remaining| {
                    remaining.min(caller_remaining)
                })
                .min(MAXIMUM_CONDVAR_WAIT_SLICE);
            if wait_for.is_zero() {
                drop(completion);
                self.request.expire_if_due();
                continue;
            }

            let (completion, _) = self
                .request
                .ready
                .wait_timeout(completion, wait_for)
                .unwrap_or_else(std::sync::PoisonError::into_inner);
            #[cfg(test)]
            self.request.wait_returns.fetch_add(1, Ordering::Relaxed);
            drop(completion);
        }
    }

    pub fn cancel(&self) -> DispatchCancelOutcome {
        self.request
            .owner
            .upgrade()
            .map_or(DispatchCancelOutcome::AlreadyTerminal, |owner| {
                owner.cancel(&self.request)
            })
    }
}

/// Stable handle to one dispatcher generation. A stale handle never redirects
/// work into a later generation.
#[derive(Clone)]
pub struct DispatchGeneration {
    controller: Arc<DispatchController>,
}

impl fmt::Debug for DispatchGeneration {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("DispatchGeneration")
            .field("id", &self.id())
            .finish_non_exhaustive()
    }
}

impl DispatchGeneration {
    pub fn id(&self) -> u64 {
        self.controller.generation
    }

    pub fn enqueue(
        &self,
        operation: DispatchOperation,
    ) -> Result<DispatchTicket, DispatchEnqueueError> {
        if active_generation_id() != Some(self.id()) {
            emit_request(EVENT_ENQUEUE_REJECTED, self.id());
            self.controller
                .diagnostics
                .rejected
                .fetch_add(1, Ordering::Relaxed);
            return Err(DispatchEnqueueError::StaleGeneration);
        }
        self.controller.enqueue(operation)
    }

    pub fn diagnostics(&self) -> DispatchDiagnostics {
        self.controller.diagnostics()
    }
}

#[derive(Default)]
struct Counters {
    accepted: AtomicU64,
    rejected: AtomicU64,
    selected: AtomicU64,
    executed: AtomicU64,
    canceled: AtomicU64,
    expired: AtomicU64,
    completed: AtomicU64,
    failed: AtomicU64,
    shutdown_retired: AtomicU64,
    incompatible_skipped: AtomicU64,
    nested_suppressed: AtomicU64,
}

struct ControllerState {
    active: bool,
    pending: VecDeque<Arc<Request>>,
    registry: HashMap<u64, Arc<Request>>,
    running: usize,
}

struct DispatchController {
    generation: u64,
    config: DispatchConfig,
    next_id: AtomicU64,
    state: Mutex<ControllerState>,
    idle: Condvar,
    diagnostics: Counters,
    #[cfg(test)]
    before_start: Mutex<Option<Arc<TestPause>>>,
    #[cfg(test)]
    panic_after_start: AtomicBool,
    #[cfg(test)]
    clear_operation_after_start: AtomicBool,
}

impl DispatchController {
    fn new(generation: u64, config: DispatchConfig) -> Arc<Self> {
        Arc::new(Self {
            generation,
            config,
            next_id: AtomicU64::new(1),
            state: Mutex::new(ControllerState {
                active: true,
                pending: VecDeque::new(),
                registry: HashMap::new(),
                running: 0,
            }),
            idle: Condvar::new(),
            diagnostics: Counters::default(),
            #[cfg(test)]
            before_start: Mutex::new(None),
            #[cfg(test)]
            panic_after_start: AtomicBool::new(false),
            #[cfg(test)]
            clear_operation_after_start: AtomicBool::new(false),
        })
    }

    fn lock_state(&self) -> std::sync::MutexGuard<'_, ControllerState> {
        self.state
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner)
    }

    fn enqueue(
        self: &Arc<Self>,
        operation: DispatchOperation,
    ) -> Result<DispatchTicket, DispatchEnqueueError> {
        let mut state = self.lock_state();
        if !state.active {
            self.diagnostics.rejected.fetch_add(1, Ordering::Relaxed);
            emit_request(EVENT_ENQUEUE_REJECTED, self.generation);
            return Err(DispatchEnqueueError::RuntimeClosing);
        }
        self.expire_locked(&mut state);
        if state.registry.len() >= self.config.maximum_pending {
            self.diagnostics.rejected.fetch_add(1, Ordering::Relaxed);
            emit_request(EVENT_ENQUEUE_REJECTED, self.generation);
            return Err(DispatchEnqueueError::QueueFull);
        }
        let id = self.next_id.fetch_add(1, Ordering::Relaxed);
        let request = Arc::new(Request {
            id,
            generation: self.generation,
            requirement: operation.requirement(),
            operation: Mutex::new(Some(operation)),
            state: AtomicU8::new(QUEUED),
            retired: AtomicBool::new(false),
            completion: Mutex::new(Completion { result: None }),
            ready: Condvar::new(),
            deadline: self
                .config
                .default_timeout
                .and_then(|timeout| Instant::now().checked_add(timeout)),
            owner: Arc::downgrade(self),
            #[cfg(test)]
            wait_returns: AtomicU64::new(0),
        });
        state.pending.push_back(request.clone());
        state.registry.insert(id, request.clone());
        drop(state);
        self.diagnostics.accepted.fetch_add(1, Ordering::Relaxed);
        emit_request(EVENT_ENQUEUE_ACCEPTED, correlation(self.generation, id));
        Ok(DispatchTicket { request })
    }

    fn expire_locked(&self, state: &mut ControllerState) -> usize {
        let now = Instant::now();
        let mut expired = Vec::new();
        state.pending.retain(|request| {
            let due = request.deadline.is_some_and(|deadline| now >= deadline);
            if due
                && request
                    .state
                    .compare_exchange(QUEUED, EXPIRED, Ordering::AcqRel, Ordering::Acquire)
                    .is_ok()
            {
                expired.push(request.clone());
                false
            } else {
                true
            }
        });
        let expired_count = expired.len();
        for request in expired {
            if request.publish(EXPIRED, Err(DispatchCompletionError::Expired)) {
                self.diagnostics.expired.fetch_add(1, Ordering::Relaxed);
                emit_request(EVENT_EXPIRED, correlation(self.generation, request.id));
            }
            self.retire_locked(state, request.id, Arc::as_ptr(&request));
            request.retired.store(true, Ordering::Release);
        }
        expired_count
    }

    fn expire(&self, request: &Request) {
        let mut state = self.lock_state();
        if request
            .state
            .compare_exchange(QUEUED, EXPIRED, Ordering::AcqRel, Ordering::Acquire)
            .is_err()
        {
            return;
        }
        state
            .pending
            .retain(|pending| !core::ptr::eq(Arc::as_ptr(pending), request));
        if request.publish(EXPIRED, Err(DispatchCompletionError::Expired)) {
            self.diagnostics.expired.fetch_add(1, Ordering::Relaxed);
            emit_request(EVENT_EXPIRED, correlation(self.generation, request.id));
        }
        self.retire_locked(&mut state, request.id, request as *const Request);
        request.retired.store(true, Ordering::Release);
    }

    fn cancel(&self, request: &Request) -> DispatchCancelOutcome {
        let mut state = self.lock_state();
        let current = request.state.load(Ordering::Acquire);
        if matches!(current, COMPLETED | FAILED | CANCELED | EXPIRED | SHUTDOWN) {
            return DispatchCancelOutcome::AlreadyTerminal;
        }
        if current == RUNNING {
            return DispatchCancelOutcome::TooLate;
        }
        if request
            .state
            .compare_exchange(current, CANCELED, Ordering::AcqRel, Ordering::Acquire)
            .is_err()
        {
            drop(state);
            return self.cancel(request);
        }
        state
            .pending
            .retain(|pending| !core::ptr::eq(Arc::as_ptr(pending), request));
        if request.publish(CANCELED, Err(DispatchCompletionError::Canceled)) {
            self.diagnostics.canceled.fetch_add(1, Ordering::Relaxed);
            emit_request(EVENT_CANCELED, correlation(self.generation, request.id));
        }
        self.retire_locked(&mut state, request.id, request as *const Request);
        request.retired.store(true, Ordering::Release);
        DispatchCancelOutcome::Canceled
    }

    fn select(&self, kind: DispatchCallbackKind) -> (Vec<Arc<Request>>, usize, usize) {
        let mut state = self.lock_state();
        if !state.active {
            return (Vec::new(), 0, 0);
        }
        let mut expired = self.expire_locked(&mut state);
        let mut selected = Vec::with_capacity(self.config.maximum_batch_per_drain);
        let mut retained = VecDeque::with_capacity(state.pending.len());
        let mut incompatible = 0;
        while let Some(request) = state.pending.pop_front() {
            if request.state.load(Ordering::Acquire) == EXPIRED {
                expired += 1;
                continue;
            }
            if selected.len() < self.config.maximum_batch_per_drain
                && kind.permits(request.requirement)
                && request
                    .state
                    .compare_exchange(QUEUED, SELECTED, Ordering::AcqRel, Ordering::Acquire)
                    .is_ok()
            {
                self.diagnostics.selected.fetch_add(1, Ordering::Relaxed);
                emit_request(EVENT_SELECTED, correlation(self.generation, request.id));
                selected.push(request);
            } else {
                if !kind.permits(request.requirement) {
                    incompatible += 1;
                }
                retained.push_back(request);
            }
        }
        state.pending = retained;
        if incompatible > 0 {
            self.diagnostics
                .incompatible_skipped
                .fetch_add(incompatible as u64, Ordering::Relaxed);
            emit_request(EVENT_INCOMPATIBLE_SKIPPED, self.generation);
        }
        (selected, incompatible, expired)
    }

    fn try_start(self: &Arc<Self>, request: Arc<Request>) -> Option<RunningRequestGuard> {
        let mut state = self.lock_state();
        let registered = state
            .registry
            .get(&request.id)
            .is_some_and(|entry| core::ptr::eq(Arc::as_ptr(entry), Arc::as_ptr(&request)));
        if !state.active || !registered || request.generation != self.generation {
            return None;
        }
        if request
            .state
            .compare_exchange(SELECTED, RUNNING, Ordering::AcqRel, Ordering::Acquire)
            .is_ok()
        {
            state.running += 1;
            Some(RunningRequestGuard {
                controller: self.clone(),
                request,
                finished: false,
            })
        } else {
            None
        }
    }

    fn finish_running(
        &self,
        request: &Request,
        result: Result<DispatchResult, DispatchExecutionError>,
    ) {
        if request.retired.swap(true, Ordering::AcqRel) {
            return;
        }
        let (state_code, completion) = match result {
            Ok(result) => (COMPLETED, Ok(result)),
            Err(error) => (FAILED, Err(DispatchCompletionError::Operation(error))),
        };
        let published = request.publish(state_code, completion);
        if published {
            let (counter, event) = if state_code == COMPLETED {
                (&self.diagnostics.completed, EVENT_COMPLETED)
            } else {
                (&self.diagnostics.failed, EVENT_FAILED)
            };
            counter.fetch_add(1, Ordering::Relaxed);
            emit_request(event, correlation(self.generation, request.id));
        }
        let mut state = self.lock_state();
        if state.running > 0 {
            state.running -= 1;
        } else {
            emit_request(
                EVENT_INTERNAL_INVARIANT,
                correlation(self.generation, request.id),
            );
        }
        self.retire_locked(&mut state, request.id, request as *const Request);
        if state.running == 0 {
            self.idle.notify_all();
        }
    }

    fn requeue_selected(&self, requests: &[Arc<Request>]) {
        let mut state = self.lock_state();
        if !state.active {
            drop(state);
            for request in requests {
                self.shutdown_selected(request);
            }
            return;
        }
        for request in requests.iter().rev() {
            if request
                .state
                .compare_exchange(SELECTED, QUEUED, Ordering::AcqRel, Ordering::Acquire)
                .is_ok()
            {
                state.pending.push_front(request.clone());
            }
        }
    }

    fn shutdown_selected(&self, request: &Request) {
        if request
            .state
            .compare_exchange(SELECTED, SHUTDOWN, Ordering::AcqRel, Ordering::Acquire)
            .is_ok()
        {
            if request.publish(SHUTDOWN, Err(DispatchCompletionError::DispatcherShutdown)) {
                self.diagnostics
                    .shutdown_retired
                    .fetch_add(1, Ordering::Relaxed);
                emit_request(EVENT_SHUTDOWN, correlation(self.generation, request.id));
            }
            request.retire();
        }
    }

    fn retire(&self, id: u64, expected: *const Request) {
        let mut state = self.lock_state();
        self.retire_locked(&mut state, id, expected);
    }

    fn retire_locked(&self, state: &mut ControllerState, id: u64, expected: *const Request) {
        if state
            .registry
            .get(&id)
            .is_some_and(|request| Arc::as_ptr(request) == expected)
        {
            state.registry.remove(&id);
        }
    }

    fn shutdown(&self) {
        let requests = {
            let mut state = self.lock_state();
            if !state.active {
                while state.running > 0 {
                    state = self
                        .idle
                        .wait(state)
                        .unwrap_or_else(std::sync::PoisonError::into_inner);
                }
                return;
            }
            self.expire_locked(&mut state);
            state.active = false;
            state.pending.clear();
            state.registry.values().cloned().collect::<Vec<_>>()
        };
        for request in &requests {
            let state = request.state.load(Ordering::Acquire);
            if matches!(state, QUEUED | SELECTED)
                && request
                    .state
                    .compare_exchange(state, SHUTDOWN, Ordering::AcqRel, Ordering::Acquire)
                    .is_ok()
            {
                if request.publish(SHUTDOWN, Err(DispatchCompletionError::DispatcherShutdown)) {
                    self.diagnostics
                        .shutdown_retired
                        .fetch_add(1, Ordering::Relaxed);
                    emit_request(EVENT_SHUTDOWN, correlation(self.generation, request.id));
                }
                request.retire();
            }
        }
        let mut state = self.lock_state();
        while state.running > 0 {
            state = self
                .idle
                .wait(state)
                .unwrap_or_else(std::sync::PoisonError::into_inner);
        }
        for request in requests {
            if !request.retired.load(Ordering::Acquire)
                && request.publish(SHUTDOWN, Err(DispatchCompletionError::DispatcherShutdown))
            {
                self.diagnostics
                    .shutdown_retired
                    .fetch_add(1, Ordering::Relaxed);
            }
            self.retire_locked(&mut state, request.id, Arc::as_ptr(&request));
            request.retired.store(true, Ordering::Release);
        }
        state.pending.clear();
        #[cfg(test)]
        debug_assert!(
            state.registry.is_empty(),
            "dispatcher request leaked during shutdown"
        );
    }

    fn diagnostics(&self) -> DispatchDiagnostics {
        let state = self.lock_state();
        DispatchDiagnostics {
            generation: self.generation,
            pending: state.pending.len(),
            running: state.running,
            accepted: self.diagnostics.accepted.load(Ordering::Relaxed),
            rejected: self.diagnostics.rejected.load(Ordering::Relaxed),
            selected: self.diagnostics.selected.load(Ordering::Relaxed),
            executed: self.diagnostics.executed.load(Ordering::Relaxed),
            canceled: self.diagnostics.canceled.load(Ordering::Relaxed),
            expired: self.diagnostics.expired.load(Ordering::Relaxed),
            completed: self.diagnostics.completed.load(Ordering::Relaxed),
            failed: self.diagnostics.failed.load(Ordering::Relaxed),
            shutdown_retired: self.diagnostics.shutdown_retired.load(Ordering::Relaxed),
            incompatible_skipped: self
                .diagnostics
                .incompatible_skipped
                .load(Ordering::Relaxed),
            nested_suppressed: self.diagnostics.nested_suppressed.load(Ordering::Relaxed),
        }
    }
}

/// Owns exact-once accounting from the Selected-to-Running commitment until
/// terminal publication and retirement.
struct RunningRequestGuard {
    controller: Arc<DispatchController>,
    request: Arc<Request>,
    finished: bool,
}

impl RunningRequestGuard {
    fn take_operation(&self) -> Option<DispatchOperation> {
        self.request
            .operation
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner)
            .take()
    }

    fn finish(mut self, result: Result<DispatchResult, DispatchExecutionError>) {
        self.controller.finish_running(&self.request, result);
        self.finished = true;
    }
}

impl Drop for RunningRequestGuard {
    fn drop(&mut self) {
        if self.finished {
            return;
        }
        let _ = self
            .request
            .operation
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner)
            .take();
        self.controller
            .finish_running(&self.request, Err(DispatchExecutionError::Panicked));
        self.finished = true;
    }
}

struct PendingConfig(DispatchConfig);

#[derive(Default)]
struct Generations {
    pending: Option<PendingConfig>,
    active: Option<Arc<DispatchController>>,
    next_generation: u64,
}

fn generations() -> &'static Mutex<Generations> {
    static GENERATIONS: OnceLock<Mutex<Generations>> = OnceLock::new();
    GENERATIONS.get_or_init(|| Mutex::new(Generations::default()))
}

fn lock_generations() -> std::sync::MutexGuard<'static, Generations> {
    generations()
        .lock()
        .unwrap_or_else(std::sync::PoisonError::into_inner)
}

pub(crate) fn install_production_config(config: DispatchConfig) -> Result<(), DispatchConfig> {
    let Some(config) = config.normalized() else {
        return Err(config);
    };
    let mut generations = lock_generations();
    if generations.pending.is_some() || generations.active.is_some() {
        return Err(config);
    }
    generations.pending = Some(PendingConfig(config));
    Ok(())
}

pub(crate) fn activate() -> bool {
    let mut generations = lock_generations();
    if generations.active.is_some() {
        return true;
    }
    let Some(PendingConfig(config)) = generations.pending.take() else {
        return false;
    };
    generations.next_generation = generations.next_generation.wrapping_add(1).max(1);
    generations.active = Some(DispatchController::new(generations.next_generation, config));
    true
}

pub(crate) fn shutdown() {
    let controller = lock_generations().active.take();
    if let Some(controller) = controller {
        controller.shutdown();
    }
}

pub fn current_generation() -> Option<DispatchGeneration> {
    lock_generations()
        .active
        .clone()
        .map(|controller| DispatchGeneration { controller })
}

fn active_generation_id() -> Option<u64> {
    lock_generations()
        .active
        .as_ref()
        .map(|controller| controller.generation)
}

pub fn enqueue(operation: DispatchOperation) -> Result<DispatchTicket, DispatchEnqueueError> {
    current_generation()
        .ok_or(DispatchEnqueueError::NoActiveGeneration)?
        .enqueue(operation)
}

pub fn diagnostics() -> Option<DispatchDiagnostics> {
    current_generation().map(|generation| generation.diagnostics())
}

thread_local! {
    static CALLBACK_DEPTH: Cell<usize> = const { Cell::new(0) };
    static DRAIN_DEPTH: Cell<usize> = const { Cell::new(0) };
}

fn callback_depth() -> usize {
    CALLBACK_DEPTH.get()
}

struct DrainGuard {
    entered: bool,
    _callback: CallbackGuard,
    _not_send: PhantomData<*mut ()>,
}

impl DrainGuard {
    fn enter() -> Self {
        let entered = DRAIN_DEPTH.get() == 0;
        if entered {
            DRAIN_DEPTH.set(1);
        }
        Self {
            entered,
            _callback: enter_callback(),
            _not_send: PhantomData,
        }
    }
}

impl Drop for DrainGuard {
    fn drop(&mut self) {
        if self.entered {
            DRAIN_DEPTH.set(0);
        }
    }
}

pub(crate) struct CallbackGuard {
    _not_send: PhantomData<*mut ()>,
}

pub(crate) fn enter_callback() -> CallbackGuard {
    CALLBACK_DEPTH.set(CALLBACK_DEPTH.get().saturating_add(1));
    CallbackGuard {
        _not_send: PhantomData,
    }
}

impl Drop for CallbackGuard {
    fn drop(&mut self) {
        CALLBACK_DEPTH.set(CALLBACK_DEPTH.get().saturating_sub(1));
    }
}

enum DrainContext<'call, 'context> {
    ThreadSafe(&'context ThreadSafeContext<'call>),
    Worksheet(&'context WorksheetContext<'call>),
    Macro(&'context MacroContext<'call>),
    Lifecycle(&'context LifecycleContext<'call>),
}

impl DrainContext<'_, '_> {
    const fn kind(&self) -> DispatchCallbackKind {
        match self {
            Self::ThreadSafe(context) => {
                let _ = context;
                DispatchCallbackKind::ThreadSafe
            }
            Self::Worksheet(context) => {
                let _ = context;
                DispatchCallbackKind::Worksheet
            }
            Self::Macro(_) => DispatchCallbackKind::Macro,
            Self::Lifecycle(context) => {
                let _ = context;
                DispatchCallbackKind::Lifecycle
            }
        }
    }
}

fn drain(context: DrainContext<'_, '_>) -> DispatchDrainReport {
    let Some(generation) = current_generation() else {
        return DispatchDrainReport::default();
    };
    let guard = DrainGuard::enter();
    if !guard.entered {
        generation
            .controller
            .diagnostics
            .nested_suppressed
            .fetch_add(1, Ordering::Relaxed);
        emit_request(EVENT_NESTED_SUPPRESSED, generation.id());
        return DispatchDrainReport {
            nested_suppressed: true,
            ..DispatchDrainReport::default()
        };
    }
    let kind = context.kind();
    let (selected, incompatible, expired) = generation.controller.select(kind);
    let mut report = DispatchDrainReport {
        selected: selected.len(),
        incompatible_skipped: incompatible,
        expired,
        ..DispatchDrainReport::default()
    };
    let started = Instant::now();
    for (index, request) in selected.iter().enumerate() {
        if generation
            .controller
            .config
            .maximum_drain_duration
            .is_some_and(|limit| started.elapsed() >= limit)
        {
            generation.controller.requeue_selected(&selected[index..]);
            report.duration_limit_reached = true;
            break;
        }
        #[cfg(test)]
        if let Some(pause) = generation
            .controller
            .before_start
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner)
            .clone()
        {
            pause.reached.wait();
            pause.release.wait();
        }
        let Some(running) = generation.controller.try_start(request.clone()) else {
            generation.controller.shutdown_selected(request);
            continue;
        };
        #[cfg(test)]
        if generation
            .controller
            .panic_after_start
            .swap(false, Ordering::AcqRel)
        {
            panic!("injected panic after dispatcher Running commitment");
        }
        #[cfg(test)]
        if generation
            .controller
            .clear_operation_after_start
            .swap(false, Ordering::AcqRel)
        {
            let _ = running.take_operation();
        }
        let Some(operation) = running.take_operation() else {
            emit_request(
                EVENT_INTERNAL_INVARIANT,
                correlation(request.generation, request.id),
            );
            running.finish(Err(DispatchExecutionError::InternalInvariant));
            report.failed += 1;
            report.processed += 1;
            continue;
        };
        generation
            .controller
            .diagnostics
            .executed
            .fetch_add(1, Ordering::Relaxed);
        emit_request(EVENT_EXECUTED, correlation(request.generation, request.id));
        let result = std::panic::catch_unwind(AssertUnwindSafe(|| execute(operation, &context)))
            .unwrap_or(Err(DispatchExecutionError::Panicked));
        if result.is_err() {
            report.failed += 1;
        }
        running.finish(result);
        report.processed += 1;
    }
    report
}

fn execute(
    operation: DispatchOperation,
    context: &DrainContext<'_, '_>,
) -> Result<DispatchResult, DispatchExecutionError> {
    if !context.kind().permits(operation.requirement()) {
        return Err(DispatchExecutionError::IncompatibleContext {
            required: operation.requirement(),
            actual: context.kind(),
        });
    }
    match operation {
        DispatchOperation::EchoOwned(value) => Ok(DispatchResult::OwnedValue(value)),
        DispatchOperation::PollCancellationPreservingBreak => match context {
            DrainContext::Macro(context) => context
                .is_cancellation_requested()
                .map(DispatchResult::CancellationRequested)
                .map_err(DispatchExecutionError::Excel),
            _ => Err(DispatchExecutionError::IncompatibleContext {
                required: DispatchRequirement::Macro,
                actual: context.kind(),
            }),
        },
        #[cfg(test)]
        DispatchOperation::TestEnqueueNested(value) => {
            enqueue(DispatchOperation::EchoOwned(value.clone()))
                .map_err(|_| DispatchExecutionError::Panicked)?;
            Ok(DispatchResult::OwnedValue(value))
        }
        #[cfg(test)]
        DispatchOperation::TestNestedMacroDrain => match context {
            DrainContext::Macro(context) => {
                let report = drain(DrainContext::Macro(context));
                Ok(DispatchResult::CancellationRequested(
                    report.nested_suppressed,
                ))
            }
            _ => unreachable!("selection enforces macro compatibility"),
        },
        #[cfg(test)]
        DispatchOperation::TestBlock => {
            if let Some(pause) = test_execution_pause()
                .lock()
                .unwrap_or_else(std::sync::PoisonError::into_inner)
                .clone()
            {
                pause.reached.wait();
                pause.release.wait();
            }
            Ok(DispatchResult::OwnedValue(ExcelValue::Integer(1)))
        }
        #[cfg(test)]
        DispatchOperation::TestPanic => panic!("dispatcher test panic"),
    }
}

pub(crate) fn drain_thread_safe(context: &ThreadSafeContext<'_>) -> DispatchDrainReport {
    drain(DrainContext::ThreadSafe(context))
}

pub(crate) fn drain_worksheet(context: &WorksheetContext<'_>) -> DispatchDrainReport {
    drain(DrainContext::Worksheet(context))
}

pub(crate) fn drain_macro(context: &MacroContext<'_>) -> DispatchDrainReport {
    drain(DrainContext::Macro(context))
}

pub(crate) fn drain_lifecycle(context: &LifecycleContext<'_>) -> DispatchDrainReport {
    drain(DrainContext::Lifecycle(context))
}

fn correlation(generation: u64, request: u64) -> u64 {
    generation.rotate_left(32) ^ request
}

fn emit_request(event: i32, correlation_id: u64) {
    crate::diagnostics::emit(
        DiagnosticEvent::new(DiagnosticCode::Dispatcher, DiagnosticSeverity::Info, event)
            .with_correlation(correlation_id),
    );
}

#[cfg(test)]
pub(crate) static TEST_SERIAL: Mutex<()> = Mutex::new(());

#[cfg(test)]
struct TestPause {
    reached: std::sync::Barrier,
    release: std::sync::Barrier,
}

#[cfg(test)]
fn test_execution_pause() -> &'static Mutex<Option<Arc<TestPause>>> {
    static PAUSE: OnceLock<Mutex<Option<Arc<TestPause>>>> = OnceLock::new();
    PAUSE.get_or_init(|| Mutex::new(None))
}

#[cfg(test)]
#[allow(
    dead_code,
    reason = "used by dispatcher/runtime test modules as they are added"
)]
pub(crate) fn reset_generations_for_test() {
    let active = {
        let mut generations = lock_generations();
        generations.pending = None;
        generations.active.take()
    };
    if let Some(active) = active {
        active.shutdown();
    }
    CALLBACK_DEPTH.set(0);
    DRAIN_DEPTH.set(0);
    *test_execution_pause()
        .lock()
        .unwrap_or_else(std::sync::PoisonError::into_inner) = None;
}

#[cfg(test)]
mod tests {
    use super::*;
    use core::ffi::c_int;
    use std::sync::{Barrier, mpsc};

    use excel_api_sys::{LPXLOPER12, XLOPER12, XLOPER12Value};

    use crate::excel_call::{CallCapability, ExcelCallBackend};

    #[derive(Default)]
    struct AbortBackend {
        linked: AtomicBool,
        calls: Mutex<Vec<(i32, usize)>>,
        code: std::sync::atomic::AtomicI32,
        result: AtomicBool,
    }

    impl AbortBackend {
        fn linked() -> Self {
            Self {
                linked: AtomicBool::new(true),
                result: AtomicBool::new(true),
                ..Self::default()
            }
        }
    }

    impl ExcelCallBackend for AbortBackend {
        fn link(&self) -> Result<(), ExcelCallError> {
            self.linked.store(true, Ordering::Release);
            Ok(())
        }

        fn unlink(&self) {
            self.linked.store(false, Ordering::Release);
        }

        fn is_linked(&self) -> bool {
            self.linked.load(Ordering::Acquire)
        }

        unsafe fn excel12v_raw(
            &self,
            function: i32,
            result: *mut XLOPER12,
            count: c_int,
            _: *mut LPXLOPER12,
        ) -> i32 {
            self.calls.lock().unwrap().push((function, count as usize));
            // SAFETY: the xlAbort wrapper supplies one live immediate result root.
            unsafe {
                (*result).val = XLOPER12Value {
                    xbool: i32::from(self.result.load(Ordering::Acquire)),
                };
                (*result).xltype = excel_api_sys::xltypeBool;
            }
            self.code.load(Ordering::Acquire)
        }
    }

    struct GenerationGuard;

    impl Drop for GenerationGuard {
        fn drop(&mut self) {
            reset_generations_for_test();
        }
    }

    fn start(config: DispatchConfig) -> (GenerationGuard, DispatchGeneration) {
        reset_generations_for_test();
        assert_eq!(install_production_config(config), Ok(()));
        assert!(activate());
        (GenerationGuard, current_generation().unwrap())
    }

    fn value(number: i32) -> DispatchOperation {
        DispatchOperation::EchoOwned(ExcelValue::Integer(number))
    }

    fn assert_owned(ticket: &DispatchTicket, expected: i32) {
        assert_eq!(
            ticket.try_result(),
            Some(Ok(DispatchResult::OwnedValue(ExcelValue::Integer(
                expected
            ))))
        );
    }

    fn clone_ticket(ticket: &DispatchTicket) -> DispatchTicket {
        DispatchTicket {
            request: ticket.request.clone(),
        }
    }

    #[test]
    fn compatibility_table_is_explicit_and_not_a_main_thread_ordering() {
        let _serial = TEST_SERIAL.lock().unwrap();
        for kind in [
            DispatchCallbackKind::ThreadSafe,
            DispatchCallbackKind::Worksheet,
            DispatchCallbackKind::Macro,
            DispatchCallbackKind::Lifecycle,
        ] {
            assert!(kind.permits(DispatchRequirement::ContextNeutral));
        }
        assert!(DispatchCallbackKind::ThreadSafe.permits(DispatchRequirement::ThreadSafeWorksheet));
        assert!(!DispatchCallbackKind::ThreadSafe.permits(DispatchRequirement::Worksheet));
        assert!(!DispatchCallbackKind::ThreadSafe.permits(DispatchRequirement::Macro));
        assert!(!DispatchCallbackKind::ThreadSafe.permits(DispatchRequirement::Lifecycle));
        assert!(DispatchCallbackKind::Worksheet.permits(DispatchRequirement::Worksheet));
        assert!(!DispatchCallbackKind::Worksheet.permits(DispatchRequirement::Macro));
        assert!(DispatchCallbackKind::Macro.permits(DispatchRequirement::Macro));
        assert!(!DispatchCallbackKind::Macro.permits(DispatchRequirement::Worksheet));
        assert!(DispatchCallbackKind::Lifecycle.permits(DispatchRequirement::Lifecycle));
        assert!(!DispatchCallbackKind::Lifecycle.permits(DispatchRequirement::Macro));
    }

    #[test]
    fn empty_queue_fifo_and_bounded_batch_are_exact() {
        let _serial = TEST_SERIAL.lock().unwrap();
        let (_guard, _) = start(DispatchConfig {
            maximum_batch_per_drain: 2,
            ..DispatchConfig::default()
        });
        let backend = AbortBackend::linked();
        let capability = CallCapability::new(&backend);
        let context = ThreadSafeContext::new(&capability);
        assert_eq!(context.drain_dispatcher(), DispatchDrainReport::default());
        let first = enqueue(value(1)).unwrap();
        let second = enqueue(value(2)).unwrap();
        let third = enqueue(value(3)).unwrap();
        let report = context.drain_dispatcher();
        assert_eq!((report.selected, report.processed), (2, 2));
        assert_owned(&first, 1);
        assert_owned(&second, 2);
        assert_eq!(third.try_result(), None);
        assert_eq!(context.drain_dispatcher().processed, 1);
        assert_owned(&third, 3);
        assert!(backend.calls.lock().unwrap().is_empty());
    }

    #[test]
    fn queue_full_and_detached_ticket_policy_remain_bounded() {
        let _serial = TEST_SERIAL.lock().unwrap();
        let (_guard, generation) = start(DispatchConfig {
            maximum_pending: 1,
            ..DispatchConfig::default()
        });
        let detached = enqueue(value(1)).unwrap();
        drop(detached);
        assert_eq!(
            enqueue(value(2)).unwrap_err(),
            DispatchEnqueueError::QueueFull
        );
        assert_eq!(generation.diagnostics().pending, 1);
        let backend = AbortBackend::linked();
        let capability = CallCapability::new(&backend);
        ThreadSafeContext::new(&capability).drain_dispatcher();
        let diagnostics = generation.diagnostics();
        assert_eq!((diagnostics.pending, diagnostics.completed), (0, 1));
    }

    #[test]
    fn compatible_request_behind_macro_head_executes_without_excel_call() {
        let _serial = TEST_SERIAL.lock().unwrap();
        let (_guard, generation) = start(DispatchConfig::default());
        let macro_ticket = enqueue(DispatchOperation::PollCancellationPreservingBreak).unwrap();
        let pure_ticket = enqueue(value(9)).unwrap();
        let backend = AbortBackend::linked();
        let capability = CallCapability::new(&backend);
        let thread_safe = ThreadSafeContext::new(&capability);
        let report = thread_safe.drain_dispatcher();
        assert_eq!((report.processed, report.incompatible_skipped), (1, 1));
        assert_owned(&pure_ticket, 9);
        assert_eq!(macro_ticket.try_result(), None);
        assert!(backend.calls.lock().unwrap().is_empty());
        let macro_context = MacroContext::new(&capability);
        assert_eq!(macro_context.drain_dispatcher().processed, 1);
        assert_eq!(
            macro_ticket.try_result(),
            Some(Ok(DispatchResult::CancellationRequested(true)))
        );
        assert_eq!(
            *backend.calls.lock().unwrap(),
            vec![(excel_api_sys::xlAbort, 0)]
        );
        assert_eq!(generation.diagnostics().incompatible_skipped, 1);
    }

    #[test]
    fn worksheet_and_lifecycle_contexts_execute_only_approved_work() {
        let _serial = TEST_SERIAL.lock().unwrap();
        let (_guard, _) = start(DispatchConfig::default());
        let macro_ticket = enqueue(DispatchOperation::PollCancellationPreservingBreak).unwrap();
        let worksheet_ticket = enqueue(value(1)).unwrap();
        let backend = AbortBackend::linked();
        let capability = CallCapability::new(&backend);
        let worksheet = WorksheetContext::new(&capability);
        assert_eq!(worksheet.drain_dispatcher().processed, 1);
        assert_owned(&worksheet_ticket, 1);
        assert_eq!(macro_ticket.try_result(), None);
        let lifecycle_ticket = enqueue(value(2)).unwrap();
        let lifecycle = LifecycleContext::new(&capability);
        assert_eq!(lifecycle.drain_dispatcher().processed, 1);
        assert_owned(&lifecycle_ticket, 2);
        assert_eq!(macro_ticket.try_result(), None);
        assert!(backend.calls.lock().unwrap().is_empty());
    }

    #[test]
    fn expiration_and_wait_timeout_are_distinct() {
        let _serial = TEST_SERIAL.lock().unwrap();
        let (_guard, _) = start(DispatchConfig {
            default_timeout: Some(Duration::ZERO),
            ..DispatchConfig::default()
        });
        let expired = enqueue(value(1)).unwrap();
        assert_eq!(
            expired.try_result(),
            Some(Err(DispatchCompletionError::Expired))
        );
        reset_generations_for_test();
        assert_eq!(install_production_config(DispatchConfig::default()), Ok(()));
        assert!(activate());
        let queued = enqueue(value(2)).unwrap();
        assert_eq!(
            queued.wait_timeout(Duration::ZERO),
            Err(DispatchCompletionError::WaitTimeout)
        );
        assert_eq!(queued.try_result(), None);

        reset_generations_for_test();
        assert_eq!(
            install_production_config(DispatchConfig {
                default_timeout: Some(Duration::ZERO),
                ..DispatchConfig::default()
            }),
            Ok(())
        );
        assert!(activate());
        let expired_during_shutdown = enqueue(value(3)).unwrap();
        shutdown();
        assert_eq!(
            expired_during_shutdown.try_result(),
            Some(Err(DispatchCompletionError::Expired))
        );
    }

    #[test]
    fn wait_observes_request_deadline_before_longer_caller_timeout() {
        let _serial = TEST_SERIAL.lock().unwrap();
        let (_guard, generation) = start(DispatchConfig {
            default_timeout: Some(Duration::from_millis(40)),
            ..DispatchConfig::default()
        });
        let ticket = enqueue(value(1)).unwrap();
        let started = Instant::now();
        assert_eq!(
            ticket.wait_timeout(Duration::from_secs(2)),
            Err(DispatchCompletionError::Expired)
        );
        assert!(started.elapsed() >= Duration::from_millis(20));
        assert!(started.elapsed() < Duration::from_secs(1));
        assert_eq!(generation.diagnostics().pending, 0);
    }

    #[test]
    fn caller_timeout_does_not_expire_a_longer_lived_request() {
        let _serial = TEST_SERIAL.lock().unwrap();
        let (_guard, generation) = start(DispatchConfig {
            default_timeout: Some(Duration::from_secs(2)),
            ..DispatchConfig::default()
        });
        let ticket = enqueue(value(1)).unwrap();
        assert_eq!(
            ticket.wait_timeout(Duration::from_millis(20)),
            Err(DispatchCompletionError::WaitTimeout)
        );
        assert_eq!(ticket.try_result(), None);
        assert_eq!(generation.diagnostics().pending, 1);
    }

    #[test]
    fn wait_returns_completion_cancellation_and_shutdown_results() {
        let _serial = TEST_SERIAL.lock().unwrap();
        let (_guard, _) = start(DispatchConfig::default());
        let backend = AbortBackend::linked();
        let capability = CallCapability::new(&backend);

        let completed = enqueue(value(7)).unwrap();
        let completed_wait = clone_ticket(&completed);
        let (ready_tx, ready_rx) = mpsc::channel();
        let waiter = std::thread::spawn(move || {
            ready_tx.send(()).unwrap();
            completed_wait.wait_timeout(Duration::from_secs(2))
        });
        ready_rx.recv().unwrap();
        MacroContext::new(&capability).drain_dispatcher();
        assert_eq!(
            waiter.join().unwrap(),
            Ok(DispatchResult::OwnedValue(ExcelValue::Integer(7)))
        );

        let canceled = enqueue(value(8)).unwrap();
        let canceled_wait = clone_ticket(&canceled);
        let waiter = std::thread::spawn(move || canceled_wait.wait_timeout(Duration::from_secs(2)));
        assert_eq!(canceled.cancel(), DispatchCancelOutcome::Canceled);
        assert_eq!(
            waiter.join().unwrap(),
            Err(DispatchCompletionError::Canceled)
        );

        let stopped = enqueue(value(9)).unwrap();
        let stopped_wait = clone_ticket(&stopped);
        let waiter = std::thread::spawn(move || stopped_wait.wait_timeout(Duration::from_secs(2)));
        shutdown();
        assert_eq!(
            waiter.join().unwrap(),
            Err(DispatchCompletionError::DispatcherShutdown)
        );
    }

    #[test]
    fn spurious_notifications_do_not_complete_a_wait() {
        let _serial = TEST_SERIAL.lock().unwrap();
        let (_guard, _) = start(DispatchConfig::default());
        let ticket = enqueue(value(1)).unwrap();
        let waiter_ticket = clone_ticket(&ticket);
        let request = ticket.request.clone();
        let waiter = std::thread::spawn(move || waiter_ticket.wait_timeout(Duration::from_secs(2)));
        while request.wait_returns.load(Ordering::Acquire) == 0 {
            request.ready.notify_all();
            std::thread::yield_now();
        }
        assert_eq!(ticket.try_result(), None);
        assert_eq!(ticket.cancel(), DispatchCancelOutcome::Canceled);
        assert_eq!(
            waiter.join().unwrap(),
            Err(DispatchCompletionError::Canceled)
        );
    }

    #[test]
    fn waiting_after_expiry_is_immediate_and_selected_or_running_work_does_not_expire() {
        let _serial = TEST_SERIAL.lock().unwrap();
        let (_guard, _generation) = start(DispatchConfig {
            default_timeout: Some(Duration::ZERO),
            ..DispatchConfig::default()
        });
        let expired = enqueue(value(1)).unwrap();
        assert_eq!(
            expired.wait_timeout(Duration::from_secs(1)),
            Err(DispatchCompletionError::Expired)
        );

        reset_generations_for_test();
        assert_eq!(
            install_production_config(DispatchConfig {
                default_timeout: Some(Duration::from_millis(5)),
                ..DispatchConfig::default()
            }),
            Ok(())
        );
        assert!(activate());
        let generation = current_generation().unwrap();
        let selected_ticket = enqueue(value(2)).unwrap();
        let (_selected, _, _) = generation.controller.select(DispatchCallbackKind::Macro);
        assert_eq!(
            selected_ticket.wait_timeout(Duration::from_millis(20)),
            Err(DispatchCompletionError::WaitTimeout)
        );
        assert_eq!(selected_ticket.try_result(), None);
        assert_eq!(selected_ticket.cancel(), DispatchCancelOutcome::Canceled);

        let running_ticket = enqueue(value(3)).unwrap();
        let (selected, _, _) = generation.controller.select(DispatchCallbackKind::Macro);
        let running = generation
            .controller
            .try_start(selected[0].clone())
            .unwrap();
        assert_eq!(
            running_ticket.wait_timeout(Duration::from_millis(20)),
            Err(DispatchCompletionError::WaitTimeout)
        );
        assert_eq!(running_ticket.try_result(), None);
        running.finish(Ok(DispatchResult::OwnedValue(ExcelValue::Integer(3))));
        assert_owned(&running_ticket, 3);
        assert_eq!(generation.diagnostics().running, 0);
    }

    #[test]
    fn cancel_before_and_after_selection_retires_exactly_once() {
        let _serial = TEST_SERIAL.lock().unwrap();
        let (_guard, generation) = start(DispatchConfig::default());
        let queued = enqueue(value(1)).unwrap();
        assert_eq!(queued.cancel(), DispatchCancelOutcome::Canceled);
        assert_eq!(queued.cancel(), DispatchCancelOutcome::AlreadyTerminal);
        assert_eq!(
            queued.try_result(),
            Some(Err(DispatchCompletionError::Canceled))
        );
        let selected_ticket = enqueue(value(2)).unwrap();
        let (selected, _, _) = generation.controller.select(DispatchCallbackKind::Macro);
        assert_eq!(selected.len(), 1);
        assert_eq!(selected_ticket.cancel(), DispatchCancelOutcome::Canceled);
        generation.controller.shutdown_selected(&selected[0]);
        assert_eq!(
            selected_ticket.try_result(),
            Some(Err(DispatchCompletionError::Canceled))
        );
        assert_eq!(generation.diagnostics().pending, 0);
    }

    #[test]
    fn cancellation_loses_after_execution_commitment() {
        let _serial = TEST_SERIAL.lock().unwrap();
        let (_guard, generation) = start(DispatchConfig::default());
        let ticket = enqueue(value(4)).unwrap();
        let (selected, _, _) = generation.controller.select(DispatchCallbackKind::Macro);
        let running = generation
            .controller
            .try_start(selected[0].clone())
            .unwrap();
        assert_eq!(ticket.cancel(), DispatchCancelOutcome::TooLate);
        running.finish(Ok(DispatchResult::OwnedValue(ExcelValue::Integer(4))));
        assert_owned(&ticket, 4);
        assert_eq!(generation.diagnostics().running, 0);
    }

    #[test]
    fn waits_are_rejected_under_callback_depth_guard() {
        let _serial = TEST_SERIAL.lock().unwrap();
        let (_guard, _) = start(DispatchConfig::default());
        let ticket = enqueue(value(1)).unwrap();
        let callback = DrainGuard::enter();
        assert_eq!(
            ticket.wait_timeout(Duration::from_secs(1)),
            Err(DispatchCompletionError::WaitFromCallback)
        );
        drop(callback);
        assert_eq!(callback_depth(), 0);
    }

    #[test]
    fn nested_enqueue_waits_for_later_drain_and_nested_drain_is_suppressed() {
        let _serial = TEST_SERIAL.lock().unwrap();
        let (_guard, generation) = start(DispatchConfig::default());
        let outer = enqueue(DispatchOperation::TestEnqueueNested(ExcelValue::Integer(8))).unwrap();
        let nested = enqueue(DispatchOperation::TestNestedMacroDrain).unwrap();
        let backend = AbortBackend::linked();
        let capability = CallCapability::new(&backend);
        let macro_context = MacroContext::new(&capability);
        let report = macro_context.drain_dispatcher();
        assert_eq!(report.processed, 2);
        assert_owned(&outer, 8);
        assert_eq!(
            nested.try_result(),
            Some(Ok(DispatchResult::CancellationRequested(true)))
        );
        assert_eq!(generation.diagnostics().pending, 1);
        assert_eq!(generation.diagnostics().nested_suppressed, 1);
        assert_eq!(macro_context.drain_dispatcher().processed, 1);
        assert_eq!(generation.diagnostics().pending, 0);
    }

    #[test]
    fn panic_clears_guard_and_later_drain_succeeds() {
        let _serial = TEST_SERIAL.lock().unwrap();
        let (_guard, _) = start(DispatchConfig::default());
        let panic_ticket = enqueue(DispatchOperation::TestPanic).unwrap();
        let backend = AbortBackend::linked();
        let capability = CallCapability::new(&backend);
        let macro_context = MacroContext::new(&capability);
        let report = macro_context.drain_dispatcher();
        assert_eq!((report.processed, report.failed), (1, 1));
        assert!(matches!(
            panic_ticket.try_result(),
            Some(Err(DispatchCompletionError::Operation(
                DispatchExecutionError::Panicked
            )))
        ));
        assert_eq!(callback_depth(), 0);
        let later = enqueue(value(3)).unwrap();
        assert_eq!(macro_context.drain_dispatcher().processed, 1);
        assert_owned(&later, 3);
    }

    #[test]
    fn running_guard_retires_repeated_pre_execution_panics_without_underflow() {
        let _serial = TEST_SERIAL.lock().unwrap();
        let (_guard, generation) = start(DispatchConfig::default());
        let backend = AbortBackend::linked();
        let capability = CallCapability::new(&backend);
        let macro_context = MacroContext::new(&capability);

        for expected_failures in 1..=3 {
            let ticket = enqueue(value(expected_failures)).unwrap();
            generation
                .controller
                .panic_after_start
                .store(true, Ordering::Release);
            assert!(
                std::panic::catch_unwind(AssertUnwindSafe(|| { macro_context.drain_dispatcher() }))
                    .is_err()
            );
            assert_eq!(
                ticket.try_result(),
                Some(Err(DispatchCompletionError::Operation(
                    DispatchExecutionError::Panicked
                )))
            );
            let diagnostics = generation.diagnostics();
            assert_eq!(diagnostics.running, 0);
            assert_eq!(diagnostics.pending, 0);
            assert_eq!(diagnostics.failed, expected_failures as u64);
            assert_eq!(callback_depth(), 0);
            assert!(ticket.request.operation.lock().unwrap().is_none());
            assert!(generation.controller.lock_state().registry.is_empty());
        }

        let later = enqueue(value(4)).unwrap();
        assert_eq!(macro_context.drain_dispatcher().processed, 1);
        assert_owned(&later, 4);

        let (done_tx, done_rx) = mpsc::channel();
        let closer = std::thread::spawn(move || {
            shutdown();
            done_tx.send(()).unwrap();
        });
        done_rx.recv_timeout(Duration::from_secs(2)).unwrap();
        closer.join().unwrap();
        assert_eq!(install_production_config(DispatchConfig::default()), Ok(()));
        assert!(activate());
        let reopened = enqueue(value(5)).unwrap();
        MacroContext::new(&capability).drain_dispatcher();
        assert_owned(&reopened, 5);
    }

    #[test]
    fn missing_running_operation_is_a_controlled_failure_and_shutdown_remains_live() {
        let _serial = TEST_SERIAL.lock().unwrap();
        let (_guard, generation) = start(DispatchConfig::default());
        let ticket = enqueue(value(1)).unwrap();
        generation
            .controller
            .clear_operation_after_start
            .store(true, Ordering::Release);
        let backend = AbortBackend::linked();
        let capability = CallCapability::new(&backend);
        let report = MacroContext::new(&capability).drain_dispatcher();
        assert_eq!((report.processed, report.failed), (1, 1));
        assert_eq!(
            ticket.try_result(),
            Some(Err(DispatchCompletionError::Operation(
                DispatchExecutionError::InternalInvariant
            )))
        );
        let diagnostics = generation.diagnostics();
        assert_eq!((diagnostics.pending, diagnostics.running), (0, 0));
        assert!(generation.controller.lock_state().registry.is_empty());

        let (done_tx, done_rx) = mpsc::channel();
        let closer = std::thread::spawn(move || {
            shutdown();
            done_tx.send(()).unwrap();
        });
        done_rx.recv_timeout(Duration::from_secs(2)).unwrap();
        closer.join().unwrap();
    }

    #[test]
    fn zero_duration_requeues_selected_work_without_execution() {
        let _serial = TEST_SERIAL.lock().unwrap();
        let (_guard, generation) = start(DispatchConfig {
            maximum_drain_duration: Some(Duration::ZERO),
            ..DispatchConfig::default()
        });
        let ticket = enqueue(value(1)).unwrap();
        let backend = AbortBackend::linked();
        let capability = CallCapability::new(&backend);
        let report = MacroContext::new(&capability).drain_dispatcher();
        assert!(report.duration_limit_reached);
        assert_eq!(report.processed, 0);
        assert_eq!(ticket.try_result(), None);
        assert_eq!(generation.diagnostics().pending, 1);
    }

    #[test]
    fn selected_request_cannot_start_after_shutdown_boundary() {
        let _serial = TEST_SERIAL.lock().unwrap();
        let (_guard, generation) = start(DispatchConfig::default());
        let ticket = enqueue(value(1)).unwrap();
        let pause = Arc::new(TestPause {
            reached: Barrier::new(2),
            release: Barrier::new(2),
        });
        *generation.controller.before_start.lock().unwrap() = Some(pause.clone());
        let backend = Arc::new(AbortBackend::linked());
        let drain = {
            let backend = backend.clone();
            std::thread::spawn(move || {
                let capability = CallCapability::new(backend.as_ref());
                MacroContext::new(&capability).drain_dispatcher()
            })
        };
        pause.reached.wait();
        shutdown();
        pause.release.wait();
        assert_eq!(drain.join().unwrap().processed, 0);
        assert_eq!(
            ticket.try_result(),
            Some(Err(DispatchCompletionError::DispatcherShutdown))
        );
    }

    #[test]
    fn shutdown_waits_for_running_operation_then_reopen_rejects_stale_generation() {
        let _serial = TEST_SERIAL.lock().unwrap();
        let (_guard, old) = start(DispatchConfig::default());
        let ticket = old.enqueue(DispatchOperation::TestBlock).unwrap();
        let pause = Arc::new(TestPause {
            reached: Barrier::new(2),
            release: Barrier::new(2),
        });
        *test_execution_pause().lock().unwrap() = Some(pause.clone());
        let backend = Arc::new(AbortBackend::linked());
        let drain = {
            let backend = backend.clone();
            std::thread::spawn(move || {
                let capability = CallCapability::new(backend.as_ref());
                MacroContext::new(&capability).drain_dispatcher()
            })
        };
        pause.reached.wait();
        let (done_tx, done_rx) = mpsc::channel();
        let closer = std::thread::spawn(move || {
            shutdown();
            done_tx.send(()).unwrap();
        });
        assert!(done_rx.try_recv().is_err());
        pause.release.wait();
        assert_eq!(drain.join().unwrap().processed, 1);
        done_rx.recv_timeout(Duration::from_secs(2)).unwrap();
        closer.join().unwrap();
        assert_owned(&ticket, 1);
        assert_eq!(install_production_config(DispatchConfig::default()), Ok(()));
        assert!(activate());
        assert_eq!(
            old.enqueue(value(2)).unwrap_err(),
            DispatchEnqueueError::StaleGeneration
        );
        let new_ticket = enqueue(value(3)).unwrap();
        let capability = CallCapability::new(backend.as_ref());
        MacroContext::new(&capability).drain_dispatcher();
        assert_owned(&new_ticket, 3);
    }

    #[test]
    fn queued_shutdown_signals_ticket_and_repeated_poll_is_stable() {
        let _serial = TEST_SERIAL.lock().unwrap();
        let (_guard, _) = start(DispatchConfig::default());
        let ticket = enqueue(value(1)).unwrap();
        shutdown();
        let expected = Some(Err(DispatchCompletionError::DispatcherShutdown));
        assert_eq!(ticket.try_result(), expected);
        assert_eq!(ticket.try_result(), expected);
        assert_eq!(
            enqueue(value(2)).unwrap_err(),
            DispatchEnqueueError::NoActiveGeneration
        );
    }

    #[test]
    fn xl_abort_failure_preserves_exact_code_and_no_auxiliary_owner() {
        let _serial = TEST_SERIAL.lock().unwrap();
        let (_guard, _) = start(DispatchConfig::default());
        let ticket = enqueue(DispatchOperation::PollCancellationPreservingBreak).unwrap();
        let backend = AbortBackend::linked();
        backend.code.store(
            excel_api_sys::xlretAbort | excel_api_sys::xlretUncalced,
            Ordering::Release,
        );
        let capability = CallCapability::new(&backend);
        MacroContext::new(&capability).drain_dispatcher();
        assert_eq!(
            ticket.try_result(),
            Some(Err(DispatchCompletionError::Operation(
                DispatchExecutionError::Excel(ExcelCallError::ExcelFailure {
                    function: "xlAbort",
                    code: crate::ExcelReturnCode(
                        excel_api_sys::xlretAbort | excel_api_sys::xlretUncalced
                    ),
                })
            )))
        );
        assert_eq!(
            *backend.calls.lock().unwrap(),
            vec![(excel_api_sys::xlAbort, 0)]
        );
    }
}
