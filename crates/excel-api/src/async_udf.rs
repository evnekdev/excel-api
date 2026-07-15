//! Owned, bounded asynchronous UDF scheduling and completion.

use core::{fmt, panic::AssertUnwindSafe};
use std::{
    collections::HashMap,
    sync::{
        Arc, Condvar, Mutex, OnceLock, Weak,
        atomic::{AtomicBool, AtomicU8, AtomicU64, Ordering},
        mpsc::{Receiver, SyncSender, TrySendError, sync_channel},
    },
    thread::JoinHandle,
};

use excel_api_sys::{LPXLOPER12, XLOPER12, XLOPER12BigData, XLOPER12BigDataHandle, XLOPER12Value};

use crate::{
    ExcelError, ExcelReturn, ExcelReturnCode, ExcelReturnValue, ThunkError,
    excel_call::ExcelCallBackend,
};

const SCHEDULED: u8 = 0;
const RUNNING: u8 = 1;
const CANCEL_REQUESTED: u8 = 2;
const COMPLETING: u8 = 3;
const COMPLETED: u8 = 4;
const CANCELED: u8 = 5;

/// One owned closure submitted to an [`AsyncExecutor`].
pub type AsyncJob = Box<dyn FnOnce() + Send + 'static>;

/// An executor rejection that returns ownership of the job it did not run.
pub struct AsyncExecuteError {
    error: AsyncSubmitError,
    job: AsyncJob,
}

impl AsyncExecuteError {
    pub fn new(error: AsyncSubmitError, job: AsyncJob) -> Self {
        Self { error, job }
    }

    pub fn error(&self) -> AsyncSubmitError {
        self.error
    }

    pub fn into_parts(self) -> (AsyncSubmitError, AsyncJob) {
        (self.error, self.job)
    }
}

impl fmt::Debug for AsyncExecuteError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("AsyncExecuteError")
            .field("error", &self.error)
            .field("job", &"<owned closure>")
            .finish()
    }
}

/// Runtime-neutral submission surface installed by an XLL application.
///
/// Each executor belongs to exactly one runtime generation. `execute` must
/// either accept the closure exactly once or return it without running it.
/// `shutdown` establishes an irreversible rejection boundary, waits until all
/// accepted closures can no longer execute XLL code, and is idempotent. It must
/// not detach work that can outlive XLL unload. The runtime never invokes
/// `shutdown` from one of the executor's jobs; third-party implementations that
/// permit that case must document it themselves. `execute` must not run the job
/// synchronously or wait for the job to finish before returning: controller
/// admission remains locked until acceptance commits.
pub trait AsyncExecutor: Send + Sync + 'static {
    fn execute(&self, job: AsyncJob) -> Result<(), AsyncExecuteError>;
    fn shutdown(&self);
}

enum ThreadPoolState {
    New,
    Running {
        sender: SyncSender<AsyncJob>,
        joins: Vec<JoinHandle<()>>,
    },
    ShuttingDown,
    Closed,
}

/// Small optional standard-library executor with a bounded queue.
///
/// Shutdown is permanent: a closed executor never creates workers again.
pub struct ThreadPoolExecutor {
    workers: usize,
    queue_bound: usize,
    state: Mutex<ThreadPoolState>,
    closed: Condvar,
    #[cfg(test)]
    fail_worker_at: Option<usize>,
}

impl ThreadPoolExecutor {
    pub fn new(workers: usize, queue_bound: usize) -> Option<Self> {
        (workers > 0 && queue_bound > 0).then_some(Self {
            workers,
            queue_bound,
            state: Mutex::new(ThreadPoolState::New),
            closed: Condvar::new(),
            #[cfg(test)]
            fail_worker_at: None,
        })
    }

    #[cfg(test)]
    fn failing_worker(workers: usize, queue_bound: usize, fail_worker_at: usize) -> Self {
        Self {
            workers,
            queue_bound,
            state: Mutex::new(ThreadPoolState::New),
            closed: Condvar::new(),
            fail_worker_at: Some(fail_worker_at),
        }
    }

    fn lock_state(&self) -> std::sync::MutexGuard<'_, ThreadPoolState> {
        self.state
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner)
    }

    fn start_locked(
        &self,
        state: &mut std::sync::MutexGuard<'_, ThreadPoolState>,
    ) -> Result<(), Vec<JoinHandle<()>>> {
        let (sender, receiver) = sync_channel(self.queue_bound);
        let receiver = Arc::new(Mutex::new(receiver));
        let mut joins = Vec::with_capacity(self.workers);
        for index in 0..self.workers {
            #[cfg(test)]
            if self.fail_worker_at == Some(index) {
                **state = ThreadPoolState::ShuttingDown;
                drop(sender);
                return Err(joins);
            }
            let receiver = receiver.clone();
            let name = format!("excel-api-async-{index}");
            match std::thread::Builder::new()
                .name(name)
                .spawn(move || worker_loop(&receiver))
            {
                Ok(join) => joins.push(join),
                Err(_) => {
                    **state = ThreadPoolState::ShuttingDown;
                    drop(sender);
                    return Err(joins);
                }
            }
        }
        **state = ThreadPoolState::Running { sender, joins };
        Ok(())
    }

    fn finish_failed_start(
        &self,
        state: std::sync::MutexGuard<'_, ThreadPoolState>,
        joins: Vec<JoinHandle<()>>,
    ) {
        drop(state);
        for join in joins {
            let _ = join.join();
        }
        let mut state = self.lock_state();
        *state = ThreadPoolState::Closed;
        self.closed.notify_all();
    }

    #[cfg(test)]
    fn lifecycle(&self) -> &'static str {
        match &*self.lock_state() {
            ThreadPoolState::New => "New",
            ThreadPoolState::Running { .. } => "Running",
            ThreadPoolState::ShuttingDown => "ShuttingDown",
            ThreadPoolState::Closed => "Closed",
        }
    }
}

fn worker_loop(receiver: &Mutex<Receiver<AsyncJob>>) {
    loop {
        let job = receiver
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner)
            .recv();
        match job {
            Ok(job) => job(),
            Err(_) => return,
        }
    }
}

impl AsyncExecutor for ThreadPoolExecutor {
    fn execute(&self, job: AsyncJob) -> Result<(), AsyncExecuteError> {
        let mut state = self.lock_state();
        if matches!(*state, ThreadPoolState::New) {
            if let Err(joins) = self.start_locked(&mut state) {
                self.finish_failed_start(state, joins);
                return Err(AsyncExecuteError::new(
                    AsyncSubmitError::WorkerStartFailed,
                    job,
                ));
            }
        }
        match &mut *state {
            ThreadPoolState::Running { sender, .. } => match sender.try_send(job) {
                Ok(()) => Ok(()),
                Err(TrySendError::Full(job)) => Err(AsyncExecuteError::new(
                    AsyncSubmitError::ExecutorQueueFull,
                    job,
                )),
                Err(TrySendError::Disconnected(job)) => Err(AsyncExecuteError::new(
                    AsyncSubmitError::ExecutorClosed,
                    job,
                )),
            },
            ThreadPoolState::ShuttingDown | ThreadPoolState::Closed => Err(AsyncExecuteError::new(
                AsyncSubmitError::ExecutorClosed,
                job,
            )),
            ThreadPoolState::New => unreachable!("startup resolved the New state"),
        }
    }

    fn shutdown(&self) {
        let joins = {
            let mut state = self.lock_state();
            loop {
                match std::mem::replace(&mut *state, ThreadPoolState::ShuttingDown) {
                    ThreadPoolState::New => {
                        *state = ThreadPoolState::Closed;
                        self.closed.notify_all();
                        return;
                    }
                    ThreadPoolState::Running { sender, joins } => {
                        drop(sender);
                        break joins;
                    }
                    ThreadPoolState::ShuttingDown => {
                        *state = ThreadPoolState::ShuttingDown;
                        state = self
                            .closed
                            .wait(state)
                            .unwrap_or_else(std::sync::PoisonError::into_inner);
                    }
                    ThreadPoolState::Closed => {
                        *state = ThreadPoolState::Closed;
                        return;
                    }
                }
            }
        };
        for join in joins {
            let _ = join.join();
        }
        let mut state = self.lock_state();
        *state = ThreadPoolState::Closed;
        self.closed.notify_all();
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum AsyncSubmitError {
    ExecutorUnavailable,
    ExecutorRejected,
    ExecutorQueueFull,
    ExecutorClosed,
    WorkerStartFailed,
    CapacityExhausted,
    RuntimeClosing,
    InvalidHandle,
}

impl fmt::Display for AsyncSubmitError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::ExecutorUnavailable => "no asynchronous executor is installed",
            Self::ExecutorRejected => "the asynchronous executor rejected the job",
            Self::ExecutorQueueFull => "the asynchronous executor queue is full",
            Self::ExecutorClosed => "the asynchronous executor generation is closed",
            Self::WorkerStartFailed => "an asynchronous executor worker could not be started",
            Self::CapacityExhausted => "the asynchronous in-flight limit was reached",
            Self::RuntimeClosing => "the asynchronous runtime is closing",
            Self::InvalidHandle => "Excel supplied an invalid asynchronous handle",
        })
    }
}

impl std::error::Error for AsyncSubmitError {}

#[derive(Debug)]
pub enum AsyncCompletionError {
    Canceled,
    AlreadyCompleted,
    RuntimeClosing,
    Return(ThunkError),
    Excel {
        code: ExcelReturnCode,
        accepted: Option<bool>,
    },
}

/// Cooperative cancellation signal available to asynchronous function bodies.
#[derive(Clone)]
pub struct AsyncCancellationToken {
    request: Weak<Request>,
}

impl AsyncCancellationToken {
    pub fn is_cancellation_requested(&self) -> bool {
        self.request.upgrade().is_none_or(|request| {
            matches!(
                request.state.load(Ordering::Acquire),
                CANCEL_REQUESTED | CANCELED
            ) || request
                .owner
                .upgrade()
                .is_none_or(|owner| !owner.is_active(request.epoch))
        })
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
struct AsyncHandle {
    value: usize,
    size: i32,
}

// The pointer-sized value is an opaque Excel token: it is copied and passed
// back but never dereferenced or freed by Rust.
// SAFETY: moving the copied integer bits does not grant pointer access.
unsafe impl Send for AsyncHandle {}
// SAFETY: shared access exposes no dereference or mutation operation.
unsafe impl Sync for AsyncHandle {}

impl AsyncHandle {
    unsafe fn copy_from(raw: LPXLOPER12) -> Result<Self, AsyncSubmitError> {
        // SAFETY: the generated thunk contract guarantees a readable callback root.
        let raw = unsafe { raw.as_ref() }.ok_or(AsyncSubmitError::InvalidHandle)?;
        if raw.xltype & excel_api_sys::XLTYPE_MASK != excel_api_sys::xltypeBigData {
            return Err(AsyncSubmitError::InvalidHandle);
        }
        // SAFETY: the validated tag selects the bigdata union member.
        let bigdata = unsafe { raw.val.bigdata };
        // SAFETY: Excel's async handle uses the hdata arm and remains opaque.
        let value = unsafe { bigdata.h.hdata } as usize;
        if value == 0 {
            return Err(AsyncSubmitError::InvalidHandle);
        }
        Ok(Self {
            value,
            size: bigdata.cbData,
        })
    }

    fn root(self) -> XLOPER12 {
        XLOPER12 {
            val: XLOPER12Value {
                bigdata: XLOPER12BigData {
                    h: XLOPER12BigDataHandle {
                        hdata: self.value as *mut core::ffi::c_void,
                    },
                    cbData: self.size,
                },
            },
            xltype: excel_api_sys::xltypeBigData,
        }
    }
}

struct Request {
    id: u64,
    epoch: u64,
    handle: AsyncHandle,
    state: AtomicU8,
    retired: AtomicBool,
    owner: Weak<AsyncController>,
}

impl Request {
    fn cancel(&self) {
        loop {
            let state = self.state.load(Ordering::Acquire);
            let next = match state {
                SCHEDULED => CANCELED,
                RUNNING => CANCEL_REQUESTED,
                _ => return,
            };
            if self
                .state
                .compare_exchange(state, next, Ordering::AcqRel, Ordering::Acquire)
                .is_ok()
            {
                return;
            }
        }
    }

    fn try_start(&self) -> bool {
        let Some(owner) = self.owner.upgrade() else {
            return false;
        };
        let state = owner.lock_state();
        let current = state.requests.get(&self.id).map(Arc::as_ptr);
        let registered = current == Some(self as *const Self);
        let may_start = state.active
            && owner.epoch == self.epoch
            && registered
            && !self.retired.load(Ordering::Acquire);
        may_start
            && self
                .state
                .compare_exchange(SCHEDULED, RUNNING, Ordering::AcqRel, Ordering::Acquire)
                .is_ok()
    }

    fn complete(&self, value: ExcelReturnValue) -> Result<(), AsyncCompletionError> {
        self.state
            .compare_exchange(RUNNING, COMPLETING, Ordering::AcqRel, Ordering::Acquire)
            .map_err(|state| match state {
                CANCEL_REQUESTED | CANCELED => AsyncCompletionError::Canceled,
                _ => AsyncCompletionError::AlreadyCompleted,
            })?;
        let Some(owner) = self.owner.upgrade() else {
            self.state.store(CANCELED, Ordering::Release);
            self.retired.store(true, Ordering::Release);
            return Err(AsyncCompletionError::RuntimeClosing);
        };
        if !owner.is_active(self.epoch) {
            self.state.store(CANCELED, Ordering::Release);
            self.retire();
            return Err(AsyncCompletionError::RuntimeClosing);
        }

        let result = value
            .plan()
            .map_err(ThunkError::ReturnPlanning)
            .and_then(|plan| plan.materialize().map_err(ThunkError::Materialization));
        let result = match result {
            Ok(result) => owner.return_to_excel(self.handle, result),
            Err(error) => Err(AsyncCompletionError::Return(error)),
        };
        self.state.store(COMPLETED, Ordering::Release);
        self.retire();
        result
    }

    fn retire(&self) {
        if self.retired.swap(true, Ordering::AcqRel) {
            return;
        }
        if let Some(owner) = self.owner.upgrade() {
            owner.retire(self.id, self as *const Self);
        }
    }
}

struct ControllerState {
    active: bool,
    requests: HashMap<u64, Arc<Request>>,
    in_flight: usize,
}

struct AsyncController {
    executor: Arc<dyn AsyncExecutor>,
    backend: Arc<dyn ExcelCallBackend>,
    maximum: usize,
    epoch: u64,
    next_id: AtomicU64,
    state: Mutex<ControllerState>,
    #[cfg(test)]
    before_submit: Mutex<Option<Arc<SubmissionPause>>>,
}

impl AsyncController {
    fn new(
        executor: Arc<dyn AsyncExecutor>,
        backend: Arc<dyn ExcelCallBackend>,
        maximum: usize,
        epoch: u64,
    ) -> Arc<Self> {
        Arc::new(Self {
            executor,
            backend,
            maximum: maximum.max(1),
            epoch,
            next_id: AtomicU64::new(1),
            state: Mutex::new(ControllerState {
                active: true,
                requests: HashMap::new(),
                in_flight: 0,
            }),
            #[cfg(test)]
            before_submit: Mutex::new(None),
        })
    }

    fn lock_state(&self) -> std::sync::MutexGuard<'_, ControllerState> {
        self.state
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner)
    }

    fn is_active(&self, epoch: u64) -> bool {
        let state = self.lock_state();
        state.active && self.epoch == epoch
    }

    fn schedule(
        self: &Arc<Self>,
        handle: AsyncHandle,
        task: Box<
            dyn FnOnce(AsyncCancellationToken) -> Result<ExcelReturnValue, ThunkError>
                + Send
                + 'static,
        >,
    ) -> Result<(), AsyncSubmitError> {
        let request = {
            let mut state = self.lock_state();
            if !state.active {
                return Err(AsyncSubmitError::RuntimeClosing);
            }
            if state.in_flight >= self.maximum {
                return Err(AsyncSubmitError::CapacityExhausted);
            }
            let id = self.next_id.fetch_add(1, Ordering::Relaxed);
            let request = Arc::new(Request {
                id,
                epoch: self.epoch,
                handle,
                state: AtomicU8::new(SCHEDULED),
                retired: AtomicBool::new(false),
                owner: Arc::downgrade(self),
            });
            state.in_flight += 1;
            state.requests.insert(id, request.clone());
            request
        };

        #[cfg(test)]
        if let Some(pause) = self
            .before_submit
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner)
            .clone()
        {
            pause.reached.wait();
            pause.release.wait();
        }

        let token = AsyncCancellationToken {
            request: Arc::downgrade(&request),
        };
        let job_request = request.clone();
        let job = Box::new(move || {
            if !job_request.try_start() {
                job_request.retire();
                return;
            }
            let outcome = std::panic::catch_unwind(AssertUnwindSafe(|| task(token)));
            let value = match outcome {
                Ok(Ok(value)) => value,
                Ok(Err(error)) => {
                    emit_async_failure();
                    ExcelReturnValue::Error(crate::thunk::error_for(&error))
                }
                Err(_) => {
                    emit_async_failure();
                    ExcelReturnValue::Error(ExcelError::Value)
                }
            };
            if let Err(AsyncCompletionError::Canceled) = job_request.complete(value) {
                job_request.state.store(CANCELED, Ordering::Release);
                job_request.retire();
            }
        }) as AsyncJob;
        // This second gate is the schedule/shutdown linearization point. If
        // shutdown won, the request is never offered to the executor. If this
        // call accepts, shutdown cannot deactivate the controller until the
        // executor has committed that acceptance. A worker attempting to start
        // waits on this same state mutex, so user code cannot run under it.
        let state = self.lock_state();
        if !state.active
            || state
                .requests
                .get(&request.id)
                .is_none_or(|registered| !Arc::ptr_eq(registered, &request))
        {
            drop(state);
            drop(job);
            request.cancel();
            request.retire();
            return Err(AsyncSubmitError::RuntimeClosing);
        }
        let submitted = self.executor.execute(job);
        drop(state);
        match submitted {
            Ok(()) => Ok(()),
            Err(error) => {
                let submit_error = error.error();
                let (_error, rejected_job) = error.into_parts();
                drop(rejected_job);
                request.cancel();
                request.retire();
                Err(submit_error)
            }
        }
    }

    fn retire(&self, id: u64, expected: *const Request) {
        let mut state = self.lock_state();
        let matches = state
            .requests
            .get(&id)
            .is_some_and(|request| Arc::as_ptr(request) == expected);
        if matches {
            state.requests.remove(&id);
            #[cfg(test)]
            debug_assert!(state.in_flight > 0, "async in-flight accounting underflow");
            if state.in_flight > 0 {
                state.in_flight -= 1;
            }
        }
    }

    fn cancel_all(&self) {
        let requests: Vec<_> = self.lock_state().requests.values().cloned().collect();
        for request in requests {
            request.cancel();
            if request.state.load(Ordering::Acquire) == CANCELED {
                request.retire();
            }
        }
    }

    fn calculation_ended(&self) {
        self.cancel_all();
    }

    fn shutdown(&self) {
        let requests = {
            let mut state = self.lock_state();
            if !state.active {
                None
            } else {
                state.active = false;
                Some(state.requests.values().cloned().collect::<Vec<_>>())
            }
        };
        let Some(requests) = requests else {
            self.executor.shutdown();
            return;
        };
        for request in &requests {
            request.cancel();
        }
        // No controller/global lock is held while the executor joins workers.
        self.executor.shutdown();
        for request in requests {
            request.state.store(CANCELED, Ordering::Release);
            request.retire();
        }
    }

    fn return_to_excel(
        &self,
        handle: AsyncHandle,
        mut value: ExcelReturn,
    ) -> Result<(), AsyncCompletionError> {
        if !self.is_active(self.epoch) || !self.backend.is_linked() {
            return Err(AsyncCompletionError::RuntimeClosing);
        }
        let mut handle = handle.root();
        let mut response = XLOPER12 {
            val: XLOPER12Value { xbool: 0 },
            xltype: excel_api_sys::xltypeNil,
        };
        let mut arguments = [
            &mut handle as LPXLOPER12,
            value.as_xloper_mut_for_callback() as LPXLOPER12,
        ];
        // SAFETY: roots remain stable for this synchronous xlAsyncReturn call.
        let raw = unsafe {
            self.backend.excel12v_raw(
                excel_api_sys::xlAsyncReturn,
                &mut response,
                2,
                arguments.as_mut_ptr(),
            )
        };
        let accepted = if response.xltype & excel_api_sys::XLTYPE_MASK == excel_api_sys::xltypeBool
        {
            // SAFETY: guarded by the Boolean tag.
            Some(unsafe { response.val.xbool != 0 })
        } else {
            None
        };
        if raw == excel_api_sys::xlretSuccess && accepted == Some(true) {
            Ok(())
        } else {
            Err(AsyncCompletionError::Excel {
                code: ExcelReturnCode(raw),
                accepted,
            })
        }
    }

    #[cfg(test)]
    fn accounting(&self) -> (usize, usize) {
        let state = self.lock_state();
        (state.in_flight, state.requests.len())
    }
}

fn emit_async_failure() {
    crate::diagnostics::emit(crate::DiagnosticEvent::new(
        crate::DiagnosticCode::ThunkFailure,
        crate::DiagnosticSeverity::Error,
        0,
    ));
}

struct PendingExecutor {
    executor: Arc<dyn AsyncExecutor>,
    maximum_in_flight: usize,
}

#[derive(Default)]
struct Generations {
    pending: Option<PendingExecutor>,
    active: Option<Arc<AsyncController>>,
    next_epoch: u64,
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

pub(crate) fn install_production_executor(
    executor: Arc<dyn AsyncExecutor>,
    maximum_in_flight: usize,
) -> Result<(), Arc<dyn AsyncExecutor>> {
    let mut generations = lock_generations();
    if generations.pending.is_some() || generations.active.is_some() {
        return Err(executor);
    }
    generations.pending = Some(PendingExecutor {
        executor,
        maximum_in_flight: maximum_in_flight.max(1),
    });
    Ok(())
}

pub(crate) fn activate(backend: Arc<dyn ExcelCallBackend>) -> Result<(), AsyncSubmitError> {
    let mut generations = lock_generations();
    if generations.active.is_some() {
        return Ok(());
    }
    let pending = generations
        .pending
        .take()
        .ok_or(AsyncSubmitError::ExecutorUnavailable)?;
    generations.next_epoch = generations.next_epoch.wrapping_add(1).max(1);
    generations.active = Some(AsyncController::new(
        pending.executor,
        backend,
        pending.maximum_in_flight,
        generations.next_epoch,
    ));
    Ok(())
}

pub(crate) fn shutdown() {
    let controller = lock_generations().active.take();
    if let Some(controller) = controller {
        controller.shutdown();
    }
}

fn active_controller() -> Option<Arc<AsyncController>> {
    lock_generations().active.clone()
}

#[cfg(test)]
pub(crate) fn reset_generations_for_test() {
    let (active, pending) = {
        let mut generations = lock_generations();
        let active = generations.active.take();
        let pending = generations.pending.take();
        (active, pending)
    };
    if let Some(active) = active {
        active.shutdown();
    }
    if let Some(pending) = pending {
        pending.executor.shutdown();
    }
}

pub fn calculation_canceled() {
    if let Some(controller) = active_controller() {
        controller.cancel_all();
    }
}

pub fn calculation_ended() {
    if let Some(controller) = active_controller() {
        controller.calculation_ended();
    }
}

/// Excel calculation-cancellation event procedure registered by name.
#[unsafe(export_name = "excel_api_calculation_canceled")]
pub extern "system" fn calculation_canceled_event() {
    let _ = std::panic::catch_unwind(calculation_canceled);
}

/// Excel calculation-ended event procedure registered by name.
#[unsafe(export_name = "excel_api_calculation_ended")]
pub extern "system" fn calculation_ended_event() {
    let _ = std::panic::catch_unwind(calculation_ended);
}

#[doc(hidden)]
pub unsafe fn schedule(
    raw_handle: LPXLOPER12,
    task: Box<
        dyn FnOnce(AsyncCancellationToken) -> Result<ExcelReturnValue, ThunkError> + Send + 'static,
    >,
) -> Result<(), AsyncSubmitError> {
    let controller = active_controller().ok_or(AsyncSubmitError::ExecutorUnavailable)?;
    // SAFETY: forwarded from the generated async thunk contract.
    let handle = unsafe { AsyncHandle::copy_from(raw_handle) }?;
    controller.schedule(handle, task)
}

#[cfg(test)]
struct SubmissionPause {
    reached: std::sync::Barrier,
    release: std::sync::Barrier,
}

#[cfg(test)]
mod tests {
    use super::*;
    use core::ffi::c_int;
    use std::{
        collections::VecDeque,
        sync::{Barrier, mpsc},
        time::Duration,
    };

    type RecordedCall = (i32, usize, usize, u32, f64);

    struct QueueState {
        jobs: VecDeque<AsyncJob>,
        closed: bool,
    }

    #[derive(Default)]
    struct QueueExecutor {
        state: Mutex<Option<QueueState>>,
        shutdowns: AtomicU64,
    }

    impl QueueExecutor {
        fn initialized() -> Self {
            Self {
                state: Mutex::new(Some(QueueState {
                    jobs: VecDeque::new(),
                    closed: false,
                })),
                shutdowns: AtomicU64::new(0),
            }
        }

        fn run_one(&self) {
            let job = self
                .state
                .lock()
                .unwrap()
                .as_mut()
                .unwrap()
                .jobs
                .pop_front()
                .expect("queued job");
            job();
        }
    }

    impl AsyncExecutor for QueueExecutor {
        fn execute(&self, job: AsyncJob) -> Result<(), AsyncExecuteError> {
            let mut state = self.state.lock().unwrap();
            let state = state.as_mut().unwrap();
            if state.closed {
                Err(AsyncExecuteError::new(
                    AsyncSubmitError::ExecutorClosed,
                    job,
                ))
            } else {
                state.jobs.push_back(job);
                Ok(())
            }
        }

        fn shutdown(&self) {
            self.shutdowns.fetch_add(1, Ordering::AcqRel);
            let jobs = {
                let mut state = self.state.lock().unwrap();
                let state = state.as_mut().unwrap();
                state.closed = true;
                std::mem::take(&mut state.jobs)
            };
            for job in jobs {
                job();
            }
        }
    }

    #[derive(Default)]
    struct MockBackend {
        linked: AtomicBool,
        code: std::sync::atomic::AtomicI32,
        accepted: AtomicBool,
        calls: Mutex<Vec<RecordedCall>>,
    }

    impl MockBackend {
        fn linked() -> Self {
            Self {
                linked: AtomicBool::new(true),
                accepted: AtomicBool::new(true),
                ..Self::default()
            }
        }
    }

    impl ExcelCallBackend for MockBackend {
        fn link(&self) -> Result<(), crate::ExcelCallError> {
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
            arguments: *mut LPXLOPER12,
        ) -> i32 {
            // SAFETY: the controller supplies the exact two initialized roots.
            let handle = unsafe { (**arguments).val.bigdata };
            // SAFETY: the second root is the numeric fixture result.
            let value = unsafe { (**arguments.add(1)).val.num };
            // SAFETY: the handle tag selects the opaque hdata arm.
            let handle_value = unsafe { handle.h.hdata } as usize;
            // SAFETY: the second argument root is initialized and live.
            let value_type = unsafe { (**arguments.add(1)).xltype };
            self.calls.lock().unwrap().push((
                function,
                count as usize,
                handle_value,
                value_type,
                value,
            ));
            // SAFETY: the test backend receives a required live result root.
            unsafe {
                (*result).val = XLOPER12Value {
                    xbool: i32::from(self.accepted.load(Ordering::Acquire)),
                };
                (*result).xltype = excel_api_sys::xltypeBool;
            }
            self.code.load(Ordering::Acquire)
        }
    }

    fn handle(value: usize) -> AsyncHandle {
        AsyncHandle { value, size: 8 }
    }

    fn number(value: f64) -> ExcelReturnValue {
        ExcelReturnValue::Number(value)
    }

    fn controller(
        executor: Arc<dyn AsyncExecutor>,
        backend: Arc<MockBackend>,
        maximum: usize,
        epoch: u64,
    ) -> Arc<AsyncController> {
        AsyncController::new(executor, backend, maximum, epoch)
    }

    #[test]
    fn thread_pool_is_lazy_bounded_and_permanently_closed() {
        let pool = Arc::new(ThreadPoolExecutor::new(1, 1).unwrap());
        assert_eq!(pool.lifecycle(), "New");
        let (started_tx, started_rx) = mpsc::channel();
        let release = Arc::new(Barrier::new(2));
        let worker_release = release.clone();
        pool.execute(Box::new(move || {
            started_tx.send(()).unwrap();
            worker_release.wait();
        }))
        .unwrap();
        started_rx.recv_timeout(Duration::from_secs(2)).unwrap();
        pool.execute(Box::new(|| {})).unwrap();
        let rejected = pool.execute(Box::new(|| {})).unwrap_err();
        assert_eq!(rejected.error(), AsyncSubmitError::ExecutorQueueFull);
        release.wait();
        pool.shutdown();
        assert_eq!(pool.lifecycle(), "Closed");
        pool.shutdown();
        assert_eq!(
            pool.execute(Box::new(|| {})).unwrap_err().error(),
            AsyncSubmitError::ExecutorClosed
        );
        assert_eq!(pool.lifecycle(), "Closed");
    }

    #[test]
    fn thread_pool_shutdown_with_empty_queue_is_permanent() {
        let pool = ThreadPoolExecutor::new(1, 1).unwrap();
        pool.shutdown();
        assert_eq!(pool.lifecycle(), "Closed");
        assert_eq!(
            pool.execute(Box::new(|| {})).unwrap_err().error(),
            AsyncSubmitError::ExecutorClosed
        );
    }

    #[test]
    fn thread_pool_runs_multiple_workers_and_joins_queued_work() {
        let pool = Arc::new(ThreadPoolExecutor::new(2, 4).unwrap());
        let barrier = Arc::new(Barrier::new(3));
        let counter = Arc::new(std::sync::atomic::AtomicUsize::new(0));
        for _ in 0..2 {
            let barrier = barrier.clone();
            let counter = counter.clone();
            pool.execute(Box::new(move || {
                counter.fetch_add(1, Ordering::AcqRel);
                barrier.wait();
            }))
            .unwrap();
        }
        barrier.wait();
        pool.execute(Box::new({
            let counter = counter.clone();
            move || {
                counter.fetch_add(1, Ordering::AcqRel);
            }
        }))
        .unwrap();
        pool.shutdown();
        assert_eq!(counter.load(Ordering::Acquire), 3);
    }

    #[test]
    fn thread_pool_partial_start_failure_closes_and_joins() {
        let pool = ThreadPoolExecutor::failing_worker(2, 2, 1);
        assert_eq!(
            pool.execute(Box::new(|| {})).unwrap_err().error(),
            AsyncSubmitError::WorkerStartFailed
        );
        assert_eq!(pool.lifecycle(), "Closed");
        assert_eq!(
            pool.execute(Box::new(|| {})).unwrap_err().error(),
            AsyncSubmitError::ExecutorClosed
        );
    }

    #[test]
    fn execute_racing_shutdown_is_linearizable() {
        for _ in 0..32 {
            let pool = Arc::new(ThreadPoolExecutor::new(1, 2).unwrap());
            let gate = Arc::new(Barrier::new(2));
            let ran = Arc::new(AtomicBool::new(false));
            let submitter = {
                let pool = pool.clone();
                let gate = gate.clone();
                let ran = ran.clone();
                std::thread::spawn(move || {
                    gate.wait();
                    pool.execute(Box::new(move || ran.store(true, Ordering::Release)))
                })
            };
            gate.wait();
            pool.shutdown();
            let result = submitter.join().unwrap();
            assert!(result.is_err() || ran.load(Ordering::Acquire));
            assert_eq!(pool.lifecycle(), "Closed");
        }
    }

    #[test]
    fn canceled_queued_work_skips_user_body_and_retires_once() {
        let executor = Arc::new(QueueExecutor::initialized());
        let backend = Arc::new(MockBackend::linked());
        let controller = controller(executor.clone(), backend.clone(), 2, 1);
        let called = Arc::new(AtomicBool::new(false));
        let body_called = called.clone();
        controller
            .schedule(
                handle(1),
                Box::new(move |_| {
                    body_called.store(true, Ordering::Release);
                    Ok(number(1.0))
                }),
            )
            .unwrap();
        controller.cancel_all();
        executor.run_one();
        assert!(!called.load(Ordering::Acquire));
        assert!(backend.calls.lock().unwrap().is_empty());
        assert_eq!(controller.accounting(), (0, 0));
    }

    #[test]
    fn running_cancellation_sets_token_and_suppresses_ignored_result() {
        let executor = Arc::new(ThreadPoolExecutor::new(1, 2).unwrap());
        let backend = Arc::new(MockBackend::linked());
        let controller = controller(executor, backend.clone(), 2, 1);
        let entered = Arc::new(Barrier::new(2));
        let leave = Arc::new(Barrier::new(2));
        let observed = Arc::new(AtomicBool::new(false));
        controller
            .schedule(
                handle(1),
                Box::new({
                    let entered = entered.clone();
                    let leave = leave.clone();
                    let observed = observed.clone();
                    move |token| {
                        entered.wait();
                        leave.wait();
                        observed.store(token.is_cancellation_requested(), Ordering::Release);
                        Ok(number(9.0))
                    }
                }),
            )
            .unwrap();
        entered.wait();
        controller.cancel_all();
        leave.wait();
        controller.shutdown();
        assert!(observed.load(Ordering::Acquire));
        assert!(backend.calls.lock().unwrap().is_empty());
        assert_eq!(controller.accounting(), (0, 0));
    }

    #[test]
    fn shutdown_between_reservation_and_submission_rejects_without_running() {
        let executor = Arc::new(ThreadPoolExecutor::new(1, 2).unwrap());
        let backend = Arc::new(MockBackend::linked());
        let controller = controller(executor, backend, 1, 1);
        let pause = Arc::new(SubmissionPause {
            reached: Barrier::new(2),
            release: Barrier::new(2),
        });
        *controller.before_submit.lock().unwrap() = Some(pause.clone());
        let called = Arc::new(AtomicBool::new(false));
        let schedule = {
            let controller = controller.clone();
            let called = called.clone();
            std::thread::spawn(move || {
                controller.schedule(
                    handle(1),
                    Box::new(move |_| {
                        called.store(true, Ordering::Release);
                        Ok(number(1.0))
                    }),
                )
            })
        };
        pause.reached.wait();
        controller.shutdown();
        pause.release.wait();
        assert_eq!(
            schedule.join().unwrap(),
            Err(AsyncSubmitError::RuntimeClosing)
        );
        assert!(!called.load(Ordering::Acquire));
        assert_eq!(controller.accounting(), (0, 0));
    }

    #[test]
    fn schedule_before_reserve_and_after_accept_obey_shutdown_boundary() {
        let executor = Arc::new(ThreadPoolExecutor::new(1, 4).unwrap());
        let backend = Arc::new(MockBackend::linked());
        let controller = controller(executor, backend.clone(), 3, 1);
        let entered = Arc::new(Barrier::new(2));
        let release = Arc::new(Barrier::new(2));
        controller
            .schedule(
                handle(1),
                Box::new({
                    let entered = entered.clone();
                    let release = release.clone();
                    move |token| {
                        entered.wait();
                        release.wait();
                        assert!(token.is_cancellation_requested());
                        Ok(number(1.0))
                    }
                }),
            )
            .unwrap();
        entered.wait();
        let queued_called = Arc::new(AtomicBool::new(false));
        controller
            .schedule(
                handle(2),
                Box::new({
                    let queued_called = queued_called.clone();
                    move |_| {
                        queued_called.store(true, Ordering::Release);
                        Ok(number(2.0))
                    }
                }),
            )
            .unwrap();
        let shutdown = {
            let controller = controller.clone();
            std::thread::spawn(move || controller.shutdown())
        };
        while controller.is_active(1) {
            std::thread::yield_now();
        }
        assert_eq!(
            controller.schedule(handle(3), Box::new(|_| Ok(number(3.0)))),
            Err(AsyncSubmitError::RuntimeClosing)
        );
        release.wait();
        shutdown.join().unwrap();
        assert!(!queued_called.load(Ordering::Acquire));
        assert!(backend.calls.lock().unwrap().is_empty());
        assert_eq!(controller.accounting(), (0, 0));
    }

    #[test]
    fn successful_completion_is_at_most_once_and_preserves_handle() {
        let executor = Arc::new(QueueExecutor::initialized());
        let backend = Arc::new(MockBackend::linked());
        let controller = controller(executor.clone(), backend.clone(), 2, 7);
        controller
            .schedule(handle(0x1234), Box::new(|_| Ok(number(4.5))))
            .unwrap();
        executor.run_one();
        assert_eq!(
            backend.calls.lock().unwrap().as_slice(),
            &[(
                excel_api_sys::xlAsyncReturn,
                2,
                0x1234,
                excel_api_sys::xltypeNum,
                4.5,
            )]
        );
        assert_eq!(controller.accounting(), (0, 0));
    }

    #[test]
    fn completion_racing_cancellation_is_linearizable_and_retires_once() {
        for _ in 0..32 {
            let executor = Arc::new(QueueExecutor::initialized());
            let backend = Arc::new(MockBackend::linked());
            let controller = controller(executor, backend.clone(), 1, 1);
            let request = {
                let request = Arc::new(Request {
                    id: 1,
                    epoch: 1,
                    handle: handle(1),
                    state: AtomicU8::new(RUNNING),
                    retired: AtomicBool::new(false),
                    owner: Arc::downgrade(&controller),
                });
                let mut state = controller.lock_state();
                state.in_flight = 1;
                state.requests.insert(1, request.clone());
                request
            };
            let barrier = Arc::new(Barrier::new(3));
            let completing = {
                let request = request.clone();
                let barrier = barrier.clone();
                std::thread::spawn(move || {
                    barrier.wait();
                    let result = request.complete(number(1.0));
                    if matches!(result, Err(AsyncCompletionError::Canceled)) {
                        request.state.store(CANCELED, Ordering::Release);
                        request.retire();
                    }
                    result
                })
            };
            let canceling = {
                let request = request.clone();
                let barrier = barrier.clone();
                std::thread::spawn(move || {
                    barrier.wait();
                    request.cancel();
                })
            };
            barrier.wait();
            let completion = completing.join().unwrap();
            canceling.join().unwrap();
            assert!(
                completion.is_ok() || matches!(completion, Err(AsyncCompletionError::Canceled))
            );
            assert!(backend.calls.lock().unwrap().len() <= 1);
            assert_eq!(controller.accounting(), (0, 0));
        }
    }

    #[test]
    fn excel_rejection_and_user_panic_each_retire_exactly_once() {
        let executor = Arc::new(QueueExecutor::initialized());
        let backend = Arc::new(MockBackend::linked());
        backend.accepted.store(false, Ordering::Release);
        let controller = controller(executor.clone(), backend.clone(), 2, 1);
        controller
            .schedule(handle(1), Box::new(|_| Ok(number(1.0))))
            .unwrap();
        controller
            .schedule(
                handle(2),
                Box::new(|_| -> Result<ExcelReturnValue, ThunkError> { panic!("fixture") }),
            )
            .unwrap();
        executor.run_one();
        executor.run_one();
        assert_eq!(backend.calls.lock().unwrap().len(), 2);
        assert_eq!(controller.accounting(), (0, 0));
    }

    #[test]
    fn excel_return_code_and_boolean_rejection_are_preserved() {
        let executor = Arc::new(QueueExecutor::initialized());
        let backend = Arc::new(MockBackend::linked());
        backend.accepted.store(false, Ordering::Release);
        backend.code.store(
            excel_api_sys::xlretInvAsynchronousContext,
            Ordering::Release,
        );
        let controller = controller(executor, backend, 1, 1);
        let request = Arc::new(Request {
            id: 1,
            epoch: 1,
            handle: handle(1),
            state: AtomicU8::new(RUNNING),
            retired: AtomicBool::new(false),
            owner: Arc::downgrade(&controller),
        });
        {
            let mut state = controller.lock_state();
            state.in_flight = 1;
            state.requests.insert(1, request.clone());
        }
        assert!(matches!(
            request.complete(number(1.0)),
            Err(AsyncCompletionError::Excel {
                code: ExcelReturnCode(excel_api_sys::xlretInvAsynchronousContext),
                accepted: Some(false),
            })
        ));
        assert_eq!(controller.accounting(), (0, 0));
    }

    #[test]
    fn invalid_async_handles_are_rejected_before_submission() {
        assert_eq!(
            // SAFETY: null intentionally exercises defensive validation.
            unsafe { AsyncHandle::copy_from(core::ptr::null_mut()) },
            Err(AsyncSubmitError::InvalidHandle)
        );
        let mut wrong = XLOPER12 {
            val: XLOPER12Value { num: 1.0 },
            xltype: excel_api_sys::xltypeNum,
        };
        assert_eq!(
            // SAFETY: `wrong` is a readable live root with a deliberately wrong tag.
            unsafe { AsyncHandle::copy_from(&mut wrong) },
            Err(AsyncSubmitError::InvalidHandle)
        );
    }

    #[test]
    fn stale_generation_cannot_complete_after_reopen_or_unlink() {
        let old_executor = Arc::new(QueueExecutor::initialized());
        let backend = Arc::new(MockBackend::linked());
        let old = controller(old_executor.clone(), backend.clone(), 1, 1);
        old.schedule(handle(1), Box::new(|_| Ok(number(1.0))))
            .unwrap();
        old.shutdown();
        let new_executor = Arc::new(QueueExecutor::initialized());
        let new = controller(new_executor.clone(), backend.clone(), 1, 2);
        new.schedule(handle(2), Box::new(|_| Ok(number(2.0))))
            .unwrap();
        backend.unlink();
        new_executor.run_one();
        assert!(backend.calls.lock().unwrap().is_empty());
        new.shutdown();
    }

    #[test]
    fn calculation_canceled_then_ended_is_idempotent() {
        let executor = Arc::new(QueueExecutor::initialized());
        let backend = Arc::new(MockBackend::linked());
        let controller = controller(executor.clone(), backend, 1, 1);
        controller
            .schedule(handle(1), Box::new(|_| Ok(number(1.0))))
            .unwrap();
        controller.cancel_all();
        controller.calculation_ended();
        executor.run_one();
        assert_eq!(controller.accounting(), (0, 0));
    }

    #[test]
    fn executor_rejection_preserves_specific_error_and_accounting() {
        let executor = Arc::new(QueueExecutor::initialized());
        executor.shutdown();
        let backend = Arc::new(MockBackend::linked());
        let controller = controller(executor, backend, 1, 1);
        assert_eq!(
            controller.schedule(handle(1), Box::new(|_| Ok(number(1.0)))),
            Err(AsyncSubmitError::ExecutorClosed)
        );
        assert_eq!(controller.accounting(), (0, 0));
    }
}
