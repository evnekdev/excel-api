//! Owned, bounded asynchronous UDF scheduling and completion.

use core::{fmt, panic::AssertUnwindSafe};
use std::{
    collections::HashMap,
    sync::{
        Arc, Mutex, OnceLock, Weak,
        atomic::{AtomicBool, AtomicU8, AtomicU64, AtomicUsize, Ordering},
        mpsc::{Receiver, SyncSender, TrySendError, sync_channel},
    },
    thread::JoinHandle,
};

use excel_api_sys::{LPXLOPER12, XLOPER12, XLOPER12BigData, XLOPER12BigDataHandle, XLOPER12Value};

use crate::{
    ExcelError, ExcelReturn, ExcelReturnCode, ExcelReturnValue, ThunkError,
    excel_call::{ExcelCallBackend, SdkExcel12vBackend},
};

const SCHEDULED: u8 = 0;
const COMPLETING: u8 = 1;
const COMPLETED: u8 = 2;
const CANCELED: u8 = 3;

/// Runtime-neutral submission surface installed by an XLL application.
///
/// `shutdown` must reject new work and wait until previously accepted jobs can
/// no longer execute XLL code. This is required before the DLL can be unloaded.
pub trait AsyncExecutor: Send + Sync + 'static {
    /// Accepts a job exactly once. Returning `Err` means the executor did not
    /// run and will never run the supplied job.
    fn execute(&self, job: Box<dyn FnOnce() + Send + 'static>) -> Result<(), AsyncSubmitError>;
    fn shutdown(&self);
}

type ExecutorJob = Box<dyn FnOnce() + Send + 'static>;

/// Small optional standard-library executor for XLLs that do not bring an
/// async runtime. It uses a bounded queue and joins every worker on shutdown.
pub struct ThreadPoolExecutor {
    workers: usize,
    queue_bound: usize,
    sender: Mutex<Option<SyncSender<ExecutorJob>>>,
    joins: Mutex<Vec<JoinHandle<()>>>,
}

impl ThreadPoolExecutor {
    pub fn new(workers: usize, queue_bound: usize) -> Option<Self> {
        (workers > 0 && queue_bound > 0).then_some(Self {
            workers,
            queue_bound,
            sender: Mutex::new(None),
            joins: Mutex::new(Vec::new()),
        })
    }

    fn ensure_started(&self) -> Result<SyncSender<ExecutorJob>, AsyncSubmitError> {
        let mut sender = self
            .sender
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner);
        if let Some(sender) = sender.as_ref() {
            return Ok(sender.clone());
        }
        let (new_sender, receiver) = sync_channel(self.queue_bound);
        let receiver = Arc::new(Mutex::new(receiver));
        let mut joins = self
            .joins
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner);
        debug_assert!(joins.is_empty());
        for index in 0..self.workers {
            let receiver = receiver.clone();
            let name = format!("excel-api-async-{index}");
            if let Ok(join) = std::thread::Builder::new()
                .name(name)
                .spawn(move || worker_loop(&receiver))
            {
                joins.push(join);
            }
        }
        if joins.is_empty() {
            return Err(AsyncSubmitError::ExecutorRejected);
        }
        *sender = Some(new_sender.clone());
        Ok(new_sender)
    }
}

fn worker_loop(receiver: &Mutex<Receiver<ExecutorJob>>) {
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
    fn execute(&self, job: ExecutorJob) -> Result<(), AsyncSubmitError> {
        match self.ensure_started()?.try_send(job) {
            Ok(()) => Ok(()),
            Err(TrySendError::Full(_)) | Err(TrySendError::Disconnected(_)) => {
                Err(AsyncSubmitError::ExecutorRejected)
            }
        }
    }

    fn shutdown(&self) {
        self.sender
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner)
            .take();
        let joins = std::mem::take(
            &mut *self
                .joins
                .lock()
                .unwrap_or_else(std::sync::PoisonError::into_inner),
        );
        for join in joins {
            let _ = join.join();
        }
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum AsyncSubmitError {
    ExecutorUnavailable,
    ExecutorRejected,
    CapacityExhausted,
    RuntimeClosing,
    InvalidHandle,
}

impl fmt::Display for AsyncSubmitError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::ExecutorUnavailable => "no asynchronous executor is installed",
            Self::ExecutorRejected => "the asynchronous executor rejected the job",
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
        self.request
            .upgrade()
            .is_none_or(|request| request.state.load(Ordering::Acquire) == CANCELED)
    }
}

#[derive(Clone, Copy)]
struct AsyncHandle {
    value: usize,
    size: i32,
}

// The pointer-sized value is an opaque Excel token: it is copied and passed
// back but never dereferenced or freed by Rust.
// SAFETY: crossing threads does not grant pointer access; only the integer
// token and size are copied into a new callback root.
unsafe impl Send for AsyncHandle {}
// SAFETY: shared access exposes no operation that dereferences or mutates the
// opaque Excel token.
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
        // SAFETY: Excel's async handle uses the hdata arm and is treated as opaque.
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
    owner: Weak<AsyncController>,
}

impl Request {
    fn cancel(&self) {
        let _ =
            self.state
                .compare_exchange(SCHEDULED, CANCELED, Ordering::AcqRel, Ordering::Acquire);
    }

    fn complete(&self, value: ExcelReturnValue) -> Result<(), AsyncCompletionError> {
        self.state
            .compare_exchange(SCHEDULED, COMPLETING, Ordering::AcqRel, Ordering::Acquire)
            .map_err(|state| match state {
                CANCELED => AsyncCompletionError::Canceled,
                _ => AsyncCompletionError::AlreadyCompleted,
            })?;
        let Some(owner) = self.owner.upgrade() else {
            self.state.store(CANCELED, Ordering::Release);
            return Err(AsyncCompletionError::RuntimeClosing);
        };
        if !owner.active.load(Ordering::Acquire)
            || owner.epoch.load(Ordering::Acquire) != self.epoch
        {
            self.state.store(CANCELED, Ordering::Release);
            owner.finish(self.id);
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
        owner.finish(self.id);
        result
    }
}

struct AsyncController {
    executor: Arc<dyn AsyncExecutor>,
    backend: Arc<dyn ExcelCallBackend>,
    maximum: usize,
    active: AtomicBool,
    epoch: AtomicU64,
    next_id: AtomicU64,
    requests: Mutex<HashMap<u64, Arc<Request>>>,
    in_flight: AtomicUsize,
}

impl AsyncController {
    fn new(
        executor: Arc<dyn AsyncExecutor>,
        backend: Arc<dyn ExcelCallBackend>,
        maximum: usize,
    ) -> Arc<Self> {
        Arc::new(Self {
            executor,
            backend,
            maximum: maximum.max(1),
            active: AtomicBool::new(false),
            epoch: AtomicU64::new(0),
            next_id: AtomicU64::new(1),
            requests: Mutex::new(HashMap::new()),
            in_flight: AtomicUsize::new(0),
        })
    }

    fn activate(&self) {
        self.epoch.fetch_add(1, Ordering::AcqRel);
        self.active.store(true, Ordering::Release);
    }

    fn reserve(&self) -> Result<(), AsyncSubmitError> {
        if !self.active.load(Ordering::Acquire) {
            return Err(AsyncSubmitError::RuntimeClosing);
        }
        self.in_flight
            .fetch_update(Ordering::AcqRel, Ordering::Acquire, |current| {
                (current < self.maximum).then_some(current + 1)
            })
            .map(|_| ())
            .map_err(|_| AsyncSubmitError::CapacityExhausted)
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
        self.reserve()?;
        let id = self.next_id.fetch_add(1, Ordering::Relaxed);
        let request = Arc::new(Request {
            id,
            epoch: self.epoch.load(Ordering::Acquire),
            handle,
            state: AtomicU8::new(SCHEDULED),
            owner: Arc::downgrade(self),
        });
        self.requests
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner)
            .insert(id, request.clone());
        let token = AsyncCancellationToken {
            request: Arc::downgrade(&request),
        };
        let job = Box::new(move || {
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
            let _ = request.complete(value);
        });
        if self.executor.execute(job).is_err() {
            if let Some(request) = self
                .requests
                .lock()
                .unwrap_or_else(std::sync::PoisonError::into_inner)
                .remove(&id)
            {
                request.cancel();
                self.in_flight.fetch_sub(1, Ordering::AcqRel);
            }
            return Err(AsyncSubmitError::ExecutorRejected);
        }
        Ok(())
    }

    fn finish(&self, id: u64) {
        if self
            .requests
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner)
            .remove(&id)
            .is_some()
        {
            self.in_flight.fetch_sub(1, Ordering::AcqRel);
        }
    }

    fn cancel_all(&self) {
        for request in self
            .requests
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner)
            .values()
        {
            request.cancel();
        }
    }

    fn calculation_ended(&self) {
        let removed = {
            let mut requests = self
                .requests
                .lock()
                .unwrap_or_else(std::sync::PoisonError::into_inner);
            let before = requests.len();
            requests.retain(|_, request| {
                matches!(
                    request.state.load(Ordering::Acquire),
                    SCHEDULED | COMPLETING
                )
            });
            before - requests.len()
        };
        self.in_flight.fetch_sub(removed, Ordering::AcqRel);
    }

    fn shutdown(&self) {
        self.active.store(false, Ordering::Release);
        self.epoch.fetch_add(1, Ordering::AcqRel);
        self.cancel_all();
        self.executor.shutdown();
        self.calculation_ended();
    }

    fn return_to_excel(
        &self,
        handle: AsyncHandle,
        mut value: ExcelReturn,
    ) -> Result<(), AsyncCompletionError> {
        if !self.active.load(Ordering::Acquire) || !self.backend.is_linked() {
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
        // SAFETY: both roots and their backing allocations remain stable for
        // the synchronous xlAsyncReturn callback. No other Excel callback is
        // exposed to this worker operation.
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
}

fn emit_async_failure() {
    crate::diagnostics::emit(crate::DiagnosticEvent::new(
        crate::DiagnosticCode::ThunkFailure,
        crate::DiagnosticSeverity::Error,
        0,
    ));
}

// Use one shared cell. Kept behind a function so lifecycle and thunk code do
// not expose the controller type.
fn controller_cell() -> &'static OnceLock<Arc<AsyncController>> {
    static CONTROLLER: OnceLock<Arc<AsyncController>> = OnceLock::new();
    &CONTROLLER
}

pub(crate) fn install_production_executor(
    executor: Arc<dyn AsyncExecutor>,
    maximum_in_flight: usize,
) -> Result<(), Arc<dyn AsyncExecutor>> {
    let backend: Arc<SdkExcel12vBackend> = crate::runtime::production_backend();
    controller_cell()
        .set(AsyncController::new(
            executor.clone(),
            backend,
            maximum_in_flight,
        ))
        .map_err(|_| executor)
}

pub(crate) fn activate() {
    if let Some(controller) = controller_cell().get() {
        controller.activate();
    }
}

pub(crate) fn shutdown() {
    if let Some(controller) = controller_cell().get() {
        controller.shutdown();
    }
}

pub fn calculation_canceled() {
    if let Some(controller) = controller_cell().get() {
        controller.cancel_all();
    }
}

pub fn calculation_ended() {
    if let Some(controller) = controller_cell().get() {
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
    let controller = controller_cell()
        .get()
        .ok_or(AsyncSubmitError::ExecutorUnavailable)?;
    // SAFETY: forwarded from the generated async thunk contract.
    let handle = unsafe { AsyncHandle::copy_from(raw_handle) }?;
    controller.schedule(handle, task)
}

#[cfg(test)]
mod tests {
    use super::*;
    use core::ffi::c_int;
    use std::{collections::VecDeque, sync::atomic::AtomicI32};

    type RecordedCall = (i32, usize, usize, u32, f64);

    #[derive(Default)]
    struct QueueExecutor {
        jobs: Mutex<VecDeque<ExecutorJob>>,
        reject: AtomicBool,
        shutdowns: AtomicUsize,
    }

    impl QueueExecutor {
        fn run_one(&self) {
            let job = self
                .jobs
                .lock()
                .unwrap_or_else(std::sync::PoisonError::into_inner)
                .pop_front()
                .expect("queued job");
            job();
        }
    }

    impl AsyncExecutor for QueueExecutor {
        fn execute(&self, job: ExecutorJob) -> Result<(), AsyncSubmitError> {
            if self.reject.load(Ordering::Acquire) {
                Err(AsyncSubmitError::ExecutorRejected)
            } else {
                self.jobs
                    .lock()
                    .unwrap_or_else(std::sync::PoisonError::into_inner)
                    .push_back(job);
                Ok(())
            }
        }

        fn shutdown(&self) {
            self.shutdowns.fetch_add(1, Ordering::AcqRel);
        }
    }

    #[derive(Default)]
    struct MockBackend {
        linked: AtomicBool,
        code: AtomicI32,
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
            // SAFETY: test controller supplies two stable roots.
            let arguments = unsafe { core::slice::from_raw_parts(arguments, count as usize) };
            // SAFETY: exact xlAsyncReturn test arguments.
            let handle = unsafe { (*arguments[0]).val.bigdata.h.hdata } as usize;
            // SAFETY: the planned test result is numeric.
            let value = unsafe { (*arguments[1]).val.num };
            // SAFETY: the controller supplies a readable second root.
            let value_type = unsafe { (*arguments[1]).xltype };
            self.calls
                .lock()
                .unwrap_or_else(std::sync::PoisonError::into_inner)
                .push((function, count as usize, handle, value_type, value));
            if !result.is_null() {
                // SAFETY: result is the controller's live writable response root.
                unsafe {
                    (*result).xltype = excel_api_sys::xltypeBool;
                    (*result).val.xbool = i32::from(self.accepted.load(Ordering::Acquire));
                }
            }
            self.code.load(Ordering::Acquire)
        }
    }

    fn handle(value: usize) -> XLOPER12 {
        AsyncHandle { value, size: 0 }.root()
    }

    fn number(value: f64) -> ExcelReturnValue {
        ExcelReturnValue::from(crate::ExcelValue::Number(value))
    }

    fn controller(maximum: usize) -> (Arc<AsyncController>, Arc<QueueExecutor>, Arc<MockBackend>) {
        let executor = Arc::new(QueueExecutor::default());
        let backend = Arc::new(MockBackend::linked());
        let controller = AsyncController::new(executor.clone(), backend.clone(), maximum);
        controller.activate();
        (controller, executor, backend)
    }

    #[test]
    fn exact_handle_and_owned_result_are_passed_once_without_dllfree() {
        let (controller, executor, backend) = controller(2);
        let mut raw = handle(0x1234);
        // SAFETY: `raw` is a live, correctly tagged test handle root.
        let copied = unsafe { AsyncHandle::copy_from(&mut raw) }.unwrap();
        controller
            .schedule(copied, Box::new(|_| Ok(number(42.0))))
            .unwrap();
        executor.run_one();
        let calls = backend.calls.lock().unwrap();
        assert_eq!(calls.len(), 1);
        assert_eq!(calls[0].0, excel_api_sys::xlAsyncReturn);
        assert_eq!(calls[0].1, 2);
        assert_eq!(calls[0].2, 0x1234);
        assert_eq!(calls[0].3, excel_api_sys::xltypeNum);
        assert_eq!(calls[0].4, 42.0);
        assert_eq!(calls[0].3 & excel_api_sys::xlbitDLLFree, 0);
        assert_eq!(controller.in_flight.load(Ordering::Acquire), 0);
    }

    #[test]
    fn cancellation_and_shutdown_suppress_late_completion() {
        let (controller, executor, backend) = controller(2);
        controller
            .schedule(
                AsyncHandle { value: 1, size: 0 },
                Box::new(|_| Ok(number(1.0))),
            )
            .unwrap();
        controller.cancel_all();
        executor.run_one();
        assert!(backend.calls.lock().unwrap().is_empty());

        controller.activate();
        controller
            .schedule(
                AsyncHandle { value: 2, size: 0 },
                Box::new(|_| Ok(number(2.0))),
            )
            .unwrap();
        controller.shutdown();
        executor.run_one();
        assert!(backend.calls.lock().unwrap().is_empty());
        assert_eq!(executor.shutdowns.load(Ordering::Acquire), 1);
    }

    #[test]
    fn capacity_executor_rejection_and_invalid_handles_are_distinct() {
        let (controller, executor, _) = controller(1);
        controller
            .schedule(
                AsyncHandle { value: 1, size: 0 },
                Box::new(|_| Ok(number(1.0))),
            )
            .unwrap();
        assert_eq!(
            controller.schedule(
                AsyncHandle { value: 2, size: 0 },
                Box::new(|_| Ok(number(2.0)))
            ),
            Err(AsyncSubmitError::CapacityExhausted)
        );
        executor.run_one();
        executor.reject.store(true, Ordering::Release);
        assert_eq!(
            controller.schedule(
                AsyncHandle { value: 3, size: 0 },
                Box::new(|_| Ok(number(3.0)))
            ),
            Err(AsyncSubmitError::ExecutorRejected)
        );
        let mut wrong = XLOPER12 {
            val: XLOPER12Value { num: 0.0 },
            xltype: excel_api_sys::xltypeNum,
        };
        // SAFETY: `wrong` is readable; the test expects tag validation before
        // the union is interpreted as an async handle.
        let invalid = unsafe { AsyncHandle::copy_from(&mut wrong) };
        assert!(matches!(invalid, Err(AsyncSubmitError::InvalidHandle)));
    }

    #[test]
    fn excel_return_code_and_false_acceptance_are_preserved() {
        let (controller, _, backend) = controller(1);
        backend.code.store(
            excel_api_sys::xlretInvAsynchronousContext,
            Ordering::Release,
        );
        backend.accepted.store(false, Ordering::Release);
        let request = Arc::new(Request {
            id: 99,
            epoch: controller.epoch.load(Ordering::Acquire),
            handle: AsyncHandle { value: 9, size: 0 },
            state: AtomicU8::new(SCHEDULED),
            owner: Arc::downgrade(&controller),
        });
        controller
            .requests
            .lock()
            .unwrap()
            .insert(99, request.clone());
        controller.in_flight.store(1, Ordering::Release);
        let error = request.complete(number(9.0)).unwrap_err();
        assert!(matches!(
            error,
            AsyncCompletionError::Excel {
                code: ExcelReturnCode(excel_api_sys::xlretInvAsynchronousContext),
                accepted: Some(false)
            }
        ));
    }

    #[test]
    fn concurrent_completion_race_calls_excel_at_most_once() {
        let (controller, _, backend) = controller(8);
        let request = Arc::new(Request {
            id: 100,
            epoch: controller.epoch.load(Ordering::Acquire),
            handle: AsyncHandle {
                value: 100,
                size: 0,
            },
            state: AtomicU8::new(SCHEDULED),
            owner: Arc::downgrade(&controller),
        });
        controller
            .requests
            .lock()
            .unwrap()
            .insert(100, request.clone());
        controller.in_flight.store(1, Ordering::Release);
        let barrier = Arc::new(std::sync::Barrier::new(8));
        let joins: Vec<_> = (0..8)
            .map(|index| {
                let request = request.clone();
                let barrier = barrier.clone();
                std::thread::spawn(move || {
                    barrier.wait();
                    request.complete(number(index as f64))
                })
            })
            .collect();
        let successes = joins
            .into_iter()
            .map(|join| join.join().unwrap())
            .filter(Result::is_ok)
            .count();
        assert_eq!(successes, 1);
        assert_eq!(backend.calls.lock().unwrap().len(), 1);
        assert_eq!(controller.in_flight.load(Ordering::Acquire), 0);
    }
}
