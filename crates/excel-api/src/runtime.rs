//! Idempotent XLL lifecycle runtime.

use core::fmt;
use std::sync::{Arc, Mutex, OnceLock};

use crate::excel_call::{CallCapability, ExcelCallBackend, ExcelCallError, SdkExcel12vBackend};
use crate::{AddInDescriptor, AsyncSubmitError, ExcelString, LifecycleContext};

#[derive(Clone, Copy)]
struct RegisteredEntry {
    id: f64,
    name: &'static str,
    kind: RegistrationKind,
    name_pending: bool,
    id_pending: bool,
}

#[derive(Clone, Copy, Eq, PartialEq)]
enum RegistrationKind {
    Worksheet,
    Command,
}

/// Observable state of one XLL runtime lifecycle generation.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum RuntimePhase {
    /// No Excel callback backend is linked and no registrations are tracked.
    Uninitialized,
    /// Initialization owns the transition and is registering functions.
    Initializing,
    /// Registrations and the current runtime generation are active.
    Initialized,
    /// Close owns the transition and is retiring active resources.
    Closing,
    /// Excel registrations remain and close must be retried. Async scheduling
    /// is disabled and the callback backend remains linked for cleanup only.
    CleanupRequired,
}

/// Counters retained by [`Runtime`] for lifecycle diagnostics.
#[derive(Clone, Debug, Default, Eq, PartialEq)]
pub struct RuntimeDiagnostics {
    /// Number of calls to [`Runtime::initialize`].
    pub initialization_attempts: usize,
    /// Initialization transitions that completed successfully.
    pub successful_initializations: usize,
    /// Idempotent initialize calls while already initialized.
    pub duplicate_initializations: usize,
    /// Number of calls to [`Runtime::close`].
    pub close_attempts: usize,
    /// Close transitions that unlinked the backend.
    pub successful_closes: usize,
    /// Total successfully registered functions and commands.
    pub registrations: usize,
    /// Number of attempted Excel unregister calls.
    pub unregister_attempts: usize,
    /// Number of rollback unregister attempts after initialization failure.
    pub rollback_attempts: usize,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
/// Outcome for an idempotent lifecycle operation.
pub enum LifecycleOutcome {
    /// This call completed the requested lifecycle transition.
    Completed,
    /// The runtime was already at the requested stable state.
    AlreadyInState,
}

#[derive(Clone, Debug, Eq, PartialEq)]
/// A lifecycle transition could not be completed safely.
pub enum LifecycleError {
    /// Another lifecycle transition owns the runtime in this phase.
    Busy(RuntimePhase),
    /// Async generation activation failed.
    Async(AsyncSubmitError),
    /// A typed Excel lifecycle call failed.
    Excel(ExcelCallError),
    /// Registration failed and one or more compensating calls also failed.
    RegistrationRollback {
        /// Original registration failure.
        primary: ExcelCallError,
        /// Failed compensating delete-name or unregister calls.
        cleanup: Vec<ExcelCallError>,
    },
    /// Closing could not unregister all remaining registrations.
    CloseFailed {
        /// Exact errors retained for a later cleanup retry.
        errors: Vec<ExcelCallError>,
    },
}

impl fmt::Display for LifecycleError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Busy(phase) => write!(f, "runtime is busy in {phase:?}"),
            Self::Async(error) => write!(f, "asynchronous runtime activation failed: {error}"),
            Self::Excel(error) => write!(f, "lifecycle Excel call failed: {error}"),
            Self::RegistrationRollback { primary, cleanup } => write!(
                f,
                "registration failed ({primary}); {} rollback calls failed",
                cleanup.len()
            ),
            Self::CloseFailed { errors } => {
                write!(f, "{} unregister calls failed during close", errors.len())
            }
        }
    }
}
impl std::error::Error for LifecycleError {}

impl From<ExcelCallError> for LifecycleError {
    fn from(value: ExcelCallError) -> Self {
        Self::Excel(value)
    }
}

struct RuntimeState {
    phase: RuntimePhase,
    module: Option<ExcelString>,
    worksheet_registrations: Vec<RegisteredEntry>,
    command_registrations: Vec<RegisteredEntry>,
    diagnostics: RuntimeDiagnostics,
    calculation_canceled_registered: bool,
    calculation_ended_registered: bool,
    async_generation_active: bool,
    dispatcher_generation_active: bool,
}

impl Default for RuntimeState {
    fn default() -> Self {
        Self {
            phase: RuntimePhase::Uninitialized,
            module: None,
            worksheet_registrations: Vec::new(),
            command_registrations: Vec::new(),
            diagnostics: RuntimeDiagnostics::default(),
            calculation_canceled_registered: false,
            calculation_ended_registered: false,
            async_generation_active: false,
            dispatcher_generation_active: false,
        }
    }
}

/// Coordinates Excel linking, registration, and orderly XLL close.
///
/// A runtime is process-local application state, not a callback context. Its
/// transition methods must be called through genuine Excel lifecycle thunks.
pub struct Runtime {
    backend: Arc<dyn ExcelCallBackend>,
    sdk_backend: Option<Arc<SdkExcel12vBackend>>,
    state: Mutex<RuntimeState>,
}

impl fmt::Debug for Runtime {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("Runtime")
            .field("phase", &self.phase())
            .finish_non_exhaustive()
    }
}

impl Runtime {
    /// Creates the process-local runtime using the production Excel12v backend.
    ///
    /// The backend remains unlinked until [`Self::initialize`] runs from a
    /// genuine lifecycle callback.
    pub fn production() -> Self {
        let backend = production_backend();
        Self {
            backend: backend.clone(),
            sdk_backend: Some(backend),
            state: Mutex::new(RuntimeState::default()),
        }
    }

    #[cfg(test)]
    fn with_backend(backend: Arc<dyn ExcelCallBackend>) -> Self {
        Self {
            backend,
            sdk_backend: None,
            state: Mutex::new(RuntimeState::default()),
        }
    }

    /// Installs Excel's `Excel12v` callback entry point in the SDK backend.
    ///
    /// Generated lifecycle thunks call this before initialization; passing an
    /// entry point does not itself link the backend or make arbitrary calls legal.
    pub fn set_excel12_entry_point(&self, callback: excel_api_sys::Excel12EntryPtFn) {
        if let Some(backend) = &self.sdk_backend {
            backend.set_entry_point(callback);
        }
    }

    fn lock(&self) -> std::sync::MutexGuard<'_, RuntimeState> {
        self.state
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner)
    }

    /// Returns the current lifecycle phase.
    pub fn phase(&self) -> RuntimePhase {
        self.lock().phase
    }
    /// Returns a copied snapshot of process-local lifecycle counters.
    pub fn diagnostics(&self) -> RuntimeDiagnostics {
        self.lock().diagnostics.clone()
    }
    /// Returns worksheet-function registration identifiers currently tracked.
    ///
    /// The IDs are meaningful only while the current runtime generation remains
    /// active and must not be used to issue raw unregister calls directly.
    pub fn registration_ids(&self) -> Vec<f64> {
        self.lock()
            .worksheet_registrations
            .iter()
            .map(|entry| entry.id)
            .collect()
    }

    /// Runs an M17 compatibility-spike operation with a lifecycle capability.
    ///
    /// This bridge is compiled only for the isolated xlcOnTime research XLL.
    /// It does not make lifecycle use of an XLM command a stable or generally
    /// supported capability.
    ///
    /// # Safety
    ///
    /// The caller must be synchronously executing on the callback thread in a
    /// genuine Excel-issued lifecycle callback. A linked backend alone is not
    /// evidence that this precondition holds.
    ///
    /// Safe code cannot invoke this bridge:
    ///
    /// ```compile_fail
    /// let runtime = excel_api::Runtime::production();
    /// runtime.experimental_with_lifecycle_context(|_| ());
    /// ```
    #[cfg(feature = "xlcontime-research")]
    #[doc(hidden)]
    pub unsafe fn experimental_with_lifecycle_context<R>(
        &self,
        body: impl FnOnce(&LifecycleContext<'_>) -> R,
    ) -> Result<R, ExcelCallError> {
        if !self.backend.is_linked() {
            return Err(ExcelCallError::BackendUnavailable);
        }
        let capability = CallCapability::new(self.backend.as_ref());
        let context = LifecycleContext::new(&capability);
        Ok(body(&context))
    }

    /// Links Excel, installs a fresh generation, and registers `add_in`.
    ///
    /// Call only from an Excel-issued open/lifecycle callback. On partial
    /// registration failure this performs bounded rollback; a rollback failure
    /// leaves [`RuntimePhase::CleanupRequired`] so [`Self::close`] can retry.
    pub fn initialize(&self, add_in: &AddInDescriptor) -> Result<LifecycleOutcome, LifecycleError> {
        self.initialize_with_help(add_in, &[])
    }

    /// Registers optional worksheet help supplements in an Excel-issued open callback.
    ///
    /// Invalid/duplicate mappings fail before linking or dispatch. Existing initialized
    /// generations are unchanged; close/reopen is required to apply new metadata.
    pub fn initialize_with_help(
        &self,
        add_in: &AddInDescriptor,
        help: &[crate::FunctionHelp],
    ) -> Result<LifecycleOutcome, LifecycleError> {
        for (index, entry) in help.iter().enumerate() {
            entry
                .topic
                .validate()
                .map_err(|e| LifecycleError::Excel(ExcelCallError::Registration(e)))?;
            if add_in
                .functions
                .iter()
                .filter(|f| f.rust_symbol == entry.rust_symbol)
                .count()
                != 1
                || help[..index]
                    .iter()
                    .any(|h| h.rust_symbol == entry.rust_symbol)
            {
                return Err(ExcelCallError::Registration(
                    crate::RegistrationError::InvalidHelpMapping,
                )
                .into());
            }
        }
        let _callback = crate::dispatcher::enter_callback();
        add_in
            .validate()
            .map_err(|error| LifecycleError::Excel(ExcelCallError::Registration(error)))?;
        {
            let mut state = self.lock();
            state.diagnostics.initialization_attempts += 1;
            match state.phase {
                RuntimePhase::Initialized => {
                    state.diagnostics.duplicate_initializations += 1;
                    return Ok(LifecycleOutcome::AlreadyInState);
                }
                RuntimePhase::Uninitialized => state.phase = RuntimePhase::Initializing,
                phase => return Err(LifecycleError::Busy(phase)),
            }
        }

        let has_async = add_in
            .functions
            .iter()
            .any(|function| function.is_asynchronous());
        if let Err(error) = self.backend.link() {
            self.lock().phase = RuntimePhase::Uninitialized;
            return Err(error.into());
        }
        if has_async {
            if let Err(error) = crate::async_udf::activate(self.backend.clone()) {
                self.backend.unlink();
                self.lock().phase = RuntimePhase::Uninitialized;
                return Err(LifecycleError::Async(error));
            }
            self.lock().async_generation_active = true;
        }
        let dispatcher_active = crate::dispatcher::activate();
        self.lock().dispatcher_generation_active = dispatcher_active;

        let capability = CallCapability::new(self.backend.as_ref());
        let context = LifecycleContext::new(&capability);
        let module = match context.get_module_name() {
            Ok(module) => module,
            Err(error) => {
                if dispatcher_active {
                    crate::dispatcher::shutdown();
                }
                if has_async {
                    crate::async_udf::shutdown();
                }
                self.backend.unlink();
                let mut state = self.lock();
                state.phase = RuntimePhase::Uninitialized;
                state.async_generation_active = false;
                state.dispatcher_generation_active = false;
                return Err(error.into());
            }
        };

        if has_async {
            if let Err(error) = self.ensure_async_events(&context) {
                if dispatcher_active {
                    crate::dispatcher::shutdown();
                }
                crate::async_udf::shutdown();
                self.backend.unlink();
                let mut state = self.lock();
                state.phase = RuntimePhase::Uninitialized;
                state.async_generation_active = false;
                state.dispatcher_generation_active = false;
                return Err(error.into());
            }
        }

        let functions = add_in.functions.iter().map(|function| {
            let topic = help
                .iter()
                .find(|h| h.rust_symbol == function.rust_symbol)
                .map(|h| h.topic);
            (
                RegistrationKind::Worksheet,
                function.excel_name,
                context.register(&module, function, topic),
            )
        });
        let commands = add_in.commands.iter().map(|command| {
            (
                RegistrationKind::Command,
                command.excel_name,
                context.register_command(&module, command),
            )
        });
        let mut entries = Vec::new();
        for (kind, name, result) in functions.chain(commands) {
            let id = match &result {
                Ok(id) => Some(*id),
                Err(ExcelCallError::RegistrationRelease { id_bits, .. }) => {
                    Some(f64::from_bits(*id_bits))
                }
                _ => None,
            };
            if let Some(id) = id {
                entries.push(RegisteredEntry {
                    id,
                    name,
                    kind,
                    name_pending: true,
                    id_pending: true,
                });
            }
            if let Err(primary) = result {
                if has_async {
                    crate::async_udf::shutdown();
                }
                if dispatcher_active {
                    crate::dispatcher::shutdown();
                }
                let (cleanup, attempts) = if primary.requires_return_to_excel() {
                    // Preserve obligations for a later genuine lifecycle callback.
                    (vec![primary.clone()], 0)
                } else {
                    cleanup_entries(&context, &mut entries)
                };
                if cleanup.is_empty() {
                    self.backend.unlink();
                }
                let mut state = self.lock();
                state.diagnostics.rollback_attempts += attempts;
                state.async_generation_active = false;
                state.dispatcher_generation_active = false;
                state.phase = if cleanup.is_empty() {
                    RuntimePhase::Uninitialized
                } else {
                    RuntimePhase::CleanupRequired
                };
                state.module = if cleanup.is_empty() {
                    None
                } else {
                    Some(module)
                };
                state.worksheet_registrations = entries
                    .iter()
                    .copied()
                    .filter(|e| e.kind == RegistrationKind::Worksheet)
                    .collect();
                state.command_registrations = entries
                    .into_iter()
                    .filter(|e| e.kind == RegistrationKind::Command)
                    .collect();
                return Err(LifecycleError::RegistrationRollback { primary, cleanup });
            }
        }
        let mut state = self.lock();
        state.phase = RuntimePhase::Initialized;
        state.module = Some(module);
        state.diagnostics.registrations += entries.len();
        state.worksheet_registrations = entries
            .iter()
            .copied()
            .filter(|e| e.kind == RegistrationKind::Worksheet)
            .collect();
        state.command_registrations = entries
            .into_iter()
            .filter(|e| e.kind == RegistrationKind::Command)
            .collect();
        state.diagnostics.successful_initializations += 1;
        Ok(LifecycleOutcome::Completed)
    }

    fn ensure_async_events(&self, context: &LifecycleContext<'_>) -> Result<(), ExcelCallError> {
        let (canceled, ended) = {
            let state = self.lock();
            (
                state.calculation_canceled_registered,
                state.calculation_ended_registered,
            )
        };
        if !canceled {
            context.register_event(
                "excel_api_calculation_canceled",
                excel_api_sys::xleventCalculationCanceled,
            )?;
            self.lock().calculation_canceled_registered = true;
        }
        if !ended {
            context.register_event(
                "excel_api_calculation_ended",
                excel_api_sys::xleventCalculationEnded,
            )?;
            self.lock().calculation_ended_registered = true;
        }
        Ok(())
    }

    /// Retires dispatcher and async generations, unregisters, then unlinks Excel.
    ///
    /// This must run from a genuine lifecycle callback. Failed unregister calls
    /// retain conservative cleanup state and return [`LifecycleError::CloseFailed`];
    /// retrying `close` never reactivates old generations.
    pub fn close(&self) -> Result<LifecycleOutcome, LifecycleError> {
        let _callback = crate::dispatcher::enter_callback();
        let (registrations, drain_async, drain_dispatcher) = {
            let mut state = self.lock();
            state.diagnostics.close_attempts += 1;
            match state.phase {
                RuntimePhase::Uninitialized => return Ok(LifecycleOutcome::AlreadyInState),
                RuntimePhase::Initialized => {
                    state.phase = RuntimePhase::Closing;
                    let mut entries = state.command_registrations.clone();
                    entries.extend(state.worksheet_registrations.iter().copied());
                    let drain_async = state.async_generation_active;
                    let drain_dispatcher = state.dispatcher_generation_active;
                    state.async_generation_active = false;
                    state.dispatcher_generation_active = false;
                    (entries, drain_async, drain_dispatcher)
                }
                RuntimePhase::CleanupRequired => {
                    state.phase = RuntimePhase::Closing;
                    let mut entries = state.command_registrations.clone();
                    entries.extend(state.worksheet_registrations.iter().copied());
                    (entries, false, false)
                }
                phase => return Err(LifecycleError::Busy(phase)),
            }
        };

        // Remove and synchronously retire cooperative dispatch before any
        // registration or callback entry point can disappear. Shutdown waits
        // only for work already executing in a real callback; it never waits
        // for a future pump.
        if drain_dispatcher {
            crate::dispatcher::shutdown();
        }
        // Disable and drain asynchronous work before registrations or the
        // Excel callback entry point can disappear.
        if drain_async {
            crate::async_udf::shutdown();
        }

        let capability = CallCapability::new(self.backend.as_ref());
        let context = LifecycleContext::new(&capability);
        let mut failed = registrations;
        let (errors, attempts) = cleanup_entries(&context, &mut failed);
        let mut state = self.lock();
        state.diagnostics.unregister_attempts += attempts;
        if errors.is_empty() {
            self.backend.unlink();
            state.phase = RuntimePhase::Uninitialized;
            state.module = None;
            state.worksheet_registrations.clear();
            state.command_registrations.clear();
            state.diagnostics.successful_closes += 1;
            Ok(LifecycleOutcome::Completed)
        } else {
            state.phase = RuntimePhase::CleanupRequired;
            state.worksheet_registrations = failed
                .iter()
                .copied()
                .filter(|entry| entry.kind == RegistrationKind::Worksheet)
                .collect();
            state.command_registrations = failed
                .into_iter()
                .filter(|entry| entry.kind == RegistrationKind::Command)
                .collect();
            Err(LifecycleError::CloseFailed { errors })
        }
    }
}

// Keep name deletion and unregister obligations independent. An accepted operation
// must not be repeated merely because xlFree subsequently failed.
fn cleanup_entries(
    context: &LifecycleContext<'_>,
    entries: &mut Vec<RegisteredEntry>,
) -> (Vec<ExcelCallError>, usize) {
    let mut errors = Vec::new();
    let mut unregister_attempts = 0;
    'entries: for entry in entries.iter_mut().rev() {
        for delete_name in [true, false] {
            let pending = if delete_name {
                &mut entry.name_pending
            } else {
                &mut entry.id_pending
            };
            if !*pending {
                continue;
            }
            let result = if delete_name {
                context.delete_defined_name(entry.name)
            } else {
                unregister_attempts += 1;
                context.unregister(entry.id)
            };
            match result {
                Ok(()) => *pending = false,
                Err(error) => {
                    if matches!(error, ExcelCallError::AcceptedCleanupRelease { .. }) {
                        *pending = false;
                    }
                    let must_return = error.requires_return_to_excel();
                    errors.push(error);
                    if must_return {
                        break 'entries;
                    }
                }
            }
        }
    }
    entries.retain(|e| e.name_pending || e.id_pending);
    (errors, unregister_attempts)
}

pub(crate) fn production_backend() -> Arc<SdkExcel12vBackend> {
    static BACKEND: OnceLock<Arc<SdkExcel12vBackend>> = OnceLock::new();
    BACKEND
        .get_or_init(|| Arc::new(SdkExcel12vBackend::new()))
        .clone()
}

impl Default for Runtime {
    fn default() -> Self {
        Self::production()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{
        ExcelArgumentType, ExcelReturnType, FunctionFlags, FunctionRegistration, FunctionSignature,
    };
    use core::ffi::c_int;
    use excel_api_sys::{LPXLOPER12, XLOPER12, XLOPER12Value};
    use std::{collections::HashMap, sync::Mutex};

    #[derive(Default)]
    struct MockState {
        linked: bool,
        calls: Vec<(i32, usize)>,
        type_texts: Vec<String>,
        registration_arguments: Vec<Vec<(u32, String)>>,
        malformed_registration: Option<usize>,
        unregistered: Vec<f64>,
        allocations: HashMap<usize, Box<[u16]>>,
        next_id: f64,
        fail_registration: Option<usize>,
        fail_unregistration: usize,
        fail_event: Option<usize>,
        event_result: Option<i32>,
        register_count: usize,
        event_count: usize,
        event_ids: Vec<i32>,
        unregister_linked: Vec<bool>,
        free_calls: usize,
        fail_free: Option<usize>,
        free_status: Option<i32>,
        reject_cleanup: Option<(i32, u32)>,
        cleanup_code: Option<(i32, i32)>,
    }

    #[derive(Default)]
    struct MockBackend {
        state: Mutex<MockState>,
    }

    impl MockBackend {
        fn failing_registration(index: usize) -> Self {
            let mock = Self::default();
            mock.state.lock().unwrap().fail_registration = Some(index);
            mock
        }
    }

    impl ExcelCallBackend for MockBackend {
        fn link(&self) -> Result<(), ExcelCallError> {
            self.state.lock().unwrap().linked = true;
            Ok(())
        }
        fn unlink(&self) {
            self.state.lock().unwrap().linked = false;
        }
        fn is_linked(&self) -> bool {
            self.state.lock().unwrap().linked
        }

        unsafe fn excel12v_raw(
            &self,
            function: i32,
            result: *mut XLOPER12,
            count: c_int,
            arguments: *mut LPXLOPER12,
        ) -> i32 {
            let mut state = self.state.lock().unwrap();
            state.calls.push((function, count as usize));
            if let Some((target, code)) = state.cleanup_code {
                if target == function {
                    return code;
                }
            }
            if function == excel_api_sys::xlGetName {
                let mut text: Box<[u16]> = [vec![15], "C:\\test\\add.xll".encode_utf16().collect()]
                    .concat()
                    .into_boxed_slice();
                let pointer = text.as_mut_ptr();
                state.allocations.insert(pointer as usize, text);
                // SAFETY: the runtime supplies a live result root.
                unsafe {
                    *result = XLOPER12 {
                        val: XLOPER12Value { str: pointer },
                        xltype: excel_api_sys::xltypeStr,
                    };
                }
            } else if function == excel_api_sys::xlfRegister {
                // SAFETY: the synchronous call keeps all `count` roots and counted strings live.
                let snapshot = unsafe { core::slice::from_raw_parts(arguments, count as usize) }
                    .iter()
                    .map(|pointer| {
                        // SAFETY: registration supplies initialized roots with string or number tags.
                        let root = unsafe { &**pointer };
                        let text = match root.xltype {
                            excel_api_sys::xltypeStr => {
                                // SAFETY: CallArguments owns the prefix and complete UTF-16 payload.
                                let pointer = unsafe { root.val.str };
                                // SAFETY: the prefix describes the live counted allocation.
                                let units = unsafe {
                                    core::slice::from_raw_parts(pointer.add(1), *pointer as usize)
                                };
                                String::from_utf16(units).unwrap()
                            }
                            excel_api_sys::xltypeNum => {
                                // SAFETY: the number tag selects the numeric union member.
                                unsafe { root.val.num }.to_string()
                            }
                            other => panic!("unexpected registration tag {other}"),
                        };
                        (root.xltype, text)
                    })
                    .collect();
                state.registration_arguments.push(snapshot);
                state.register_count += 1;
                if state.fail_registration == Some(state.register_count) {
                    return excel_api_sys::xlretFailed;
                }
                // SAFETY: registration always supplies at least three initialized argument roots.
                let type_root = unsafe { *arguments.add(2) };
                // SAFETY: the third root is a counted string kept live by CallArguments.
                let type_pointer = unsafe { (*type_root).val.str };
                // SAFETY: counted prefix and payload are initialized.
                let length = unsafe { *type_pointer } as usize;
                // SAFETY: the payload has `length` code units.
                let units = unsafe { core::slice::from_raw_parts(type_pointer.add(1), length) };
                state.type_texts.push(String::from_utf16(units).unwrap());
                if state.malformed_registration == Some(state.register_count) {
                    // SAFETY: the caller supplies a live result root; simulate Excel #VALUE!.
                    unsafe {
                        *result = XLOPER12 {
                            val: XLOPER12Value {
                                err: excel_api_sys::xlerrValue,
                            },
                            xltype: excel_api_sys::xltypeErr,
                        };
                    }
                    return excel_api_sys::xlretSuccess;
                }
                state.next_id += 1.0;
                // SAFETY: the runtime supplies a live result root.
                unsafe {
                    *result = XLOPER12 {
                        val: XLOPER12Value { num: state.next_id },
                        xltype: excel_api_sys::xltypeNum,
                    };
                }
            } else if function == excel_api_sys::xlfUnregister {
                let linked = state.linked;
                state.unregister_linked.push(linked);
                if state.fail_unregistration > 0 {
                    state.fail_unregistration -= 1;
                    return excel_api_sys::xlretFailed;
                }
                // SAFETY: unregister supplies one number root.
                let id = unsafe { (**arguments).val.num };
                state.unregistered.push(id);
                // SAFETY: the runtime supplies a live result root.
                unsafe {
                    *result = XLOPER12 {
                        val: XLOPER12Value { xbool: 1 },
                        xltype: excel_api_sys::xltypeBool,
                    };
                }
            } else if function == excel_api_sys::xlfSetName {
                // SAFETY: the runtime supplies a live result root.
                unsafe {
                    *result = XLOPER12 {
                        val: XLOPER12Value { xbool: 1 },
                        xltype: excel_api_sys::xltypeBool,
                    };
                }
            } else if function == excel_api_sys::xlEventRegister {
                state.event_count += 1;
                if state.fail_event == Some(state.event_count) {
                    return excel_api_sys::xlretFailed;
                }
                // SAFETY: event registration supplies a string followed by an integer.
                let event = unsafe { (**arguments.add(1)).val.w };
                state.event_ids.push(event);
                // SAFETY: the runtime supplies a live result root.
                unsafe {
                    *result = XLOPER12 {
                        val: XLOPER12Value {
                            w: state.event_result.unwrap_or(state.event_count as i32),
                        },
                        xltype: excel_api_sys::xltypeInt,
                    };
                }
            } else if function == excel_api_sys::xlFree {
                state.free_calls += 1;
                if state.fail_free == Some(state.free_calls) {
                    return state.free_status.unwrap_or(excel_api_sys::xlretFailed);
                }
                // SAFETY: xlFree receives one live C-API result root.
                let root = unsafe { *arguments };
                // SAFETY: root remains live for this synchronous call.
                if unsafe { (*root).xltype & excel_api_sys::XLTYPE_MASK }
                    == excel_api_sys::xltypeStr
                {
                    // SAFETY: active union member is string.
                    let address = unsafe { (*root).val.str } as usize;
                    state.allocations.remove(&address);
                }
            }
            if let Some((target, tag)) = state.reject_cleanup {
                if function == target {
                    // SAFETY: cleanup supplies a live result; immediate rejection fixture.
                    unsafe {
                        *result = XLOPER12 {
                            val: XLOPER12Value { err: 0 },
                            xltype: tag,
                        };
                    }
                }
            }
            excel_api_sys::xlretSuccess
        }
    }

    const EMPTY: &[ExcelArgumentType] = &[];
    static FUNCTIONS: &[FunctionRegistration] = &[
        FunctionRegistration::new(
            "one",
            "ONE",
            FunctionSignature::new(ExcelReturnType::Xloper12, EMPTY),
        )
        .arguments(&[], &[]),
        FunctionRegistration::new(
            "two",
            "TWO",
            FunctionSignature::new(ExcelReturnType::Xloper12, EMPTY),
        )
        .arguments(&[], &[])
        .flags(FunctionFlags {
            volatile: false,
            thread_safe: true,
            macro_type: false,
            cluster_safe: false,
        }),
    ];
    static ADD_IN: AddInDescriptor = AddInDescriptor::new("test", "test", FUNCTIONS);

    static COMMANDS: &[crate::CommandRegistration] = &[
        crate::CommandRegistration::new("run_command", "RUST.RUN")
            .shortcut("A")
            .description("Run command"),
        crate::CommandRegistration::new("plain_command", "RUST.PLAIN"),
    ];
    static COMMAND_ADD_IN: AddInDescriptor =
        AddInDescriptor::new("commands", "commands", &[]).commands(COMMANDS);
    static MIXED_ADD_IN: AddInDescriptor =
        AddInDescriptor::new("mixed", "mixed", FUNCTIONS).commands(COMMANDS);

    #[test]
    fn help_supplements_fill_slot_nine_and_preserve_cleanup() {
        let _runtime = runtime_test_guard();
        let help = [crate::FunctionHelp::new(
            "one",
            crate::HelpTopic::new("guide.chm", 42),
        )];
        let backend = Arc::new(MockBackend::default());
        let runtime = Runtime::with_backend(backend.clone());
        runtime.initialize_with_help(&MIXED_ADD_IN, &help).unwrap();
        {
            let state = backend.state.lock().unwrap();
            let first = &state.registration_arguments[0];
            assert_eq!(first.len(), 10);
            assert_eq!(first[8], (excel_api_sys::xltypeStr, "guide.chm!42".into()));
            assert_eq!(first[5], (excel_api_sys::xltypeNum, "1".into()));
            assert_eq!(first[6].1, "Rust");
            assert_eq!(first[7].1, "");
            assert_eq!(first[9].1, "");
            assert_eq!(state.registration_arguments[1][8].1, "");
            assert_command_arguments(
                &state.registration_arguments[2],
                "A",
                "Run command",
                "run_command",
                "RUST.RUN",
            );
        }
        runtime.close().unwrap();
        assert_eq!(
            backend.state.lock().unwrap().unregistered,
            [2.0, 1.0, 4.0, 3.0]
        );

        let backend = Arc::new(MockBackend::failing_registration(2));
        let runtime = Runtime::with_backend(backend.clone());
        assert!(runtime.initialize_with_help(&ADD_IN, &help).is_err());
        let state = backend.state.lock().unwrap();
        assert_eq!(state.unregistered, [1.0]);
        assert!(state.allocations.is_empty());
    }

    #[test]
    fn invalid_help_and_mapping_fail_before_link_or_dispatch() {
        let _runtime = runtime_test_guard();
        let valid = crate::FunctionHelp::new("one", crate::HelpTopic::new("guide.chm", u32::MAX));
        for help in [
            vec![crate::FunctionHelp::new(
                "one",
                crate::HelpTopic::new("", 0),
            )],
            vec![crate::FunctionHelp::new("unknown", valid.topic)],
            vec![valid, valid],
            vec![crate::FunctionHelp::new(
                "one",
                crate::HelpTopic::new("https://example.test/help", 1),
            )],
        ] {
            let backend = Arc::new(MockBackend::default());
            let runtime = Runtime::with_backend(backend.clone());
            assert!(runtime.initialize_with_help(&ADD_IN, &help).is_err());
            let state = backend.state.lock().unwrap();
            assert!(!state.linked);
            assert!(state.calls.is_empty());
        }
    }

    fn assert_command_arguments(
        arguments: &[(u32, String)],
        shortcut: &str,
        help: &str,
        procedure: &str,
        name: &str,
    ) {
        let values = [
            "C:\\test\\add.xll",
            procedure,
            "I",
            name,
            "",
            "2",
            "Commands",
            shortcut,
            "",
            help,
        ];
        assert_eq!(arguments.len(), values.len());
        for (index, ((tag, actual), expected)) in arguments.iter().zip(values).enumerate() {
            assert_eq!(
                *tag,
                if index == 5 {
                    excel_api_sys::xltypeNum
                } else {
                    excel_api_sys::xltypeStr
                },
                "slot {}",
                index + 1
            );
            assert_eq!(actual, expected, "slot {}", index + 1);
        }
    }

    #[test]
    fn commands_dispatch_exact_positions_and_unregister() {
        let _runtime = runtime_test_guard();
        let backend = Arc::new(MockBackend::default());
        let runtime = Runtime::with_backend(backend.clone());
        assert_eq!(
            runtime.initialize(&COMMAND_ADD_IN),
            Ok(LifecycleOutcome::Completed)
        );
        assert_eq!(runtime.phase(), RuntimePhase::Initialized);
        assert_eq!(
            runtime
                .lock()
                .command_registrations
                .iter()
                .map(|entry| entry.id)
                .collect::<Vec<_>>(),
            vec![1.0, 2.0]
        );
        {
            let state = backend.state.lock().unwrap();
            assert_command_arguments(
                &state.registration_arguments[0],
                "A",
                "Run command",
                "run_command",
                "RUST.RUN",
            );
            assert_command_arguments(
                &state.registration_arguments[1],
                "",
                "",
                "plain_command",
                "RUST.PLAIN",
            );
            assert_eq!(
                state
                    .calls
                    .iter()
                    .filter(|(id, _)| *id == excel_api_sys::xlfRegister)
                    .copied()
                    .collect::<Vec<_>>(),
                vec![(excel_api_sys::xlfRegister, 10); 2]
            );
        }
        assert_eq!(runtime.close(), Ok(LifecycleOutcome::Completed));
        let state = backend.state.lock().unwrap();
        assert_eq!(state.unregistered, [2.0, 1.0]);
        assert_eq!(state.unregister_linked, [true, true]);
        assert_eq!(state.free_calls, 7);
        assert!(state.allocations.is_empty());
        assert!(!state.linked);
    }

    #[test]
    fn mixed_functions_and_commands_preserve_registration_order() {
        let _runtime = runtime_test_guard();
        let backend = Arc::new(MockBackend::default());
        let runtime = Runtime::with_backend(backend.clone());
        assert_eq!(
            runtime.initialize(&MIXED_ADD_IN),
            Ok(LifecycleOutcome::Completed)
        );
        {
            let state = backend.state.lock().unwrap();
            assert_eq!(state.type_texts, ["Q", "Q$", "I", "I"]);
            assert_eq!(
                state.registration_arguments[0][5],
                (excel_api_sys::xltypeNum, "1".into())
            );
            assert_eq!(state.registration_arguments[0][6].1, "Rust");
            assert_command_arguments(
                &state.registration_arguments[2],
                "A",
                "Run command",
                "run_command",
                "RUST.RUN",
            );
        }
        assert_eq!(runtime.registration_ids(), vec![1.0, 2.0]);
        assert_eq!(
            runtime
                .lock()
                .command_registrations
                .iter()
                .map(|entry| entry.id)
                .collect::<Vec<_>>(),
            vec![3.0, 4.0]
        );
        assert_eq!(runtime.close(), Ok(LifecycleOutcome::Completed));
        // Preserve the existing close path: worksheet group then command group, each reversed.
        assert_eq!(
            backend.state.lock().unwrap().unregistered,
            [2.0, 1.0, 4.0, 3.0]
        );
    }

    #[test]
    fn command_failures_roll_back_functions_and_prior_commands() {
        let _runtime = runtime_test_guard();
        for malformed in [false, true] {
            let backend = Arc::new(MockBackend::default());
            if malformed {
                backend.state.lock().unwrap().malformed_registration = Some(4);
            } else {
                backend.state.lock().unwrap().fail_registration = Some(4);
            }
            let runtime = Runtime::with_backend(backend.clone());
            let Err(LifecycleError::RegistrationRollback { primary, cleanup }) =
                runtime.initialize(&MIXED_ADD_IN)
            else {
                panic!("expected rollback");
            };
            if malformed {
                assert!(matches!(primary, ExcelCallError::MalformedResult { .. }));
            } else {
                assert!(matches!(primary, ExcelCallError::ExcelFailure { .. }));
            }
            assert!(cleanup.is_empty());
            assert_eq!(runtime.phase(), RuntimePhase::Uninitialized);
            let state = backend.state.lock().unwrap();
            assert_eq!(state.register_count, 4);
            assert_eq!(state.unregistered, [3.0, 2.0, 1.0]);
            assert!(state.allocations.is_empty());
            assert!(!state.linked);
        }
    }

    #[test]
    fn command_rollback_cleanup_failure_is_reported_and_close_retries() {
        let _runtime = runtime_test_guard();
        let backend = Arc::new(MockBackend::failing_registration(2));
        backend.state.lock().unwrap().fail_unregistration = 1;
        let runtime = Runtime::with_backend(backend.clone());
        let Err(LifecycleError::RegistrationRollback { cleanup, .. }) =
            runtime.initialize(&COMMAND_ADD_IN)
        else {
            panic!("expected rollback");
        };
        assert_eq!(cleanup.len(), 1);
        assert_eq!(runtime.phase(), RuntimePhase::CleanupRequired);
        assert!(backend.state.lock().unwrap().linked);
        assert_eq!(runtime.close(), Ok(LifecycleOutcome::Completed));
        assert_eq!(backend.state.lock().unwrap().unregistered, [1.0]);
        assert_eq!(runtime.phase(), RuntimePhase::Uninitialized);
    }

    #[test]
    fn invalid_command_metadata_never_reaches_backend() {
        let _runtime = runtime_test_guard();
        static INVALID: &[crate::CommandRegistration] =
            &[crate::CommandRegistration::new("", "BAD")];
        let add_in = AddInDescriptor::new("bad", "bad", &[]).commands(INVALID);
        let backend = Arc::new(MockBackend::default());
        let runtime = Runtime::with_backend(backend.clone());
        assert!(runtime.initialize(&add_in).is_err());
        assert!(backend.state.lock().unwrap().calls.is_empty());
    }

    const ASYNC_ARGS: &[ExcelArgumentType] =
        &[ExcelArgumentType::Number, ExcelArgumentType::AsyncHandle];
    static ASYNC_FUNCTIONS: &[FunctionRegistration] = &[FunctionRegistration::new(
        "async_one",
        "ASYNC.ONE",
        FunctionSignature::new(ExcelReturnType::AsyncVoid, ASYNC_ARGS),
    )
    .arguments(&["value"], &["Value"])
    .flags(FunctionFlags {
        volatile: false,
        thread_safe: true,
        macro_type: false,
        cluster_safe: false,
    })];
    static ASYNC_ADD_IN: AddInDescriptor =
        AddInDescriptor::new("async test", "async test", ASYNC_FUNCTIONS);

    fn install_test_executor() {
        crate::async_udf::reset_generations_for_test();
        let executor = crate::ThreadPoolExecutor::new(1, 4).unwrap();
        assert!(crate::install_async_executor(Arc::new(executor), 4).is_ok());
    }

    fn async_test_guard() -> std::sync::MutexGuard<'static, ()> {
        static GUARD: OnceLock<Mutex<()>> = OnceLock::new();
        GUARD
            .get_or_init(|| Mutex::new(()))
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner)
    }

    struct RuntimeTestGuard {
        _runtime: std::sync::MutexGuard<'static, ()>,
        _dispatcher: std::sync::MutexGuard<'static, ()>,
    }

    fn runtime_test_guard() -> RuntimeTestGuard {
        static GUARD: OnceLock<Mutex<()>> = OnceLock::new();
        let runtime = GUARD
            .get_or_init(|| Mutex::new(()))
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner);
        let dispatcher = crate::dispatcher::TEST_SERIAL
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner);
        RuntimeTestGuard {
            _runtime: runtime,
            _dispatcher: dispatcher,
        }
    }

    #[test]
    fn duplicate_sequences_register_once_and_close_is_idempotent() {
        let _runtime = runtime_test_guard();
        let backend = Arc::new(MockBackend::default());
        let runtime = Runtime::with_backend(backend.clone());
        assert_eq!(runtime.initialize(&ADD_IN), Ok(LifecycleOutcome::Completed));
        assert_eq!(
            runtime.initialize(&ADD_IN),
            Ok(LifecycleOutcome::AlreadyInState)
        );
        assert_eq!(runtime.registration_ids(), vec![1.0, 2.0]);
        assert_eq!(runtime.close(), Ok(LifecycleOutcome::Completed));
        assert_eq!(runtime.close(), Ok(LifecycleOutcome::AlreadyInState));
        let state = backend.state.lock().unwrap();
        assert_eq!(state.type_texts, ["Q", "Q$"]);
        assert_eq!(state.unregistered, [2.0, 1.0]);
        assert_eq!(state.unregister_linked, [true, true]);
        assert!(!state.linked);
        assert!(state.allocations.is_empty());
        assert_eq!(state.free_calls, 7); // xlGetName, registrations, name deletion, unregisters.
    }

    #[test]
    fn partial_registration_rolls_back_and_retry_succeeds() {
        let _runtime = runtime_test_guard();
        let backend = Arc::new(MockBackend::failing_registration(2));
        let runtime = Runtime::with_backend(backend.clone());
        assert!(matches!(
            runtime.initialize(&ADD_IN),
            Err(LifecycleError::RegistrationRollback { .. })
        ));
        assert_eq!(runtime.phase(), RuntimePhase::Uninitialized);
        assert_eq!(backend.state.lock().unwrap().unregistered, [1.0]);
        backend.state.lock().unwrap().fail_registration = None;
        assert_eq!(runtime.initialize(&ADD_IN), Ok(LifecycleOutcome::Completed));
    }

    #[test]
    fn unavailable_backend_restores_uninitialized_state() {
        let _runtime = runtime_test_guard();
        let runtime = Runtime::with_backend(Arc::new(
            crate::excel_call::test_support::UnavailableBackend,
        ));
        assert_eq!(
            runtime.initialize(&ADD_IN),
            Err(LifecycleError::Excel(ExcelCallError::BackendUnavailable))
        );
        assert_eq!(runtime.phase(), RuntimePhase::Uninitialized);
    }

    #[test]
    fn async_events_register_once_across_close_and_reopen() {
        let _runtime = runtime_test_guard();
        let _guard = async_test_guard();
        install_test_executor();
        let backend = Arc::new(MockBackend::default());
        let runtime = Runtime::with_backend(backend.clone());
        assert_eq!(
            runtime.initialize(&ASYNC_ADD_IN),
            Ok(LifecycleOutcome::Completed)
        );
        assert_eq!(runtime.close(), Ok(LifecycleOutcome::Completed));
        let executor = crate::ThreadPoolExecutor::new(1, 4).unwrap();
        assert!(crate::install_async_executor(Arc::new(executor), 4).is_ok());
        assert_eq!(
            runtime.initialize(&ASYNC_ADD_IN),
            Ok(LifecycleOutcome::Completed)
        );
        assert_eq!(runtime.close(), Ok(LifecycleOutcome::Completed));
        let state = backend.state.lock().unwrap();
        assert_eq!(state.event_count, 2);
        assert_eq!(
            state.event_ids,
            [2, 1] // Independent pinned XLCALL.H: canceled=2, ended=1.
        );
        crate::async_udf::reset_generations_for_test();
    }

    #[test]
    fn failed_second_event_registration_retries_only_missing_event() {
        let _runtime = runtime_test_guard();
        let _guard = async_test_guard();
        install_test_executor();
        let backend = Arc::new(MockBackend::default());
        backend.state.lock().unwrap().fail_event = Some(2);
        let runtime = Runtime::with_backend(backend.clone());
        assert!(matches!(
            runtime.initialize(&ASYNC_ADD_IN),
            Err(LifecycleError::Excel(_))
        ));
        assert_eq!(runtime.phase(), RuntimePhase::Uninitialized);
        backend.state.lock().unwrap().fail_event = None;
        let executor = crate::ThreadPoolExecutor::new(1, 4).unwrap();
        assert!(crate::install_async_executor(Arc::new(executor), 4).is_ok());
        assert_eq!(
            runtime.initialize(&ASYNC_ADD_IN),
            Ok(LifecycleOutcome::Completed)
        );
        assert_eq!(backend.state.lock().unwrap().event_count, 3);
        assert_eq!(runtime.close(), Ok(LifecycleOutcome::Completed));
        crate::async_udf::reset_generations_for_test();
    }

    #[test]
    fn close_failure_is_cleanup_required_and_retry_can_reopen() {
        let _runtime = runtime_test_guard();
        let _guard = async_test_guard();
        install_test_executor();
        let backend = Arc::new(MockBackend::default());
        let runtime = Runtime::with_backend(backend.clone());
        assert_eq!(
            runtime.initialize(&ASYNC_ADD_IN),
            Ok(LifecycleOutcome::Completed)
        );
        backend.state.lock().unwrap().fail_unregistration = 1;
        assert!(matches!(
            runtime.close(),
            Err(LifecycleError::CloseFailed { .. })
        ));
        assert_eq!(runtime.phase(), RuntimePhase::CleanupRequired);
        assert!(matches!(
            runtime.initialize(&ASYNC_ADD_IN),
            Err(LifecycleError::Busy(RuntimePhase::CleanupRequired))
        ));
        assert_eq!(runtime.close(), Ok(LifecycleOutcome::Completed));
        assert_eq!(runtime.phase(), RuntimePhase::Uninitialized);
        let executor = crate::ThreadPoolExecutor::new(1, 4).unwrap();
        assert!(crate::install_async_executor(Arc::new(executor), 4).is_ok());
        assert_eq!(
            runtime.initialize(&ASYNC_ADD_IN),
            Ok(LifecycleOutcome::Completed)
        );
        assert_eq!(runtime.close(), Ok(LifecycleOutcome::Completed));
        crate::async_udf::reset_generations_for_test();
    }

    #[test]
    fn initialization_failure_after_events_shuts_generation_and_reuses_events() {
        let _runtime = runtime_test_guard();
        let _guard = async_test_guard();
        install_test_executor();
        let backend = Arc::new(MockBackend::failing_registration(1));
        let runtime = Runtime::with_backend(backend.clone());
        assert!(matches!(
            runtime.initialize(&ASYNC_ADD_IN),
            Err(LifecycleError::RegistrationRollback { .. })
        ));
        assert_eq!(runtime.phase(), RuntimePhase::Uninitialized);
        backend.state.lock().unwrap().fail_registration = None;
        let executor = crate::ThreadPoolExecutor::new(1, 4).unwrap();
        assert!(crate::install_async_executor(Arc::new(executor), 4).is_ok());
        assert_eq!(
            runtime.initialize(&ASYNC_ADD_IN),
            Ok(LifecycleOutcome::Completed)
        );
        assert_eq!(backend.state.lock().unwrap().event_count, 2);
        assert_eq!(runtime.close(), Ok(LifecycleOutcome::Completed));
        crate::async_udf::reset_generations_for_test();
    }

    #[test]
    fn dispatcher_shutdown_precedes_cleanup_required_and_backend_unlink() {
        let _runtime = runtime_test_guard();
        crate::dispatcher::reset_generations_for_test();
        assert_eq!(
            crate::install_dispatcher(crate::DispatchConfig::default()),
            Ok(())
        );
        let backend = Arc::new(MockBackend::default());
        let runtime = Runtime::with_backend(backend.clone());
        assert_eq!(runtime.initialize(&ADD_IN), Ok(LifecycleOutcome::Completed));
        let ticket = crate::enqueue_dispatch(crate::DispatchOperation::EchoOwned(
            crate::ExcelValue::Integer(7),
        ))
        .unwrap();
        backend.state.lock().unwrap().fail_unregistration = 1;
        assert!(matches!(
            runtime.close(),
            Err(LifecycleError::CloseFailed { .. })
        ));
        assert_eq!(runtime.phase(), RuntimePhase::CleanupRequired);
        assert_eq!(
            ticket.try_result(),
            Some(Err(crate::DispatchCompletionError::DispatcherShutdown))
        );
        assert_eq!(
            crate::enqueue_dispatch(crate::DispatchOperation::EchoOwned(
                crate::ExcelValue::Integer(8)
            ))
            .unwrap_err(),
            crate::DispatchEnqueueError::NoActiveGeneration
        );
        assert!(backend.state.lock().unwrap().linked);
        assert_eq!(runtime.close(), Ok(LifecycleOutcome::Completed));
        assert!(!backend.state.lock().unwrap().linked);
        crate::dispatcher::reset_generations_for_test();
    }
    #[test]
    fn registration_release_failure_preserves_function_and_command_ids() {
        let _guard = runtime_test_guard();
        for free_index in [2, 4] {
            let backend = Arc::new(MockBackend::default());
            backend.state.lock().unwrap().fail_free = Some(free_index);
            let runtime = Runtime::with_backend(backend.clone());
            assert!(matches!(
                runtime.initialize(&MIXED_ADD_IN),
                Err(LifecycleError::RegistrationRollback {
                    primary: ExcelCallError::RegistrationRelease { .. },
                    ..
                })
            ));
            let state = backend.state.lock().unwrap();
            let expected = if free_index == 2 {
                vec![1.0]
            } else {
                vec![3.0, 2.0, 1.0]
            };
            assert_eq!(state.unregistered, expected);
            assert!(!state.linked);
        }
    }

    #[test]
    fn cleanup_retry_never_repeats_the_other_accepted_operation() {
        let _guard = runtime_test_guard();
        for target in [excel_api_sys::xlfSetName, excel_api_sys::xlfUnregister] {
            for tag in [
                excel_api_sys::xltypeBool,
                excel_api_sys::xltypeErr,
                excel_api_sys::xltypeNil,
            ] {
                let backend = Arc::new(MockBackend::default());
                let runtime = Runtime::with_backend(backend.clone());
                runtime.initialize(&ADD_IN).unwrap();
                backend.state.lock().unwrap().reject_cleanup = Some((target, tag));
                assert!(matches!(
                    runtime.close(),
                    Err(LifecycleError::CloseFailed { .. })
                ));
                assert!(matches!(
                    runtime.initialize(&ADD_IN),
                    Err(LifecycleError::Busy(RuntimePhase::CleanupRequired))
                ));
                let before = backend.state.lock().unwrap().calls.len();
                backend.state.lock().unwrap().reject_cleanup = None;
                runtime.close().unwrap();
                let state = backend.state.lock().unwrap();
                let other = if target == excel_api_sys::xlfSetName {
                    excel_api_sys::xlfUnregister
                } else {
                    excel_api_sys::xlfSetName
                };
                assert!(!state.calls[before..].iter().any(|(id, _)| *id == other));
                assert_eq!(
                    state.calls[before..]
                        .iter()
                        .filter(|(id, _)| *id == target)
                        .count(),
                    2
                );
            }
        }
    }

    #[test]
    fn accepted_cleanup_with_release_failure_is_not_retried() {
        let _guard = runtime_test_guard();
        for free_index in [4, 5] {
            let backend = Arc::new(MockBackend::default());
            let runtime = Runtime::with_backend(backend.clone());
            runtime.initialize(&ADD_IN).unwrap();
            backend.state.lock().unwrap().fail_free = Some(free_index);
            assert!(runtime.close().is_err());
            let before = backend.state.lock().unwrap().calls.len();
            runtime.close().unwrap();
            assert_eq!(backend.state.lock().unwrap().calls.len(), before);
        }
    }

    #[test]
    fn rollback_retains_only_failed_obligations() {
        let _guard = runtime_test_guard();
        let backend = Arc::new(MockBackend::failing_registration(2));
        backend.state.lock().unwrap().reject_cleanup =
            Some((excel_api_sys::xlfSetName, excel_api_sys::xltypeBool));
        let runtime = Runtime::with_backend(backend.clone());
        assert!(runtime.initialize(&ADD_IN).is_err());
        backend.state.lock().unwrap().reject_cleanup = None;
        runtime.close().unwrap();
        assert_eq!(backend.state.lock().unwrap().unregistered, [1.0]);
    }

    #[test]
    fn event_registration_preserves_nonzero_signed_ids_and_rejects_zero() {
        let backend = MockBackend::default();
        backend.link().unwrap();
        let capability = CallCapability::new(&backend);
        let context = LifecycleContext::new(&capability);
        for id in [1, i32::MAX, i32::MIN, -853147637] {
            backend.state.lock().unwrap().event_result = Some(id);
            assert_eq!(context.register_event("handler", 1).unwrap(), id);
        }
        backend.state.lock().unwrap().event_result = Some(0);
        assert!(context.register_event("handler", 1).is_err());
    }

    #[test]
    fn invalid_event_metadata_never_dispatches() {
        let backend = MockBackend::default();
        backend.link().unwrap();
        let capability = CallCapability::new(&backend);
        let context = LifecycleContext::new(&capability);
        for (name, event) in [
            ("handler", -1),
            ("", excel_api_sys::xleventCalculationEnded),
            ("bad\0name", excel_api_sys::xleventCalculationCanceled),
        ] {
            assert!(context.register_event(name, event).is_err());
        }
        assert!(backend.state.lock().unwrap().calls.is_empty());
    }
    #[test]
    fn abort_during_cleanup_defers_all_remaining_calls() {
        let _guard = runtime_test_guard();
        let backend = Arc::new(MockBackend::default());
        let runtime = Runtime::with_backend(backend.clone());
        runtime.initialize(&ADD_IN).unwrap();
        let before = backend.state.lock().unwrap().calls.len();
        backend.state.lock().unwrap().cleanup_code =
            Some((excel_api_sys::xlfSetName, excel_api_sys::xlretAbort));
        assert!(runtime.close().is_err());
        assert_eq!(backend.state.lock().unwrap().calls.len(), before + 1);
        backend.state.lock().unwrap().cleanup_code = None;
        runtime.close().unwrap();
        assert_eq!(backend.state.lock().unwrap().unregistered, [2.0, 1.0]);
    }

    #[test]
    fn cleanup_rejection_preserves_additional_release_error() {
        let _guard = runtime_test_guard();
        let backend = Arc::new(MockBackend::default());
        let runtime = Runtime::with_backend(backend.clone());
        runtime.initialize(&ADD_IN).unwrap();
        {
            let mut state = backend.state.lock().unwrap();
            state.reject_cleanup = Some((excel_api_sys::xlfSetName, excel_api_sys::xltypeErr));
            state.fail_free = Some(4);
        }
        let Err(LifecycleError::CloseFailed { errors }) = runtime.close() else {
            panic!("expected failure")
        };
        assert!(matches!(
            &errors[0],
            ExcelCallError::CleanupRejected {
                excel_error: Some(0),
                release: Some(_),
                ..
            }
        ));
        backend.state.lock().unwrap().reject_cleanup = None;
        runtime.close().unwrap();
    }

    #[test]
    fn release_abort_defers_rollback_and_preserves_the_accepted_id() {
        let _guard = runtime_test_guard();
        for code in [
            excel_api_sys::xlretAbort,
            excel_api_sys::xlretUncalced | excel_api_sys::xlretFailed,
        ] {
            let backend = Arc::new(MockBackend::default());
            {
                let mut state = backend.state.lock().unwrap();
                state.fail_free = Some(2);
                state.free_status = Some(code);
            }
            let runtime = Runtime::with_backend(backend.clone());
            assert!(matches!(
                runtime.initialize(&ADD_IN),
                Err(LifecycleError::RegistrationRollback {
                    primary: ExcelCallError::RegistrationRelease { .. },
                    ..
                })
            ));
            assert_eq!(runtime.phase(), RuntimePhase::CleanupRequired);
            assert_eq!(runtime.registration_ids(), [1.0]);
            {
                let mut state = backend.state.lock().unwrap();
                assert_eq!(state.calls.len(), 4);
                assert!(state.unregistered.is_empty());
                state.fail_free = None;
            }
            runtime.close().unwrap();
            assert_eq!(backend.state.lock().unwrap().unregistered, [1.0]);
        }
    }

    #[test]
    fn release_abort_after_cleanup_stops_and_does_not_repeat_accepted_side_effect() {
        let _guard = runtime_test_guard();
        for rejected in [false, true] {
            let backend = Arc::new(MockBackend::default());
            let runtime = Runtime::with_backend(backend.clone());
            runtime.initialize(&ADD_IN).unwrap();
            let before = {
                let mut state = backend.state.lock().unwrap();
                state.fail_free = Some(4);
                state.free_status = Some(excel_api_sys::xlretAbort | excel_api_sys::xlretFailed);
                if rejected {
                    state.reject_cleanup =
                        Some((excel_api_sys::xlfSetName, excel_api_sys::xltypeBool));
                }
                state.calls.len()
            };
            assert!(runtime.close().is_err());
            {
                let mut state = backend.state.lock().unwrap();
                assert_eq!(state.calls.len(), before + 2);
                assert!(state.unregistered.is_empty());
                state.fail_free = None;
                state.reject_cleanup = None;
            }
            runtime.close().unwrap();
            let state = backend.state.lock().unwrap();
            assert_eq!(state.unregistered, [2.0, 1.0]);
            let name_calls = state.calls[before..]
                .iter()
                .filter(|(f, _)| *f == excel_api_sys::xlfSetName)
                .count();
            assert_eq!(name_calls, if rejected { 3 } else { 2 });
        }
    }
}
