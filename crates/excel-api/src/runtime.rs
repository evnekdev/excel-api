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
}

#[derive(Clone, Copy, Eq, PartialEq)]
enum RegistrationKind {
    Worksheet,
    Command,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum RuntimePhase {
    Uninitialized,
    Initializing,
    Initialized,
    Closing,
    /// Excel registrations remain and close must be retried. Async scheduling
    /// is disabled and the callback backend remains linked for cleanup only.
    CleanupRequired,
}

#[derive(Clone, Debug, Default, Eq, PartialEq)]
pub struct RuntimeDiagnostics {
    pub initialization_attempts: usize,
    pub successful_initializations: usize,
    pub duplicate_initializations: usize,
    pub close_attempts: usize,
    pub successful_closes: usize,
    pub registrations: usize,
    pub unregister_attempts: usize,
    pub rollback_attempts: usize,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum LifecycleOutcome {
    Completed,
    AlreadyInState,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub enum LifecycleError {
    Busy(RuntimePhase),
    Async(AsyncSubmitError),
    Excel(ExcelCallError),
    RegistrationRollback {
        primary: ExcelCallError,
        cleanup: Vec<ExcelCallError>,
    },
    CloseFailed {
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

    pub fn phase(&self) -> RuntimePhase {
        self.lock().phase
    }
    pub fn diagnostics(&self) -> RuntimeDiagnostics {
        self.lock().diagnostics.clone()
    }
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

    pub fn initialize(&self, add_in: &AddInDescriptor) -> Result<LifecycleOutcome, LifecycleError> {
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

        let mut registered = Vec::with_capacity(add_in.functions.len());
        for function in add_in.functions {
            match context.register(&module, function) {
                Ok(id) => registered.push(id),
                Err(primary) => {
                    let mut cleanup = Vec::new();
                    for (index, id) in registered.iter().copied().enumerate().rev() {
                        if let Err(error) =
                            context.delete_defined_name(add_in.functions[index].excel_name)
                        {
                            cleanup.push(error);
                        }
                        if let Err(error) = context.unregister(id) {
                            cleanup.push(error);
                        }
                    }
                    if has_async {
                        crate::async_udf::shutdown();
                    }
                    if dispatcher_active {
                        crate::dispatcher::shutdown();
                    }
                    if cleanup.is_empty() {
                        self.backend.unlink();
                    }
                    let mut state = self.lock();
                    state.diagnostics.rollback_attempts += registered.len();
                    state.async_generation_active = false;
                    state.dispatcher_generation_active = false;
                    if cleanup.is_empty() {
                        state.phase = RuntimePhase::Uninitialized;
                        state.module = None;
                        state.worksheet_registrations.clear();
                        state.command_registrations.clear();
                    } else {
                        // Keep the backend and conservative cleanup metadata
                        // alive; a later close can retry anything Excel may
                        // still consider registered.
                        state.phase = RuntimePhase::CleanupRequired;
                        state.module = Some(module);
                        state.worksheet_registrations = add_in.functions[..registered.len()]
                            .iter()
                            .zip(registered.iter().copied())
                            .map(|(function, id)| RegisteredEntry {
                                id,
                                name: function.excel_name,
                                kind: RegistrationKind::Worksheet,
                            })
                            .collect();
                        state.command_registrations.clear();
                    }
                    return Err(LifecycleError::RegistrationRollback { primary, cleanup });
                }
            }
        }

        let mut command_registered = Vec::with_capacity(add_in.commands.len());
        for command in add_in.commands {
            match context.register_command(&module, command) {
                Ok(id) => command_registered.push(id),
                Err(primary) => {
                    let mut cleanup = Vec::new();
                    for (index, id) in command_registered.iter().copied().enumerate().rev() {
                        if let Err(error) =
                            context.delete_defined_name(add_in.commands[index].excel_name)
                        {
                            cleanup.push(error);
                        }
                        if let Err(error) = context.unregister(id) {
                            cleanup.push(error);
                        }
                    }
                    for (index, id) in registered.iter().copied().enumerate().rev() {
                        if let Err(error) =
                            context.delete_defined_name(add_in.functions[index].excel_name)
                        {
                            cleanup.push(error);
                        }
                        if let Err(error) = context.unregister(id) {
                            cleanup.push(error);
                        }
                    }
                    if has_async {
                        crate::async_udf::shutdown();
                    }
                    if dispatcher_active {
                        crate::dispatcher::shutdown();
                    }
                    if cleanup.is_empty() {
                        self.backend.unlink();
                    }
                    let mut state = self.lock();
                    state.diagnostics.rollback_attempts +=
                        command_registered.len() + registered.len();
                    state.async_generation_active = false;
                    state.dispatcher_generation_active = false;
                    if cleanup.is_empty() {
                        state.phase = RuntimePhase::Uninitialized;
                        state.module = None;
                        state.worksheet_registrations.clear();
                        state.command_registrations.clear();
                    } else {
                        state.phase = RuntimePhase::CleanupRequired;
                        state.module = Some(module);
                        state.worksheet_registrations = registered
                            .iter()
                            .copied()
                            .zip(add_in.functions.iter())
                            .map(|(id, function)| RegisteredEntry {
                                id,
                                name: function.excel_name,
                                kind: RegistrationKind::Worksheet,
                            })
                            .collect();
                        state.command_registrations = command_registered
                            .iter()
                            .copied()
                            .zip(add_in.commands.iter())
                            .map(|(id, command)| RegisteredEntry {
                                id,
                                name: command.excel_name,
                                kind: RegistrationKind::Command,
                            })
                            .collect();
                    }
                    return Err(LifecycleError::RegistrationRollback { primary, cleanup });
                }
            }
        }

        let mut state = self.lock();
        state.phase = RuntimePhase::Initialized;
        state.module = Some(module);
        state.diagnostics.registrations += registered.len() + command_registered.len();
        state.worksheet_registrations = registered
            .iter()
            .copied()
            .zip(add_in.functions.iter())
            .map(|(id, function)| RegisteredEntry {
                id,
                name: function.excel_name,
                kind: RegistrationKind::Worksheet,
            })
            .collect();
        state.command_registrations = command_registered
            .iter()
            .copied()
            .zip(add_in.commands.iter())
            .map(|(id, command)| RegisteredEntry {
                id,
                name: command.excel_name,
                kind: RegistrationKind::Command,
            })
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
        let mut failed = Vec::new();
        let mut errors = Vec::new();
        for RegisteredEntry { id, name, kind } in registrations.iter().rev().copied() {
            if let Err(error) = context.delete_defined_name(name) {
                errors.push(error);
            }
            if let Err(error) = context.unregister(id) {
                failed.push(RegisteredEntry { id, name, kind });
                errors.push(error);
            }
        }
        let mut state = self.lock();
        state.diagnostics.unregister_attempts += registrations.len();
        if errors.is_empty() {
            self.backend.unlink();
            state.phase = RuntimePhase::Uninitialized;
            state.module = None;
            state.worksheet_registrations.clear();
            state.command_registrations.clear();
            state.diagnostics.successful_closes += 1;
            Ok(LifecycleOutcome::Completed)
        } else {
            failed.reverse();
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
        unregistered: Vec<f64>,
        allocations: HashMap<usize, Box<[u16]>>,
        next_id: f64,
        fail_registration: Option<usize>,
        fail_unregistration: usize,
        fail_event: Option<usize>,
        register_count: usize,
        event_count: usize,
        event_ids: Vec<i32>,
        unregister_linked: Vec<bool>,
        free_calls: usize,
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
                            w: state.event_count as i32,
                        },
                        xltype: excel_api_sys::xltypeInt,
                    };
                }
            } else if function == excel_api_sys::xlFree {
                state.free_calls += 1;
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

    fn runtime_test_guard() -> std::sync::MutexGuard<'static, ()> {
        static GUARD: OnceLock<Mutex<()>> = OnceLock::new();
        GUARD
            .get_or_init(|| Mutex::new(()))
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner)
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
            [
                excel_api_sys::xleventCalculationCanceled,
                excel_api_sys::xleventCalculationEnded,
            ]
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
        let _guard = crate::dispatcher::TEST_SERIAL.lock().unwrap();
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
}
