//! Idempotent XLL lifecycle runtime.

use core::fmt;
use std::sync::{Arc, Mutex, OnceLock};

use crate::excel_call::{CallCapability, ExcelCallBackend, ExcelCallError, SdkExcel12vBackend};
use crate::{AddInDescriptor, ExcelString, LifecycleContext};

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
}

impl Default for RuntimeState {
    fn default() -> Self {
        Self {
            phase: RuntimePhase::Uninitialized,
            module: None,
            worksheet_registrations: Vec::new(),
            command_registrations: Vec::new(),
            diagnostics: RuntimeDiagnostics::default(),
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

    pub fn initialize(&self, add_in: &AddInDescriptor) -> Result<LifecycleOutcome, LifecycleError> {
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

        if let Err(error) = self.backend.link() {
            self.lock().phase = RuntimePhase::Uninitialized;
            return Err(error.into());
        }

        let capability = CallCapability::new(self.backend.as_ref());
        let context = LifecycleContext::new(&capability);
        let module = match context.get_module_name() {
            Ok(module) => module,
            Err(error) => {
                self.backend.unlink();
                self.lock().phase = RuntimePhase::Uninitialized;
                return Err(error.into());
            }
        };

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
                    let mut state = self.lock();
                    state.diagnostics.rollback_attempts += registered.len();
                    if cleanup.is_empty() {
                        self.backend.unlink();
                        state.phase = RuntimePhase::Uninitialized;
                        state.module = None;
                        state.worksheet_registrations.clear();
                        state.command_registrations.clear();
                    } else {
                        // Keep the backend and conservative cleanup metadata
                        // alive; a later close can retry anything Excel may
                        // still consider registered.
                        state.phase = RuntimePhase::Initialized;
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
                    let mut state = self.lock();
                    state.diagnostics.rollback_attempts +=
                        command_registered.len() + registered.len();
                    if cleanup.is_empty() {
                        self.backend.unlink();
                        state.phase = RuntimePhase::Uninitialized;
                        state.module = None;
                        state.worksheet_registrations.clear();
                        state.command_registrations.clear();
                    } else {
                        state.phase = RuntimePhase::Initialized;
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

    pub fn close(&self) -> Result<LifecycleOutcome, LifecycleError> {
        let registrations = {
            let mut state = self.lock();
            state.diagnostics.close_attempts += 1;
            match state.phase {
                RuntimePhase::Uninitialized => return Ok(LifecycleOutcome::AlreadyInState),
                RuntimePhase::Initialized => {
                    state.phase = RuntimePhase::Closing;
                    let mut entries = state.command_registrations.clone();
                    entries.extend(state.worksheet_registrations.iter().copied());
                    entries
                }
                phase => return Err(LifecycleError::Busy(phase)),
            }
        };

        let capability = CallCapability::new(self.backend.as_ref());
        let context = LifecycleContext::new(&capability);
        let mut failed = Vec::new();
        let mut errors = Vec::new();
        for RegisteredEntry { id, name, kind } in registrations.iter().rev().copied() {
            let result = context
                .delete_defined_name(name)
                .and_then(|()| context.unregister(id));
            if let Err(error) = result {
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
            state.phase = RuntimePhase::Initialized;
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
        register_count: usize,
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

    #[test]
    fn duplicate_sequences_register_once_and_close_is_idempotent() {
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
        assert!(state.allocations.is_empty());
        assert_eq!(state.free_calls, 7); // xlGetName, registrations, name deletion, unregisters.
    }

    #[test]
    fn partial_registration_rolls_back_and_retry_succeeds() {
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
        let runtime = Runtime::with_backend(Arc::new(
            crate::excel_call::test_support::UnavailableBackend,
        ));
        assert_eq!(
            runtime.initialize(&ADD_IN),
            Err(LifecycleError::Excel(ExcelCallError::BackendUnavailable))
        );
        assert_eq!(runtime.phase(), RuntimePhase::Uninitialized);
    }
}
