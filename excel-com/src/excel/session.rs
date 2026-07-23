//! Explicit ownership, attachment, and diagnostics for Excel applications.

use std::ffi::c_void;
use std::marker::PhantomData;
use std::ops::Deref;
use std::rc::Rc;
use std::time::Duration;

use windows_sys::Win32::Foundation::{CloseHandle, WAIT_OBJECT_0};
use windows_sys::Win32::System::Com::CLSIDFromProgID;
use windows_sys::Win32::System::Ole::GetActiveObject;
use windows_sys::Win32::System::Threading::{
    OpenProcess, PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_SYNCHRONIZE, WaitForSingleObject,
};
use windows_sys::Win32::UI::WindowsAndMessaging::GetWindowThreadProcessId;
use windows_sys::core::GUID;

use super::{Application, CalculationState};
use crate::automation::{ComMessageFilterGuard, ComRetryPolicy};
use crate::internal::{ComPtr, Dispatch, Unknown, wide_nul};
use crate::{ComApartment, ExcelComError, ExcelRuntimeError};

const IID_IDISPATCH: GUID = GUID::from_u128(0x00020400_0000_0000_c000_000000000046);
const MK_E_UNAVAILABLE: i32 = 0x8004_01E3_u32 as i32;

/// Distinguishes a process created by this crate from a shared attachment.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum SessionOwnership {
    /// This crate activated the Excel local server.
    Owned,
    /// This crate obtained a shared reference from Excel's active-object registration.
    Attached,
}

/// Office process architecture when it is safely observable.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum OfficeBitness {
    /// 32-bit Office.
    X86,
    /// 64-bit x86 Office.
    X64,
    /// ARM64 Office.
    Arm64,
}

/// A bounded request for attachment to an existing Excel Automation server.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct AttachOptions {
    /// Selection supported by the active-object registration mechanism.
    pub selection: ExistingInstanceSelection,
}

impl Default for AttachOptions {
    fn default() -> Self {
        Self {
            selection: ExistingInstanceSelection::ActiveInstance,
        }
    }
}

/// Selection modes that are reliable for Excel's `GetActiveObject` registration.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum ExistingInstanceSelection {
    /// Attach to the active Excel Automation registration.
    FirstAvailable,
    /// Attach to the active Excel Automation registration.
    ActiveInstance,
}

/// A non-owning view of the Excel session held by an owned or attached handle.
pub enum ExcelSession<'apartment> {
    /// A crate-created Excel server that can be explicitly shut down.
    Owned(OwnedApplication<'apartment>),
    /// A user or other-process Excel server that can only be observed or used.
    Attached(AttachedApplication<'apartment>),
}

impl ExcelSession<'_> {
    /// Returns the shared, apartment-bound Application surface.
    pub fn application(&self) -> &Application {
        match self {
            Self::Owned(application) => application.application(),
            Self::Attached(application) => application.application(),
        }
    }

    /// Captures the session diagnostics without exposing a process handle.
    pub fn diagnostics(&self) -> Result<ExcelSessionDiagnostics, ExcelComError> {
        match self {
            Self::Owned(application) => application.diagnostics(),
            Self::Attached(application) => application.diagnostics(),
        }
    }
}

/// A crate-created Excel session. Only this type can request Excel shutdown.
///
/// The type is both apartment-bound and deliberately not thread-safe:
///
/// ```compile_fail
/// fn require_send<T: Send>() {}
/// require_send::<excel_com::OwnedApplication<'static>>();
/// ```
///
/// It also cannot outlive its COM apartment:
///
/// ```compile_fail
/// fn invalid() -> excel_com::OwnedApplication<'static> {
///     let apartment = excel_com::ComApartment::sta().unwrap();
///     excel_com::OwnedApplication::new(&apartment).unwrap()
/// }
/// ```
pub struct OwnedApplication<'apartment> {
    inner: Application,
    process: OwnedExcelProcess,
    message_filter: Option<ComMessageFilterGuard>,
    _apartment: PhantomData<&'apartment ComApartment>,
    _not_send_or_sync: PhantomData<Rc<()>>,
}

impl std::fmt::Debug for OwnedApplication<'_> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("OwnedApplication")
            .field("application", &self.inner)
            .field("process_id", &self.process.process_id)
            .field("message_filter", &self.message_filter.is_some())
            .finish()
    }
}

impl<'apartment> OwnedApplication<'apartment> {
    /// Activates a fresh local Excel server without a retry policy.
    pub fn new(apartment: &'apartment ComApartment) -> Result<Self, ExcelComError> {
        Self::new_inner(apartment, None)
    }

    /// Activates a fresh local Excel server and installs an STA-local policy
    /// for safe COM retry operations.
    pub fn new_with_retry_policy(
        apartment: &'apartment ComApartment,
        policy: ComRetryPolicy,
    ) -> Result<Self, ExcelComError> {
        let filter = ComMessageFilterGuard::install(apartment, policy)?;
        Self::new_inner(apartment, Some(filter))
    }

    fn new_inner(
        apartment: &'apartment ComApartment,
        message_filter: Option<ComMessageFilterGuard>,
    ) -> Result<Self, ExcelComError> {
        apartment.assert_current()?;
        let inner = Application::from_dispatch(crate::automation::activate_excel()?);
        let process = OwnedExcelProcess::observe(&inner);
        Ok(Self {
            inner,
            process,
            message_filter,
            _apartment: PhantomData,
            _not_send_or_sync: PhantomData,
        })
    }

    /// Returns the shared Excel object-model wrapper without transferring quit rights.
    pub fn application(&self) -> &Application {
        &self.inner
    }

    /// Captures process and Excel-state diagnostics for this owned session.
    pub fn diagnostics(&self) -> Result<ExcelSessionDiagnostics, ExcelComError> {
        collect_diagnostics(
            &self.inner,
            SessionOwnership::Owned,
            self.process.process_id,
        )
    }

    /// Waits for the exact owned process observed from this server's own window.
    ///
    /// This never searches by process name and never terminates a process. A
    /// caller normally invokes [`Self::quit`] or [`Self::quit_and_wait`] first.
    pub fn wait_for_exit(&self, timeout: Duration) -> Result<ProcessExitReport, ExcelComError> {
        self.process.wait_for_exit(timeout)
    }

    /// Requests normal Excel shutdown and releases this owned wrapper.
    ///
    /// `Drop` never calls `Quit`. Other child wrappers should be released before
    /// calling this method so Excel can naturally terminate.
    pub fn quit(self) -> Result<(), ExcelComError> {
        self.inner.request_quit()
    }

    /// Requests normal Excel shutdown, releases COM references, then waits for
    /// the exact owned process to exit naturally.
    pub fn quit_and_wait(self, timeout: Duration) -> Result<ProcessExitReport, ExcelComError> {
        self.inner.request_quit()?;
        let process = self.process;
        drop(self);
        process.wait_for_exit(timeout)
    }
}

impl Deref for OwnedApplication<'_> {
    type Target = Application;

    fn deref(&self) -> &Self::Target {
        &self.inner
    }
}

/// A shared reference to an existing Excel server. It never owns shutdown.
///
/// Attached sessions are apartment-bound, not transferable between threads,
/// and intentionally have no `quit` method:
///
/// ```compile_fail
/// fn require_sync<T: Sync>() {}
/// require_sync::<excel_com::AttachedApplication<'static>>();
/// ```
///
/// ```compile_fail
/// fn must_not_quit(session: excel_com::AttachedApplication<'_>) {
///     session.quit();
/// }
/// ```
pub struct AttachedApplication<'apartment> {
    inner: Application,
    _apartment: PhantomData<&'apartment ComApartment>,
    _not_send_or_sync: PhantomData<Rc<()>>,
}

impl std::fmt::Debug for AttachedApplication<'_> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_tuple("AttachedApplication")
            .field(&self.inner)
            .finish()
    }
}

impl<'apartment> AttachedApplication<'apartment> {
    /// Attaches through Excel's active-object registration on `apartment`'s STA.
    ///
    /// Excel does not expose a reliable process-id selection through this
    /// registration. The API intentionally omits a `ProcessId` selector rather
    /// than guessing from process names or arbitrary window enumeration.
    pub fn attach(
        apartment: &'apartment ComApartment,
        _options: &AttachOptions,
    ) -> Result<Self, ExcelComError> {
        apartment.assert_current()?;
        let name = wide_nul("Excel.Application");
        let mut class = GUID::default();
        // SAFETY: valid ProgID and output storage.
        let status = unsafe { CLSIDFromProgID(name.as_ptr(), &mut class) };
        if ExcelComError::failed(status) {
            return Err(ExcelComError::Runtime(ExcelRuntimeError::RotAccessFailed {
                hresult: status,
            }));
        }
        let mut raw: *mut c_void = std::ptr::null_mut();
        // SAFETY: class and output storage remain valid through the lookup.
        let status = unsafe { GetActiveObject(&class, std::ptr::null_mut(), &mut raw) };
        if status == MK_E_UNAVAILABLE {
            return Err(ExcelComError::Runtime(ExcelRuntimeError::NoRunningInstance));
        }
        if ExcelComError::failed(status) {
            return Err(ExcelComError::Runtime(ExcelRuntimeError::RotAccessFailed {
                hresult: status,
            }));
        }
        // SAFETY: GetActiveObject returned one IUnknown reference on success.
        let unknown: ComPtr<Unknown> = unsafe { ComPtr::from_owned(raw) }?;
        let dispatch = unknown.query_interface::<Dispatch>(&IID_IDISPATCH)?;
        Ok(Self {
            inner: Application::from_dispatch(dispatch),
            _apartment: PhantomData,
            _not_send_or_sync: PhantomData,
        })
    }

    /// Returns the shared, non-owning Application surface.
    pub fn application(&self) -> &Application {
        &self.inner
    }

    /// Captures Excel-state diagnostics without assuming process ownership.
    pub fn diagnostics(&self) -> Result<ExcelSessionDiagnostics, ExcelComError> {
        collect_diagnostics(&self.inner, SessionOwnership::Attached, None)
    }
}

impl Deref for AttachedApplication<'_> {
    type Target = Application;

    fn deref(&self) -> &Self::Target {
        &self.inner
    }
}

/// Sanitized process and object-model observations for an Excel session.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct ExcelSessionDiagnostics {
    /// Whether this crate owns the Excel server.
    pub ownership: SessionOwnership,
    /// Process id observed from the owned server window, when available.
    pub process_id: Option<u32>,
    /// Excel's reported application version.
    pub version: String,
    /// Office architecture, when safely observed.
    pub bitness: Option<OfficeBitness>,
    /// Count of workbooks currently open in this Application.
    pub workbook_count: usize,
    /// Count of Excel windows currently open in this Application.
    pub window_count: usize,
    /// Whether Excel reports its main UI as visible.
    pub visible: bool,
    /// Excel `Ready`, when exposed by this host state.
    pub ready: Option<bool>,
    /// Excel `Interactive`, when exposed by this host state.
    pub interactive: Option<bool>,
    /// Excel calculation state, when exposed by this host state.
    pub calculation_state: Option<CalculationState>,
}

/// Result of waiting for an exact owned Excel process.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct ProcessExitReport {
    /// Process id observed from the owned Excel window.
    pub process_id: u32,
    /// Whether the process exited before the requested timeout.
    pub exited: bool,
    /// Duration requested by the caller.
    pub timeout: Duration,
}

#[derive(Clone, Copy)]
struct OwnedExcelProcess {
    process_id: Option<u32>,
}

impl OwnedExcelProcess {
    fn observe(application: &Application) -> Self {
        let process_id = application.hwnd().ok().and_then(|raw| {
            let mut process_id = 0;
            // SAFETY: Application.Hwnd is an Excel-owned window handle and the
            // output address is valid for this direct diagnostic query.
            let thread = unsafe { GetWindowThreadProcessId(raw as isize as _, &mut process_id) };
            (thread != 0 && process_id != 0).then_some(process_id)
        });
        Self { process_id }
    }

    fn wait_for_exit(&self, timeout: Duration) -> Result<ProcessExitReport, ExcelComError> {
        let process_id = self.process_id.ok_or(ExcelComError::Unsupported {
            detail: "owned Excel process id was not observable from Application.Hwnd",
        })?;
        // SAFETY: access rights are query/synchronize only; no termination right
        // is requested and the returned handle is closed before this method exits.
        let handle = unsafe {
            OpenProcess(
                PROCESS_QUERY_LIMITED_INFORMATION | PROCESS_SYNCHRONIZE,
                0,
                process_id,
            )
        };
        if handle.is_null() {
            return Err(ExcelComError::Runtime(
                ExcelRuntimeError::SessionDisappeared,
            ));
        }
        let milliseconds = timeout.as_millis().min(u128::from(u32::MAX)) as u32;
        // SAFETY: `handle` is a valid process handle and the bounded timeout is valid.
        let status = unsafe { WaitForSingleObject(handle, milliseconds) };
        // SAFETY: this method owns the handle returned by OpenProcess.
        unsafe { CloseHandle(handle) };
        if status == WAIT_OBJECT_0 {
            return Ok(ProcessExitReport {
                process_id,
                exited: true,
                timeout,
            });
        }
        Err(ExcelComError::Runtime(
            ExcelRuntimeError::ProcessExitTimeout {
                process_id: Some(process_id),
                timeout,
            },
        ))
    }
}

fn collect_diagnostics(
    application: &Application,
    ownership: SessionOwnership,
    process_id: Option<u32>,
) -> Result<ExcelSessionDiagnostics, ExcelComError> {
    Ok(ExcelSessionDiagnostics {
        ownership,
        process_id,
        version: application.version()?,
        // The public Automation surface does not provide an authoritative
        // bitness value. Returning None is safer than guessing from process
        // names or the caller architecture.
        bitness: None,
        workbook_count: application.workbook_count()?,
        window_count: application.window_count()?,
        visible: application.visible()?,
        ready: application.ready().ok(),
        interactive: application.interactive().ok(),
        calculation_state: application.calculation_state().ok(),
    })
}
