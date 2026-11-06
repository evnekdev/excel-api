use core::marker::PhantomData;

use crate::excel_call::CallCapability;

/// Capability token for a genuine worksheet-function callback.
///
/// It borrows the callback capability and cannot be constructed by ordinary
/// safe code. It does not imply that every Excel call is legal.
///
/// ```compile_fail
/// fn send<T: Send>() {}
/// fn sync<T: Sync>() {}
/// send::<excel_api::WorksheetContext<'static>>();
/// sync::<excel_api::WorksheetContext<'static>>();
/// ```
#[derive(Debug)]
pub struct WorksheetContext<'call> {
    #[allow(
        dead_code,
        reason = "worksheet-safe calls are added to this capability whitelist incrementally"
    )]
    capability: &'call CallCapability<'call>,
    _call: PhantomData<(&'call mut &'call (), std::rc::Rc<()>)>,
}

/// Capability token for a genuine multi-thread-safe worksheet callback.
///
/// Only calls explicitly documented as thread-safe may use this context.
/// MTR permission does not permit moving the capability to a Rust-created thread.
///
/// ```compile_fail
/// fn send<T: Send>() {}
/// fn sync<T: Sync>() {}
/// send::<excel_api::ThreadSafeContext<'static>>();
/// sync::<excel_api::ThreadSafeContext<'static>>();
/// ```
#[derive(Debug)]
pub struct ThreadSafeContext<'call> {
    #[allow(
        dead_code,
        reason = "xlFree is consumed internally by Excel-owned result RAII"
    )]
    capability: &'call CallCapability<'call>,
    _call: PhantomData<(&'call mut &'call (), std::rc::Rc<()>)>,
}

#[derive(Debug)]
/// Lifecycle callbacks cannot poll user cancellation: `xlAbort` is exposed
/// only to the verified worksheet, thread-safe, and macro contexts.
///
/// ```compile_fail
/// fn send<T: Send>() {}
/// fn sync<T: Sync>() {}
/// send::<excel_api::LifecycleContext<'static>>();
/// sync::<excel_api::LifecycleContext<'static>>();
/// ```
///
/// ```compile_fail
/// use excel_api::LifecycleContext;
///
/// fn cannot_poll(context: &LifecycleContext<'_>) {
///     let _ = context.is_cancellation_requested();
/// }
/// ```
pub struct LifecycleContext<'call> {
    capability: &'call CallCapability<'call>,
    _call: PhantomData<(&'call mut &'call (), std::rc::Rc<()>)>,
}

/// Capability token for a macro-sheet-equivalent function or command callback.
///
/// It is callback-borrowed and must not be retained or sent to another thread.
/// Generated thunks retain the origin internally: macro-sheet worksheet functions
/// may query document information but cannot write cells or request recalculation.
///
/// ```compile_fail
/// fn send<T: Send>() {}
/// fn sync<T: Sync>() {}
/// send::<excel_api::MacroContext<'static>>();
/// sync::<excel_api::MacroContext<'static>>();
/// ```
#[derive(Debug)]
#[cfg_attr(
    not(feature = "xlcontime-research"),
    doc = r#"
The ordinary build exposes no xlcOnTime scheduling operation:

```compile_fail
use excel_api::MacroContext;

fn no_production_ontime(context: &MacroContext<'_>) {
    let _ = context.experimental_schedule_on_time(45_000.0, "COMMAND");
}
```
"#
)]
pub struct MacroContext<'call> {
    #[allow(dead_code, reason = "macro operations are outside the M8 catalogue")]
    capability: &'call CallCapability<'call>,
    command: bool,
    _call: PhantomData<(&'call mut &'call (), std::rc::Rc<()>)>,
}

macro_rules! context_impl {
    ($type:ident) => {
        impl<'call> $type<'call> {
            #[allow(
                dead_code,
                reason = "constructed as each callback class gains operations"
            )]
            pub(crate) const fn new(capability: &'call CallCapability<'call>) -> Self {
                Self {
                    capability,
                    _call: PhantomData,
                }
            }

            #[allow(dead_code, reason = "consumed as each callback class gains operations")]
            pub(crate) const fn capability(&self) -> &'call CallCapability<'call> {
                self.capability
            }
        }
    };
}

context_impl!(WorksheetContext);
context_impl!(ThreadSafeContext);
context_impl!(LifecycleContext);
impl<'call> MacroContext<'call> {
    pub(crate) const fn new(capability: &'call CallCapability<'call>) -> Self {
        Self {
            capability,
            command: true,
            _call: PhantomData,
        }
    }
    pub(crate) const fn macro_sheet(capability: &'call CallCapability<'call>) -> Self {
        Self {
            capability,
            command: false,
            _call: PhantomData,
        }
    }
    pub(crate) const fn capability(&self) -> &'call CallCapability<'call> {
        self.capability
    }
    pub(crate) fn require_command(
        &self,
        function: &'static str,
    ) -> Result<(), crate::ExcelCallError> {
        if self.command {
            Ok(())
        } else {
            Err(crate::ExcelCallError::IllegalContext {
                function,
                context: crate::CallPermission::Worksheet,
            })
        }
    }
}

impl ThreadSafeContext<'_> {
    /// Drains one bounded batch containing only thread-safe-compatible or pure work.
    pub fn drain_dispatcher(&self) -> crate::DispatchDrainReport {
        crate::dispatcher::drain_thread_safe(self)
    }
}

impl WorksheetContext<'_> {
    /// Drains one bounded batch containing only worksheet-compatible or pure work.
    pub fn drain_dispatcher(&self) -> crate::DispatchDrainReport {
        crate::dispatcher::drain_worksheet(self)
    }
}

impl MacroContext<'_> {
    /// Drains one bounded cooperative batch under this genuine macro callback.
    ///
    /// Enqueueing a request does not invoke this method or otherwise wake Excel.
    pub fn drain_dispatcher(&self) -> crate::DispatchDrainReport {
        crate::dispatcher::drain_macro(self)
    }
}

impl LifecycleContext<'_> {
    /// Drains one bounded batch containing only lifecycle-compatible or pure work.
    /// Runtime close does not call this; close retires queued work instead.
    pub fn drain_dispatcher(&self) -> crate::DispatchDrainReport {
        crate::dispatcher::drain_lifecycle(self)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::excel_call::test_support::UnavailableBackend;

    #[test]
    fn contexts_carry_a_real_borrowed_capability() {
        let backend = UnavailableBackend;
        let capability = CallCapability::new(&backend);
        let worksheet = WorksheetContext::new(&capability);
        let thread_safe = ThreadSafeContext::new(&capability);
        let lifecycle = LifecycleContext::new(&capability);
        assert!(core::ptr::eq(worksheet.capability(), &capability));
        assert!(core::ptr::eq(thread_safe.capability(), &capability));
        assert!(core::ptr::eq(lifecycle.capability(), &capability));
    }
}
