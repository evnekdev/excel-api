use core::marker::PhantomData;

/// Token for an ordinary worksheet-function invocation.
#[derive(Debug)]
pub struct WorksheetContext<'call> {
    _call: PhantomData<&'call mut &'call ()>,
}

/// Token for a function registered as thread-safe.
///
/// This type intentionally exposes no general Excel C API access.
#[derive(Debug)]
pub struct ThreadSafeContext<'call> {
    _call: PhantomData<&'call mut &'call ()>,
}

/// Token for main-thread macro/command execution.
#[derive(Debug)]
pub struct MacroContext<'call> {
    _call: PhantomData<&'call mut &'call ()>,
}

impl WorksheetContext<'_> {
    /// Contexts may only be created by the runtime.
    pub(crate) const fn new() -> Self {
        Self { _call: PhantomData }
    }
}

impl ThreadSafeContext<'_> {
    /// Contexts may only be created by the runtime.
    pub(crate) const fn new() -> Self {
        Self { _call: PhantomData }
    }
}

impl MacroContext<'_> {
    /// Contexts may only be created by the runtime.
    pub(crate) const fn new() -> Self {
        Self { _call: PhantomData }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn context_tokens_are_zero_sized() {
        assert_eq!(core::mem::size_of::<WorksheetContext<'_>>(), 0);
        let _ = WorksheetContext::new();
        let _ = ThreadSafeContext::new();
        let _ = MacroContext::new();
    }
}
