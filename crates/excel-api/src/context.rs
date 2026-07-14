use core::marker::PhantomData;

use crate::excel_call::CallCapability;

#[derive(Debug)]
pub struct WorksheetContext<'call> {
    #[allow(
        dead_code,
        reason = "worksheet-safe calls are added to this capability whitelist incrementally"
    )]
    capability: &'call CallCapability<'call>,
    _call: PhantomData<&'call mut &'call ()>,
}

#[derive(Debug)]
pub struct ThreadSafeContext<'call> {
    #[allow(
        dead_code,
        reason = "xlFree is consumed internally by Excel-owned result RAII"
    )]
    capability: &'call CallCapability<'call>,
    _call: PhantomData<&'call mut &'call ()>,
}

#[derive(Debug)]
pub struct LifecycleContext<'call> {
    capability: &'call CallCapability<'call>,
    _call: PhantomData<&'call mut &'call ()>,
}

#[derive(Debug)]
pub struct MacroContext<'call> {
    #[allow(dead_code, reason = "macro operations are outside the M8 catalogue")]
    capability: &'call CallCapability<'call>,
    _call: PhantomData<&'call mut &'call ()>,
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
context_impl!(MacroContext);

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
