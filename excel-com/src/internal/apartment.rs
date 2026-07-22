use std::marker::PhantomData;
use std::rc::Rc;
use std::thread::ThreadId;

use windows_sys::Win32::System::Com::{COINIT_APARTMENTTHREADED, CoInitializeEx, CoUninitialize};

use crate::ExcelComError;

/// Balances a caller-owned STA COM initialization on its creating thread.
#[derive(Debug)]
pub struct ComApartment {
    thread: ThreadId,
    _not_send_or_sync: PhantomData<Rc<()>>,
}

impl ComApartment {
    /// Initializes the current thread for STA COM Automation.
    pub fn sta() -> Result<Self, ExcelComError> {
        // SAFETY: null reserved pointer and documented STA flag are valid inputs.
        let status = unsafe { CoInitializeEx(std::ptr::null(), COINIT_APARTMENTTHREADED as u32) };
        if status >= 0 {
            Ok(Self {
                thread: std::thread::current().id(),
                _not_send_or_sync: PhantomData,
            })
        } else {
            Err(ExcelComError::Initialization { hresult: status })
        }
    }

    pub(crate) fn assert_current(&self) -> Result<(), ExcelComError> {
        (self.thread == std::thread::current().id())
            .then_some(())
            .ok_or(ExcelComError::Ownership {
                detail: "COM apartment used from a different thread",
            })
    }
}

impl Drop for ComApartment {
    fn drop(&mut self) {
        // SAFETY: this guard is constructed only after successful COM initialization.
        unsafe { CoUninitialize() }
    }
}
