//! STA apartment ownership for the research-only raw Automation kernel.

use windows_sys::Win32::System::Com::{CoInitializeEx, CoUninitialize, COINIT_APARTMENTTHREADED};

/// Balances a successful `CoInitializeEx(COINIT_APARTMENTTHREADED)` call.
pub(super) struct ComApartment;

impl ComApartment {
    pub(super) fn initialize() -> Result<Self, String> {
        let status = unsafe { CoInitializeEx(std::ptr::null(), COINIT_APARTMENTTHREADED as u32) };
        if status >= 0 {
            Ok(Self)
        } else {
            Err(format!("CoInitializeEx failed: {}", super::hex(status)))
        }
    }
}

impl Drop for ComApartment {
    fn drop(&mut self) {
        unsafe { CoUninitialize() }
    }
}
