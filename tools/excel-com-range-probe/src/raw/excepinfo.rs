//! EXCEPINFO cleanup and normalized copied diagnostics.

use serde_json::{json, Value};
use windows_sys::Win32::Foundation::SysFreeString;
use windows_sys::Win32::System::Com::EXCEPINFO;

/// Owns returned EXCEPINFO BSTRs and invokes deferred fill-in before copying.
pub(super) struct OwnedExcepInfo(pub(super) EXCEPINFO);

impl OwnedExcepInfo {
    pub(super) fn new() -> Self {
        Self(EXCEPINFO::default())
    }

    pub(super) fn take(&mut self) -> Value {
        let deferred = self.0.pfnDeferredFillIn.is_some();
        let deferred_status = self.0.pfnDeferredFillIn.map(|fill| unsafe { fill(&mut self.0) });
        let scode = self.0.scode;
        unsafe {
            for value in [
                &mut self.0.bstrSource,
                &mut self.0.bstrDescription,
                &mut self.0.bstrHelpFile,
            ] {
                if !(*value).is_null() {
                    SysFreeString(*value);
                    *value = std::ptr::null();
                }
            }
        }
        json!({
            "deferred_fill_in_present": deferred,
            "deferred_fill_in_hresult": deferred_status.map(super::hex),
            "scode": super::hex(scode),
            "wcode": self.0.wCode,
        })
    }
}

impl Drop for OwnedExcepInfo {
    fn drop(&mut self) {
        let _ = self.take();
    }
}
