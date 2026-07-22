use std::ptr::NonNull;
use windows_sys::Win32::System::Com::SAFEARRAY;
use windows_sys::Win32::System::Ole::SafeArrayDestroy;
/// Private SAFEARRAY owner for future array encode/decode work.
#[allow(dead_code)]
pub(crate) struct SafeArray(NonNull<SAFEARRAY>);
#[allow(dead_code)]
impl SafeArray {
    pub(crate) unsafe fn from_owned(raw: *mut SAFEARRAY) -> Option<Self> {
        NonNull::new(raw).map(Self)
    }
}
impl Drop for SafeArray {
    fn drop(&mut self) {
        // SAFETY: the constructor accepts only a SAFEARRAY ownership transfer and Drop runs once.
        unsafe {
            let _ = SafeArrayDestroy(self.0.as_ptr());
        }
    }
}
