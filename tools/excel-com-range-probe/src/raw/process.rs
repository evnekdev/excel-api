//! Owned Excel process tracking without persisting process identities.

use windows_sys::Win32::Foundation::{CloseHandle, FILETIME, HANDLE, HWND, WAIT_OBJECT_0};
use windows_sys::Win32::System::Threading::{
    GetProcessTimes, OpenProcess, WaitForSingleObject, PROCESS_QUERY_LIMITED_INFORMATION,
    PROCESS_SYNCHRONIZE,
};
use windows_sys::Win32::UI::WindowsAndMessaging::GetWindowThreadProcessId;

use super::com_ptr::{ComPtr, Dispatch};
use super::dispatch::{call, Frame, DISPATCH_PROPERTYGET};

/// Owns only the synchronization handle for the Excel process created by this run.
pub(super) struct OwnedProcess(HANDLE);

impl OwnedProcess {
    pub(super) fn from_app(app: &ComPtr<Dispatch>, lcid: u32) -> Result<Self, String> {
        let hwnd_call = call(app, "Hwnd", DISPATCH_PROPERTYGET, lcid, Frame::empty());
        if hwnd_call.hr != 0 {
            return Err(format!("Hwnd failed: {}", super::hex(hwnd_call.hr)));
        }
        let hwnd = hwnd_call.result.i4_value().ok_or("Hwnd did not return VT_I4")?;
        let mut pid = 0;
        unsafe { GetWindowThreadProcessId(hwnd as isize as HWND, &mut pid) };
        if pid == 0 {
            return Err("Hwnd ownership lookup returned zero".to_owned());
        }
        let handle = unsafe {
            OpenProcess(
                PROCESS_QUERY_LIMITED_INFORMATION | PROCESS_SYNCHRONIZE,
                0,
                pid,
            )
        };
        if handle.is_null() {
            return Err("OpenProcess for owned Excel returned null".to_owned());
        }
        let mut creation = FILETIME::default();
        if unsafe {
            GetProcessTimes(
                handle,
                &mut creation,
                &mut FILETIME::default(),
                &mut FILETIME::default(),
                &mut FILETIME::default(),
            )
        } == 0
        {
            unsafe { CloseHandle(handle) };
            return Err("GetProcessTimes for owned Excel failed".to_owned());
        }
        Ok(Self(handle))
    }

    pub(super) fn wait(&self) -> bool {
        unsafe { WaitForSingleObject(self.0, 15_000) == WAIT_OBJECT_0 }
    }
}

impl Drop for OwnedProcess {
    fn drop(&mut self) {
        if !self.0.is_null() {
            unsafe { CloseHandle(self.0) };
        }
    }
}
