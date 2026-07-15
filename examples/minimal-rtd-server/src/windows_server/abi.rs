//! Raw RTD ABI verified from the installed Excel type library 1.9.
//!
//! The type library exposes dual interfaces. These definitions use the hidden
//! `TKIND_INTERFACE` vtables, not the Automation dispatch projection.

use core::ffi::c_void;
use windows::Win32::Foundation::VARIANT_BOOL;
use windows::Win32::System::Com::{IDispatch, IDispatch_Vtbl, SAFEARRAY};
use windows::Win32::System::Variant::VARIANT;
use windows::core::{GUID, HRESULT, Interface, Result};

pub(crate) const CLSID_MINIMAL_RTD: GUID = GUID::from_u128(0xdc738fe5_30ee_40e8_a8c2_3d16f217c52d);
pub(crate) const IID_IRTD_SERVER: GUID = GUID::from_u128(0xec0e6191_db51_11d3_8f3e_00c04f3651b8);
pub(crate) const IID_IRTD_UPDATE_EVENT: GUID =
    GUID::from_u128(0xa43788c1_d91b_11d3_8f39_00c04f3651b8);
pub(crate) const EXCEL_TYPELIB: GUID = GUID::from_u128(0x00020813_0000_0000_c000_000000000046);
pub(crate) const EXCEL_TYPELIB_MAJOR: u16 = 1;
pub(crate) const EXCEL_TYPELIB_MINOR: u16 = 9;

windows_core::imp::define_interface!(
    IRtdUpdateEvent,
    IRtdUpdateEvent_Vtbl,
    0xa43788c1_d91b_11d3_8f39_00c04f3651b8
);
windows_core::imp::interface_hierarchy!(IRtdUpdateEvent, IDispatch, windows_core::IUnknown);

#[repr(C)]
pub struct IRtdUpdateEvent_Vtbl {
    pub base__: IDispatch_Vtbl,
    pub update_notify: unsafe extern "system" fn(*mut c_void) -> HRESULT,
    pub heartbeat_interval_get: unsafe extern "system" fn(*mut c_void, *mut i32) -> HRESULT,
    pub heartbeat_interval_put: unsafe extern "system" fn(*mut c_void, i32) -> HRESULT,
    pub disconnect: unsafe extern "system" fn(*mut c_void) -> HRESULT,
}

impl IRtdUpdateEvent {
    pub(crate) unsafe fn update_notify(&self) -> Result<()> {
        // SAFETY: `self` is an apartment-valid COM proxy and the slot is fixed
        // by Excel type library 1.9's hidden dual-interface vtable.
        unsafe { (Interface::vtable(self).update_notify)(Interface::as_raw(self)).ok() }
    }
}

#[repr(C)]
pub(crate) struct IRtdServer_Vtbl {
    pub base__: IDispatch_Vtbl,
    pub server_start: unsafe extern "system" fn(*mut c_void, *mut c_void, *mut i32) -> HRESULT,
    pub connect_data: unsafe extern "system" fn(
        *mut c_void,
        i32,
        *mut *mut SAFEARRAY,
        *mut VARIANT_BOOL,
        *mut VARIANT,
    ) -> HRESULT,
    pub refresh_data:
        unsafe extern "system" fn(*mut c_void, *mut i32, *mut *mut SAFEARRAY) -> HRESULT,
    pub disconnect_data: unsafe extern "system" fn(*mut c_void, i32) -> HRESULT,
    pub heartbeat: unsafe extern "system" fn(*mut c_void, *mut i32) -> HRESULT,
    pub server_terminate: unsafe extern "system" fn(*mut c_void) -> HRESULT,
}

const _: () = {
    assert!(size_of::<GUID>() == 16);
    assert!(size_of::<HRESULT>() == 4);
    assert!(size_of::<i32>() == 4);
    assert!(size_of::<VARIANT_BOOL>() == 2);
    assert!(size_of::<IRtdServer_Vtbl>() == size_of::<usize>() * 13);
    assert!(size_of::<IRtdUpdateEvent_Vtbl>() == size_of::<usize>() * 11);
};

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn installed_typelib_iids_and_vtable_sizes_are_pinned() {
        assert_eq!(
            IID_IRTD_SERVER,
            GUID::from_u128(0xec0e6191_db51_11d3_8f3e_00c04f3651b8)
        );
        assert_eq!(
            IID_IRTD_UPDATE_EVENT,
            GUID::from_u128(0xa43788c1_d91b_11d3_8f39_00c04f3651b8)
        );
        assert_eq!(size_of::<IRtdServer_Vtbl>(), size_of::<usize>() * 13);
        assert_eq!(size_of::<IRtdUpdateEvent_Vtbl>(), size_of::<usize>() * 11);
    }
}
