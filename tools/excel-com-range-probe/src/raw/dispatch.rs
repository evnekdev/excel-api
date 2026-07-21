//! Generic SDK `IDispatch` ABI and invocation frame ownership.

use std::ffi::c_void;

use serde_json::Value;
use windows_sys::core::{IUnknown_Vtbl, GUID, HRESULT};
use windows_sys::Win32::System::Com::{DISPPARAMS, EXCEPINFO};
use windows_sys::Win32::System::Variant::VARIANT;

use super::com_ptr::{ComPtr, Dispatch};
use super::excepinfo::OwnedExcepInfo;
use super::variant::OwnedVariant;

pub(super) const DISPATCH_METHOD: u16 = 1;
pub(super) const DISPATCH_PROPERTYGET: u16 = 2;
pub(super) const DISPATCH_PROPERTYPUT: u16 = 4;
const DISPID_PROPERTYPUT: i32 = -3;

/// Exact generic SDK `IDispatch` vtable order; deliberately not an Excel vtable.
#[repr(C)]
pub(super) struct IDispatchVtbl {
    pub(super) base: IUnknown_Vtbl,
    pub(super) get_type_info_count: unsafe extern "system" fn(*mut c_void, *mut u32) -> HRESULT,
    pub(super) get_type_info:
        unsafe extern "system" fn(*mut c_void, u32, u32, *mut *mut c_void) -> HRESULT,
    pub(super) get_ids_of_names: unsafe extern "system" fn(
        *mut c_void,
        *const GUID,
        *const *const u16,
        u32,
        u32,
        *mut i32,
    ) -> HRESULT,
    pub(super) invoke: unsafe extern "system" fn(
        *mut c_void,
        i32,
        *const GUID,
        u32,
        u16,
        *const DISPPARAMS,
        *mut VARIANT,
        *mut EXCEPINFO,
        *mut u32,
    ) -> HRESULT,
}

/// Keeps positional VARIANTs and named DISPIDs alive during one Invoke.
pub(super) struct Frame {
    args: Vec<OwnedVariant>,
    named: Vec<i32>,
}

impl Frame {
    pub(super) fn empty() -> Self {
        Self { args: vec![], named: vec![] }
    }

    pub(super) fn positional(mut args: Vec<OwnedVariant>) -> Self {
        args.reverse();
        Self { args, named: vec![] }
    }

    pub(super) fn put(value: OwnedVariant) -> Self {
        Self { args: vec![value], named: vec![DISPID_PROPERTYPUT] }
    }

    fn params(&mut self) -> DISPPARAMS {
        DISPPARAMS {
            rgvarg: if self.args.is_empty() { std::ptr::null_mut() } else { self.args.as_mut_ptr().cast() },
            rgdispidNamedArgs: if self.named.is_empty() { std::ptr::null_mut() } else { self.named.as_mut_ptr() },
            cArgs: self.args.len() as u32,
            cNamedArgs: self.named.len() as u32,
        }
    }
}

/// Copied result and diagnostics from a single generic `IDispatch::Invoke`.
pub(super) struct Call {
    pub(super) hr: HRESULT,
    pub(super) result: OwnedVariant,
    pub(super) exception: Value,
    pub(super) arg_error: u32,
}

pub(super) fn call(
    target: &ComPtr<Dispatch>,
    member: &str,
    flags: u16,
    lcid: u32,
    mut frame: Frame,
) -> Call {
    let mut result = OwnedVariant::empty();
    let mut exception = OwnedExcepInfo::new();
    let mut argument_error = u32::MAX;
    let params = frame.params();
    let status = match unsafe { target.dispid(member, lcid) } {
        Ok(dispid) => unsafe {
            (target.vtbl().invoke)(
                target.raw,
                dispid,
                &GUID::default(),
                lcid,
                flags,
                &params,
                &mut result.0,
                &mut exception.0,
                &mut argument_error,
            )
        },
        Err(status) => status,
    };
    let exception = exception.take();
    Call { hr: status, result, exception, arg_error: argument_error }
}
