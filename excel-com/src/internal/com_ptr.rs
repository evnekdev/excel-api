use std::ffi::c_void;
use std::marker::PhantomData;
use std::ptr::NonNull;
use std::rc::Rc;

use windows_sys::core::IUnknown_Vtbl;

use crate::ExcelComError;

/// Marker for generic Automation dispatch references.
pub(crate) enum Dispatch {}

#[repr(C)]
pub(crate) struct DispatchVtbl {
    pub(crate) base: IUnknown_Vtbl,
    pub(crate) get_type_info_count: unsafe extern "system" fn(*mut c_void, *mut u32) -> i32,
    pub(crate) get_type_info:
        unsafe extern "system" fn(*mut c_void, u32, u32, *mut *mut c_void) -> i32,
    pub(crate) get_ids_of_names: unsafe extern "system" fn(
        *mut c_void,
        *const windows_sys::core::GUID,
        *const *const u16,
        u32,
        u32,
        *mut i32,
    ) -> i32,
    pub(crate) invoke: unsafe extern "system" fn(
        *mut c_void,
        i32,
        *const windows_sys::core::GUID,
        u32,
        u16,
        *const windows_sys::Win32::System::Com::DISPPARAMS,
        *mut windows_sys::Win32::System::Variant::VARIANT,
        *mut windows_sys::Win32::System::Com::EXCEPINFO,
        *mut u32,
    ) -> i32,
}

/// Exactly-one-reference COM owner. It is deliberately private and !Send/!Sync.
pub(crate) struct ComPtr<T> {
    raw: NonNull<c_void>,
    _type: PhantomData<T>,
    _not_send_or_sync: PhantomData<Rc<()>>,
}

impl<T> ComPtr<T> {
    pub(crate) unsafe fn from_owned(raw: *mut c_void) -> Result<Self, ExcelComError> {
        let raw = NonNull::new(raw).ok_or(ExcelComError::Ownership {
            detail: "COM activation returned a null interface",
        })?;
        Ok(Self {
            raw,
            _type: PhantomData,
            _not_send_or_sync: PhantomData,
        })
    }
    pub(crate) fn raw(&self) -> *mut c_void {
        self.raw.as_ptr()
    }
    pub(crate) unsafe fn vtbl(&self) -> &DispatchVtbl {
        // SAFETY: constructors require an owned generic IDispatch pointer with this vtable layout.
        unsafe { &**(self.raw() as *const *const DispatchVtbl) }
    }
}

impl<T> Clone for ComPtr<T> {
    fn clone(&self) -> Self {
        // SAFETY: this is a valid IDispatch reference; AddRef establishes the cloned ownership.
        unsafe { (self.vtbl().base.AddRef)(self.raw()) };
        Self {
            raw: self.raw,
            _type: PhantomData,
            _not_send_or_sync: PhantomData,
        }
    }
}

impl<T> Drop for ComPtr<T> {
    fn drop(&mut self) {
        // SAFETY: this owner holds exactly one reference that has not yet been released.
        unsafe { (self.vtbl().base.Release)(self.raw()) };
    }
}
