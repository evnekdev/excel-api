use std::ffi::c_void;
use std::marker::PhantomData;
use std::ptr::NonNull;
use std::rc::Rc;

use windows_sys::core::GUID;
use windows_sys::core::IUnknown_Vtbl;

use crate::ExcelComError;

/// Marker for generic Automation dispatch references.
pub(crate) enum Dispatch {}
/// Marker for an owned IUnknown reference that is never dispatched through.
pub(crate) enum Unknown {}
/// Marker for an owned IEnumVARIANT reference.
pub(crate) enum EnumVariantInterface {}

pub(crate) const IID_IUNKNOWN: GUID = GUID::from_u128(0x00000000_0000_0000_c000_000000000046);

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
    /// Transfers this one owned COM reference to an ABI owner such as VARIANT.
    pub(crate) fn into_raw(self) -> *mut c_void {
        let raw = self.raw();
        std::mem::forget(self);
        raw
    }
    unsafe fn iunknown_vtbl(&self) -> &IUnknown_Vtbl {
        // SAFETY: every owned interface starts with the IUnknown vtable prefix.
        unsafe { &**(self.raw() as *const *const IUnknown_Vtbl) }
    }
    pub(crate) fn query_interface<U>(&self, iid: &GUID) -> Result<ComPtr<U>, ExcelComError> {
        let mut raw = std::ptr::null_mut();
        // SAFETY: the interface, IID, and writable output pointer are valid.
        let status = unsafe { (self.iunknown_vtbl().QueryInterface)(self.raw(), iid, &mut raw) };
        if ExcelComError::failed(status) {
            return Err(ExcelComError::QueryInterface { hresult: status });
        }
        // SAFETY: a successful QueryInterface returns one owned reference.
        unsafe { ComPtr::from_owned(raw) }
    }
}

impl ComPtr<Dispatch> {
    pub(crate) unsafe fn vtbl(&self) -> &DispatchVtbl {
        // SAFETY: this specialization only owns IDispatch pointers.
        unsafe { &**(self.raw() as *const *const DispatchVtbl) }
    }

    pub(crate) fn canonical_unknown(&self) -> Result<ComPtr<Unknown>, ExcelComError> {
        self.query_interface(&IID_IUNKNOWN)
    }
}

impl<T> Clone for ComPtr<T> {
    fn clone(&self) -> Self {
        // SAFETY: this is a valid interface reference; AddRef establishes cloned ownership.
        unsafe { (self.iunknown_vtbl().AddRef)(self.raw()) };
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
        unsafe { (self.iunknown_vtbl().Release)(self.raw()) };
    }
}
