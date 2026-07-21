//! Owning generic `IDispatch` references; no Excel-specific interface layout.

use std::ffi::c_void;
use std::marker::PhantomData;

use windows_sys::core::HRESULT;

use super::dispatch::IDispatchVtbl;

/// Marker for an Automation `IDispatch` interface.
pub(super) struct Dispatch;

/// Owns exactly one COM interface reference.
pub(super) struct ComPtr<T> {
    pub(super) raw: *mut c_void,
    pub(super) _marker: PhantomData<T>,
}

impl<T> ComPtr<T> {
    /// Reads the SDK-layout generic `IDispatch` vtable for this reference.
    pub(super) unsafe fn vtbl(&self) -> &IDispatchVtbl {
        unsafe { &**(self.raw as *const *const IDispatchVtbl) }
    }

    /// Takes a COM reference already owned by the caller.
    ///
    /// `raw` must be a non-null `IDispatch` pointer with one reference that this
    /// owner is responsible for releasing.
    pub(super) unsafe fn from_owned(raw: *mut c_void) -> Option<Self> {
        (!raw.is_null()).then_some(Self {
            raw,
            _marker: PhantomData,
        })
    }

    /// Clones a borrowed `IDispatch` reference with `AddRef`.
    ///
    /// `raw` must be null or point to a valid SDK-layout `IDispatch`.
    pub(super) unsafe fn from_borrowed(raw: *mut c_void) -> Option<Self> {
        if raw.is_null() {
            None
        } else {
            unsafe { ((&**(raw as *const *const IDispatchVtbl)).base.AddRef)(raw) };
            Some(Self {
                raw,
                _marker: PhantomData,
            })
        }
    }

    /// Resolves a member name through the generic Automation vtable.
    pub(super) unsafe fn dispid(&self, name: &str, lcid: u32) -> Result<i32, HRESULT> {
        let wide: Vec<u16> = name.encode_utf16().chain(Some(0)).collect();
        let names = [wide.as_ptr()];
        let mut id = 0;
        let status = unsafe {
            (self.vtbl().get_ids_of_names)(
                self.raw,
                &windows_sys::core::GUID::default(),
                names.as_ptr(),
                1,
                lcid,
                &mut id,
            )
        };
        if status == 0 { Ok(id) } else { Err(status) }
    }
}

impl<T> Drop for ComPtr<T> {
    fn drop(&mut self) {
        if !self.raw.is_null() {
            unsafe { (self.vtbl().base.Release)(self.raw) };
        }
    }
}
