//! Private `IEnumVARIANT` ownership and one-item iteration.

use std::ffi::c_void;

use windows_sys::Win32::System::Variant::VARIANT;
use windows_sys::core::{GUID, IUnknown_Vtbl};

use super::{OwnedVariant, property_get};
use crate::{
    ExcelComError,
    excel::DispatchObject,
    internal::{ComPtr, Dispatch, EnumVariantInterface},
    object_model::MemberId,
};

const IID_IENUMVARIANT: GUID = GUID::from_u128(0x00020404_0000_0000_c000_000000000046);
const S_OK: i32 = 0;
const S_FALSE: i32 = 1;

#[repr(C)]
struct EnumVariantVtbl {
    base: IUnknown_Vtbl,
    next: unsafe extern "system" fn(*mut c_void, u32, *mut VARIANT, *mut u32) -> i32,
    skip: unsafe extern "system" fn(*mut c_void, u32) -> i32,
    reset: unsafe extern "system" fn(*mut c_void) -> i32,
    clone: unsafe extern "system" fn(*mut c_void, *mut *mut c_void) -> i32,
}

/// Private owning `IEnumVARIANT` wrapper. It never exposes raw interfaces.
pub(crate) struct EnumVariant {
    inner: ComPtr<EnumVariantInterface>,
    collection: &'static str,
    next_index: usize,
}

impl EnumVariant {
    pub(crate) fn from_new_enum(
        target: &DispatchObject,
        member: MemberId,
        collection: &'static str,
    ) -> Result<Self, ExcelComError> {
        let mut result = property_get(
            &target.dispatch,
            crate::object_model::member(member, false),
            vec![],
        )?;
        let enum_pointer = match result.vt() {
            windows_sys::Win32::System::Variant::VT_UNKNOWN => {
                result.take_unknown()?.query_interface(&IID_IENUMVARIANT)?
            }
            windows_sys::Win32::System::Variant::VT_DISPATCH => {
                result.take_dispatch()?.query_interface(&IID_IENUMVARIANT)?
            }
            _ => {
                return Err(ExcelComError::Enumeration {
                    collection,
                    item_index: 0,
                    hresult: None,
                    detail: "_NewEnum did not return VT_UNKNOWN or VT_DISPATCH",
                });
            }
        };
        Ok(Self {
            inner: enum_pointer,
            collection,
            next_index: 0,
        })
    }

    pub(crate) fn next(&mut self) -> Result<Option<OwnedVariant>, ExcelComError> {
        let mut value = OwnedVariant::empty();
        let mut fetched = 0u32;
        // SAFETY: IEnumVARIANT has the documented vtable prefix and all output storage is valid.
        let status = unsafe { (self.vtbl().next)(self.inner.raw(), 1, &mut value.0, &mut fetched) };
        match status {
            S_OK if fetched == 1 => {
                self.next_index += 1;
                Ok(Some(value))
            }
            S_FALSE if fetched == 0 => Ok(None),
            S_OK | S_FALSE => Err(ExcelComError::Enumeration {
                collection: self.collection,
                item_index: self.next_index,
                hresult: Some(status),
                detail: "IEnumVARIANT::Next returned an invalid fetched count",
            }),
            _ => Err(ExcelComError::Enumeration {
                collection: self.collection,
                item_index: self.next_index,
                hresult: Some(status),
                detail: "IEnumVARIANT::Next failed",
            }),
        }
    }

    unsafe fn vtbl(&self) -> &EnumVariantVtbl {
        // SAFETY: construction QueryInterfaces explicitly for IID_IEnumVARIANT.
        unsafe { &**(self.inner.raw() as *const *const EnumVariantVtbl) }
    }
}

pub(crate) fn enumerated_dispatch(
    value: &mut OwnedVariant,
    collection: &'static str,
    item_index: usize,
) -> Result<ComPtr<Dispatch>, ExcelComError> {
    match value.vt() {
        windows_sys::Win32::System::Variant::VT_DISPATCH => {
            value
                .take_dispatch()
                .map_err(|_| ExcelComError::Enumeration {
                    collection,
                    item_index,
                    hresult: None,
                    detail: "enumerator yielded a null dispatch reference",
                })
        }
        windows_sys::Win32::System::Variant::VT_UNKNOWN => value
            .take_unknown()?
            .query_interface(&GUID::from_u128(0x00020400_0000_0000_c000_000000000046)),
        _ => Err(ExcelComError::Enumeration {
            collection,
            item_index,
            hresult: None,
            detail: "enumerator yielded a non-object VARIANT",
        }),
    }
}
