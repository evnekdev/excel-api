//! Centralized OLE Automation allocation and decoding.
//!
//! # Safety contract
//! Every raw SAFEARRAY pointer is non-null and borrowed only for the COM call,
//! or is owned by `OwnedSafeArray`. Bounds precede access. GetElement writes an
//! initialized VARIANT cleared once; PutElement copies initialized VT_I4 data.
//! BSTR length comes from Automation. `into_raw` transfers ownership to Excel;
//! otherwise Drop destroys the array once. Tests keep constructor-owned
//! pointers alive around every inspection.

use crate::model::RefreshItem;
use core::ffi::c_void;
use core::mem::ManuallyDrop;
use windows::Win32::Foundation::{E_INVALIDARG, E_OUTOFMEMORY};
use windows::Win32::System::Com::{SAFEARRAY, SAFEARRAYBOUND};
use windows::Win32::System::Ole::{
    SafeArrayCreate, SafeArrayDestroy, SafeArrayGetDim, SafeArrayGetElement, SafeArrayGetLBound,
    SafeArrayGetUBound, SafeArrayPutElement,
};
use windows::Win32::System::Variant::{VARIANT, VT_BSTR, VT_VARIANT};
use windows::core::{BSTR, HRESULT};

pub(crate) unsafe fn read_topic_components(
    strings: *mut *mut SAFEARRAY,
) -> Result<Vec<Vec<u16>>, HRESULT> {
    if strings.is_null() {
        return Err(E_INVALIDARG);
    }
    // SAFETY: the type-library contract supplies an in pointer to a SAFEARRAY
    // pointer for the duration of ConnectData.
    let array = unsafe { strings.read() };
    if array.is_null() || unsafe { SafeArrayGetDim(array) } != 1 {
        return Err(E_INVALIDARG);
    }
    let lower = unsafe { SafeArrayGetLBound(array, 1) }.map_err(|error| error.code())?;
    let upper = unsafe { SafeArrayGetUBound(array, 1) }.map_err(|error| error.code())?;
    if upper < lower {
        return Ok(Vec::new());
    }
    let count = usize::try_from(upper - lower + 1).map_err(|_| E_INVALIDARG)?;
    if count > 16 {
        return Err(E_INVALIDARG);
    }
    let mut components = Vec::with_capacity(count);
    for index in lower..=upper {
        let mut value = VARIANT::default();
        unsafe {
            SafeArrayGetElement(array, &index, (&mut value as *mut VARIANT).cast::<c_void>())
        }
        .map_err(|error| error.code())?;
        components.push(unsafe { copy_bstr(&value) }?);
    }
    Ok(components)
}

unsafe fn copy_bstr(value: &VARIANT) -> Result<Vec<u16>, HRESULT> {
    if value.vt() != VT_BSTR {
        return Err(E_INVALIDARG);
    }
    // SAFETY: VT_BSTR selects the bstrVal union member. The borrow does not
    // take ownership; the surrounding VARIANT remains responsible for clear.
    let raw = unsafe {
        &value.Anonymous.Anonymous.Anonymous.bstrVal as *const ManuallyDrop<BSTR> as *const BSTR
    };
    // SAFETY: `raw` points to the selected live BSTR member for this borrow.
    Ok(unsafe { (&*raw).to_vec() })
}

pub(crate) fn initial_counter(value: i32) -> VARIANT {
    value.into()
}

pub(crate) struct OwnedSafeArray(*mut SAFEARRAY);

impl OwnedSafeArray {
    pub(crate) fn refresh_payload(items: &[RefreshItem]) -> Result<Self, HRESULT> {
        let topic_count = u32::try_from(items.len()).map_err(|_| E_INVALIDARG)?;
        let bounds = [
            SAFEARRAYBOUND {
                cElements: 2,
                lLbound: 0,
            },
            SAFEARRAYBOUND {
                cElements: topic_count,
                lLbound: 0,
            },
        ];
        // SAFETY: both bounds live through the call and request a two-dimensional
        // VT_VARIANT array. The returned owner destroys it on every failure.
        let array = unsafe { SafeArrayCreate(VT_VARIANT, 2, bounds.as_ptr()) };
        if array.is_null() {
            return Err(E_OUTOFMEMORY);
        }
        let owner = Self(array);
        for (column, item) in items.iter().enumerate() {
            let column = i32::try_from(column).map_err(|_| E_INVALIDARG)?;
            owner.put([0, column], &VARIANT::from(item.topic_id))?;
            owner.put([1, column], &VARIANT::from(item.value))?;
        }
        Ok(owner)
    }

    fn put(&self, indices: [i32; 2], value: &VARIANT) -> Result<(), HRESULT> {
        // SAFETY: the array is a live VT_VARIANT SAFEARRAY and `value` remains
        // initialized for the duration of the copying API call.
        unsafe {
            SafeArrayPutElement(
                self.0,
                indices.as_ptr(),
                (value as *const VARIANT).cast::<c_void>(),
            )
        }
        .map_err(|error| error.code())
    }

    pub(crate) fn into_raw(self) -> *mut SAFEARRAY {
        let raw = self.0;
        core::mem::forget(self);
        raw
    }
}

impl Drop for OwnedSafeArray {
    fn drop(&mut self) {
        // SAFETY: this owner is unique until `into_raw`; SafeArrayDestroy clears
        // every copied VARIANT/BSTR element exactly once.
        unsafe {
            let _ = SafeArrayDestroy(self.0);
        }
    }
}

#[cfg(test)]
pub(crate) unsafe fn read_i32_cell(array: *mut SAFEARRAY, indices: [i32; 2]) -> i32 {
    let mut value = VARIANT::default();
    unsafe {
        SafeArrayGetElement(
            array,
            indices.as_ptr(),
            (&mut value as *mut VARIANT).cast::<c_void>(),
        )
        .unwrap();
    }
    i32::try_from(&value).unwrap()
}

#[cfg(test)]
mod tests {
    use super::*;
    use windows::Win32::System::Ole::{SafeArrayGetDim, SafeArrayGetLBound, SafeArrayGetUBound};

    #[test]
    fn refresh_payload_has_exact_two_by_n_shape_and_values() {
        let items = [
            RefreshItem {
                topic_id: 8,
                value: 12,
                version: 1,
            },
            RefreshItem {
                topic_id: 9,
                value: 13,
                version: 2,
            },
        ];
        let payload = OwnedSafeArray::refresh_payload(&items).unwrap();
        assert_eq!(unsafe { SafeArrayGetDim(payload.0) }, 2);
        assert_eq!(unsafe { SafeArrayGetLBound(payload.0, 1) }.unwrap(), 0);
        assert_eq!(unsafe { SafeArrayGetUBound(payload.0, 1) }.unwrap(), 1);
        assert_eq!(unsafe { SafeArrayGetLBound(payload.0, 2) }.unwrap(), 0);
        assert_eq!(unsafe { SafeArrayGetUBound(payload.0, 2) }.unwrap(), 1);
        assert_eq!(unsafe { read_i32_cell(payload.0, [0, 0]) }, 8);
        assert_eq!(unsafe { read_i32_cell(payload.0, [1, 0]) }, 12);
        assert_eq!(unsafe { read_i32_cell(payload.0, [0, 1]) }, 9);
        assert_eq!(unsafe { read_i32_cell(payload.0, [1, 1]) }, 13);
    }

    #[test]
    fn repeated_payload_construction_and_clear_is_stable() {
        let item = [RefreshItem {
            topic_id: 1,
            value: 2,
            version: 3,
        }];
        for _ in 0..1_000 {
            drop(OwnedSafeArray::refresh_payload(&item).unwrap());
        }
    }

    #[test]
    fn zero_update_payload_is_a_two_by_zero_safearray() {
        let payload = OwnedSafeArray::refresh_payload(&[]).unwrap();
        assert_eq!(unsafe { SafeArrayGetDim(payload.0) }, 2);
        assert_eq!(unsafe { SafeArrayGetLBound(payload.0, 1) }.unwrap(), 0);
        assert_eq!(unsafe { SafeArrayGetUBound(payload.0, 1) }.unwrap(), 1);
        assert_eq!(unsafe { SafeArrayGetLBound(payload.0, 2) }.unwrap(), 0);
        assert_eq!(unsafe { SafeArrayGetUBound(payload.0, 2) }.unwrap(), -1);
    }

    #[test]
    fn bstr_copy_preserves_embedded_nul_and_non_ascii_utf16() {
        let units = [b'A' as u16, 0, 0x03bb, 0xd83d, 0xde00];
        let value = VARIANT::from(BSTR::from_wide(&units));
        assert_eq!(unsafe { copy_bstr(&value) }.unwrap(), units);
    }
}
