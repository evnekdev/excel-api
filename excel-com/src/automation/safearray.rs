//! Private SAFEARRAY ownership using only SDK APIs.

use core::ffi::c_void;
use std::ptr::NonNull;

use windows_sys::Win32::System::Com::{SAFEARRAY, SAFEARRAYBOUND};
use windows_sys::Win32::System::Ole::{
    SafeArrayCreate, SafeArrayDestroy, SafeArrayGetDim, SafeArrayGetElement, SafeArrayGetLBound,
    SafeArrayGetUBound, SafeArrayGetVartype, SafeArrayPutElement,
};
use windows_sys::Win32::System::Variant::{VARIANT, VT_VARIANT};

use super::OwnedVariant;

/// Metadata copied from a borrowed SAFEARRAY without exposing its address.
pub(crate) struct ArrayMetadata {
    pub(crate) rank: u32,
    pub(crate) dimensions: Vec<ArrayDimension>,
    pub(crate) element_vartype: Option<u16>,
}

/// A physical SAFEARRAY dimension.
pub(crate) struct ArrayDimension {
    pub(crate) lower_bound: i32,
    pub(crate) element_count: u32,
}

/// An owned SAFEARRAY transferred into or out of a VARIANT exactly once.
pub(crate) struct SafeArray(NonNull<SAFEARRAY>);

impl SafeArray {
    pub(crate) fn create_variant(bounds: &[SAFEARRAYBOUND]) -> Option<Self> {
        let rank = u32::try_from(bounds.len()).ok()?;
        (rank != 0).then_some(())?;
        // SAFETY: the nonempty bounds slice has `rank` valid SAFEARRAYBOUND entries.
        let raw = unsafe { SafeArrayCreate(VT_VARIANT, rank, bounds.as_ptr()) };
        Self::from_owned(raw)
    }

    pub(crate) fn from_owned(raw: *mut SAFEARRAY) -> Option<Self> {
        NonNull::new(raw).map(Self)
    }

    pub(crate) fn into_raw(self) -> *mut SAFEARRAY {
        let raw = self.0.as_ptr();
        std::mem::forget(self);
        raw
    }

    pub(crate) fn put_variant(&self, indices: &[i32], value: &OwnedVariant) -> bool {
        // SAFETY: this owner holds a valid descriptor and the initialized VARIANT stays valid for the call.
        unsafe {
            SafeArrayPutElement(
                self.0.as_ptr(),
                indices.as_ptr(),
                &value.0 as *const VARIANT as *const c_void,
            ) == 0
        }
    }

    pub(crate) fn get_variant_borrowed(
        raw: *mut SAFEARRAY,
        indices: &[i32],
    ) -> Option<OwnedVariant> {
        (!raw.is_null()).then_some(())?;
        let mut result = OwnedVariant::empty();
        // SAFETY: the borrowed descriptor and caller-provided indices are valid for the SDK call.
        let status = unsafe {
            SafeArrayGetElement(
                raw,
                indices.as_ptr(),
                &mut result.0 as *mut VARIANT as *mut c_void,
            )
        };
        (status == 0).then_some(result)
    }

    pub(crate) fn metadata(raw: *mut SAFEARRAY) -> Option<ArrayMetadata> {
        (!raw.is_null()).then_some(())?;
        // SAFETY: `raw` is checked non-null and is borrowed only for SDK metadata inspection.
        let rank = unsafe { SafeArrayGetDim(raw) };
        let mut dimensions = Vec::with_capacity(rank as usize);
        for dimension in 1..=rank {
            let mut lower_bound = 0;
            let mut upper_bound = -1;
            // SAFETY: the descriptor is borrowed and both output pointers are writable local storage.
            let lower_status = unsafe { SafeArrayGetLBound(raw, dimension, &mut lower_bound) };
            // SAFETY: the descriptor is borrowed and both output pointers are writable local storage.
            let upper_status = unsafe { SafeArrayGetUBound(raw, dimension, &mut upper_bound) };
            if lower_status != 0 || upper_status != 0 {
                return None;
            }
            let element_count = upper_bound
                .checked_sub(lower_bound)?
                .checked_add(1)
                .and_then(|count| u32::try_from(count).ok())?;
            dimensions.push(ArrayDimension {
                lower_bound,
                element_count,
            });
        }
        let mut element_vartype = 0;
        // SAFETY: the borrowed descriptor and output VARTYPE storage are valid for this call.
        let vartype_status = unsafe { SafeArrayGetVartype(raw, &mut element_vartype) };
        let element_vartype = (vartype_status == 0).then_some(element_vartype);
        Some(ArrayMetadata {
            rank,
            dimensions,
            element_vartype,
        })
    }
}

impl Drop for SafeArray {
    fn drop(&mut self) {
        // SAFETY: this unique owner is constructed from a SAFEARRAY ownership transfer.
        unsafe {
            let _ = SafeArrayDestroy(self.0.as_ptr());
        }
    }
}
