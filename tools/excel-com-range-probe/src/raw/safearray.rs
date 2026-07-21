//! SAFEARRAY ownership and SDK metadata inspection for the Value/Value2 matrix.

use core::ffi::c_void;

use serde::Serialize;
use windows_sys::Win32::System::Com::{SAFEARRAY, SAFEARRAYBOUND};
use windows_sys::Win32::System::Ole::{
    SafeArrayAccessData, SafeArrayCreate, SafeArrayCreateVector, SafeArrayDestroy,
    SafeArrayGetDim, SafeArrayGetElement, SafeArrayGetLBound, SafeArrayGetUBound,
    SafeArrayGetVartype, SafeArrayPutElement, SafeArrayUnaccessData,
};
use windows_sys::Win32::System::Variant::{VARIANT, VT_VARIANT};

use super::variant::OwnedVariant;

/// Bounds and element VARTYPE copied from the SDK APIs without pointer values.
#[derive(Clone, Debug, Eq, PartialEq, Serialize)]
pub(super) struct ObservedSafeArray {
    pub(super) rank: u32,
    pub(super) dimensions: Vec<ObservedSafeArrayDimension>,
    pub(super) element_vartype: Option<u16>,
    pub(super) metadata_dimension_order: &'static str,
    pub(super) storage_traversal: &'static str,
    pub(super) access_balanced: bool,
}

/// One SAFEARRAY dimension in physical COM order.
#[derive(Clone, Debug, Eq, PartialEq, Serialize)]
pub(super) struct ObservedSafeArrayDimension {
    pub(super) physical_dimension: u32,
    pub(super) lower_bound: i32,
    pub(super) upper_bound: i32,
    pub(super) element_count: u32,
}

impl ObservedSafeArray {
    /// Inspects metadata through SDK SAFEARRAY APIs and balances Access/Unaccess.
    ///
    /// `value` must be null or point to a valid SAFEARRAY for the duration of
    /// this call. The method never takes ownership of the descriptor.
    pub(super) unsafe fn inspect(value: *mut SAFEARRAY) -> Option<Self> {
        if value.is_null() {
            return None;
        }
        let rank = unsafe { SafeArrayGetDim(value) };
        let mut dimensions = Vec::with_capacity(rank as usize);
        for physical_dimension in 1..=rank {
            let mut lower_bound = 0;
            let mut upper_bound = -1;
            if unsafe { SafeArrayGetLBound(value, physical_dimension, &mut lower_bound) } != 0
                || unsafe { SafeArrayGetUBound(value, physical_dimension, &mut upper_bound) } != 0
            {
                return None;
            }
            let element_count = upper_bound
                .checked_sub(lower_bound)
                .and_then(|value| value.checked_add(1))
                .and_then(|value| u32::try_from(value).ok())?;
            dimensions.push(ObservedSafeArrayDimension {
                physical_dimension,
                lower_bound,
                upper_bound,
                element_count,
            });
        }
        let mut element_vartype = 0;
        let element_vartype = (unsafe { SafeArrayGetVartype(value, &mut element_vartype) } == 0)
            .then_some(element_vartype);
        let mut data = std::ptr::null_mut();
        let access_balanced = if (unsafe { SafeArrayAccessData(value, &mut data) }) == 0 {
            (unsafe { SafeArrayUnaccessData(value) }) == 0
        } else {
            false
        };
        Some(Self {
            rank,
            dimensions,
            element_vartype,
            metadata_dimension_order: "SafeArrayGetLBound/GetUBound dimensions 1..rank",
            storage_traversal:
                "SDK SafeArrayGetElement/PutElement index order [physical_dimension_1, ...]",
            access_balanced,
        })
    }
}

/// Owns a SAFEARRAY returned with transferable ownership by Automation.
#[allow(dead_code)]
pub(super) struct OwnedSafeArray(*mut SAFEARRAY);

#[allow(dead_code)]
impl OwnedSafeArray {
    /// Creates an owned `SAFEARRAY(VARIANT)`. Its elements are initialized by
    /// `SafeArrayPutElement`; a partial construction is still destroyed by
    /// this owner's `Drop` implementation.
    pub(super) fn create_variant(bounds: &[SAFEARRAYBOUND]) -> Result<Self, String> {
        let rank = u32::try_from(bounds.len()).map_err(|_| "SAFEARRAY rank overflow")?;
        if rank == 0 {
            return Err("SAFEARRAY rank must be at least one".to_owned());
        }
        let value = unsafe { SafeArrayCreate(VT_VARIANT, rank, bounds.as_ptr()) };
        unsafe { Self::from_owned(value) }.ok_or_else(|| "SafeArrayCreate returned null".to_owned())
    }

    /// Creates a one-dimensional `SAFEARRAY(VARIANT)` through the SDK vector
    /// constructor, retaining the explicit lower bound.
    pub(super) fn create_variant_vector(lower_bound: i32, element_count: u32) -> Result<Self, String> {
        let value = unsafe { SafeArrayCreateVector(VT_VARIANT, lower_bound, element_count) };
        unsafe { Self::from_owned(value) }
            .ok_or_else(|| "SafeArrayCreateVector returned null".to_owned())
    }

    /// Takes ownership of a SAFEARRAY pointer returned by Automation.
    ///
    /// `value` must be null or be destroyable exactly once with `SafeArrayDestroy`.
    pub(super) unsafe fn from_owned(value: *mut SAFEARRAY) -> Option<Self> {
        (!value.is_null()).then_some(Self(value))
    }

    pub(super) fn as_ptr(&self) -> *mut SAFEARRAY {
        self.0
    }

    pub(super) fn into_raw(mut self) -> *mut SAFEARRAY {
        let value = self.0;
        self.0 = std::ptr::null_mut();
        value
    }

    /// Writes a fully initialized `VARIANT` through the SDK API. `indices`
    /// are ordered by physical SAFEARRAY dimension, never by a guessed Excel
    /// row/column convention.
    pub(super) fn put_variant(
        &self,
        indices: &[i32],
        value: &OwnedVariant,
    ) -> Result<(), String> {
        let result = unsafe {
            SafeArrayPutElement(
                self.0,
                indices.as_ptr(),
                &value.0 as *const VARIANT as *const c_void,
            )
        };
        (result == 0).then_some(()).ok_or_else(|| {
            format!("SafeArrayPutElement failed with 0x{:08X}", result as u32)
        })
    }

    /// Retrieves a copied, initialized `VARIANT` through the SDK API.
    pub(super) fn get_variant(&self, indices: &[i32]) -> Result<OwnedVariant, String> {
        let mut value = OwnedVariant::empty();
        let result = unsafe {
            SafeArrayGetElement(
                self.0,
                indices.as_ptr(),
                &mut value.0 as *mut VARIANT as *mut c_void,
            )
        };
        (result == 0)
            .then_some(value)
            .ok_or_else(|| format!("SafeArrayGetElement failed with 0x{:08X}", result as u32))
    }
}

pub(super) fn checked_element_count(dimensions: &[ObservedSafeArrayDimension]) -> Option<usize> {
    dimensions.iter().try_fold(1_usize, |total, dimension| {
        total.checked_mul(usize::try_from(dimension.element_count).ok()?)
    })
}

/// Converts a zero-based logical row/column position into physical SAFEARRAY
/// indices only after the caller has established that physical dimension one
/// maps to rows and dimension two maps to columns.
pub(super) fn row_column_indices(
    dimensions: &[ObservedSafeArrayDimension],
    row: u32,
    column: u32,
) -> Option<[i32; 2]> {
    let [rows, columns]: &[ObservedSafeArrayDimension; 2] = dimensions.try_into().ok()?;
    (row < rows.element_count && column < columns.element_count).then_some([
        rows.lower_bound.checked_add(i32::try_from(row).ok()?)?,
        columns.lower_bound.checked_add(i32::try_from(column).ok()?)?,
    ])
}

impl Drop for OwnedSafeArray {
    fn drop(&mut self) {
        if !self.0.is_null() {
            unsafe {
                let _ = SafeArrayDestroy(self.0);
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn metadata_preserves_non_zero_bounds_in_physical_order() {
        let bounds = [
            SAFEARRAYBOUND {
                cElements: 2,
                lLbound: 7,
            },
            SAFEARRAYBOUND {
                cElements: 3,
                lLbound: -2,
            },
        ];
        let raw = unsafe { SafeArrayCreate(VT_VARIANT, 2, bounds.as_ptr()) };
        let owner = unsafe { OwnedSafeArray::from_owned(raw) }.expect("SAFEARRAY allocation");
        let layout = unsafe { ObservedSafeArray::inspect(owner.as_ptr()) }.expect("metadata");
        assert_eq!(layout.rank, 2);
        assert_eq!(
            layout.dimensions,
            vec![
                ObservedSafeArrayDimension {
                    physical_dimension: 1,
                    lower_bound: 7,
                    upper_bound: 8,
                    element_count: 2,
                },
                ObservedSafeArrayDimension {
                    physical_dimension: 2,
                    lower_bound: -2,
                    upper_bound: 0,
                    element_count: 3,
                },
            ]
        );
        assert!(layout.access_balanced);
        assert_eq!(checked_element_count(&layout.dimensions), Some(6));
        assert_eq!(row_column_indices(&layout.dimensions, 1, 2), Some([8, 0]));
    }

    #[test]
    fn vector_and_element_round_trip_use_sdk_apis() {
        let owner = OwnedSafeArray::create_variant_vector(-3, 2).expect("SAFEARRAY vector");
        let first = OwnedVariant::i4(41);
        let second = OwnedVariant::i4(42);
        owner.put_variant(&[-3], &first).expect("first element");
        owner.put_variant(&[-2], &second).expect("second element");
        assert_eq!(owner.get_variant(&[-3]).expect("first read").i4_value(), Some(41));
        assert_eq!(owner.get_variant(&[-2]).expect("second read").i4_value(), Some(42));
    }
}
