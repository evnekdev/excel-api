use crate::ExcelComError;
use crate::automation::property_get;
use crate::excel::{DispatchObject, Workbook};
use crate::internal::{ComPtr, Dispatch};
use crate::object_model::{MemberId, member};
use std::fmt::{Debug, Formatter};

/// Experimental wrapper for an Excel Workbooks collection.
pub struct Workbooks {
    inner: DispatchObject,
}
impl Debug for Workbooks {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_tuple("Workbooks")
            .field(&self.inner)
            .finish()
    }
}
impl Clone for Workbooks {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl Workbooks {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "Workbooks",
            },
        }
    }
    pub fn count(&self) -> Result<i32, ExcelComError> {
        property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.workbooks.count"), false),
            vec![],
        )?
        .as_i32()
        .ok_or(ExcelComError::Conversion {
            detail: "Count did not return VT_I4",
        })
    }
    pub fn add(&self) -> Result<Workbook, ExcelComError> {
        let mut result = property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.workbooks.add"), false),
            vec![],
        )?;
        Ok(Workbook::from_dispatch(result.take_dispatch()?))
    }
}
