use crate::ExcelComError;
use crate::automation::{OwnedVariant, invoke, property_get, property_put};
use crate::excel::DispatchObject;
use crate::internal::{ComPtr, Dispatch};
use crate::object_model::{MemberId, member};
use std::fmt::{Debug, Formatter};

/// Experimental wrapper for a single Excel Workbook.
pub struct Workbook {
    inner: DispatchObject,
}
impl Debug for Workbook {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_tuple("Workbook")
            .field(&self.inner)
            .finish()
    }
}
impl Clone for Workbook {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl Workbook {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "Workbook",
            },
        }
    }
    pub fn name(&self) -> Result<String, ExcelComError> {
        property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.workbook.name"), false),
            vec![],
        )?
        .as_string()
    }
    pub fn saved(&self) -> Result<bool, ExcelComError> {
        property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.workbook.saved"), false),
            vec![],
        )?
        .as_bool()
        .ok_or(ExcelComError::Conversion {
            detail: "Saved did not return VT_BOOL",
        })
    }
    pub fn set_saved(&self, value: bool) -> Result<(), ExcelComError> {
        let _ = property_put(
            &self.inner.dispatch,
            member(MemberId::new("excel.workbook.saved"), true),
            OwnedVariant::bool(value),
        )?;
        Ok(())
    }
    pub fn close_without_saving(self) -> Result<(), ExcelComError> {
        let _ = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.workbook.close"), false),
            vec![OwnedVariant::bool(false)],
            false,
        )?;
        Ok(())
    }
}
