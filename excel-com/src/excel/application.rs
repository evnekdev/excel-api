use crate::automation::{OwnedVariant, activate_excel, invoke, property_get, property_put};
use crate::excel::{DispatchObject, Workbooks};
use crate::object_model::{MemberId, member};
use crate::{ComApartment, ExcelComError};
use std::fmt::{Debug, Formatter};

/// Experimental wrapper for a crate-created local Excel Application instance.
pub struct Application {
    inner: DispatchObject,
}
impl Debug for Application {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_tuple("Application")
            .field(&self.inner)
            .finish()
    }
}
impl Clone for Application {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl Application {
    pub fn new(apartment: &ComApartment) -> Result<Self, ExcelComError> {
        apartment.assert_current()?;
        Ok(Self {
            inner: DispatchObject {
                dispatch: activate_excel()?,
                kind: "Application",
            },
        })
    }
    pub fn version(&self) -> Result<String, ExcelComError> {
        property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.application.version"), false),
            vec![],
        )?
        .as_string()
    }
    pub fn visible(&self) -> Result<bool, ExcelComError> {
        property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.application.visible"), false),
            vec![],
        )?
        .as_bool()
        .ok_or(ExcelComError::Conversion {
            detail: "Visible did not return VT_BOOL",
        })
    }
    pub fn set_visible(&self, value: bool) -> Result<(), ExcelComError> {
        let _ = property_put(
            &self.inner.dispatch,
            member(MemberId::new("excel.application.visible"), true),
            OwnedVariant::bool(value),
        )?;
        Ok(())
    }
    pub fn workbooks(&self) -> Result<Workbooks, ExcelComError> {
        let mut result = property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.application.workbooks"), false),
            vec![],
        )?;
        Ok(Workbooks::from_dispatch(result.take_dispatch()?))
    }
    /// Explicitly asks the crate-created application to quit. `Drop` never does this.
    pub fn quit(self) -> Result<(), ExcelComError> {
        let _ = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.application.quit"), false),
            vec![],
            false,
        )?;
        Ok(())
    }
}
