use crate::automation::{OwnedVariant, activate_excel, invoke, property_get, property_put};
use crate::excel::{DispatchObject, Workbooks};
use crate::object_model::{MemberId, member};
use crate::{ComApartment, ConversionError, ExcelComError};
use std::fmt::{Debug, Formatter};

/// Restores an [`Application`]'s prior `DisplayAlerts` value on drop.
///
/// Restoration failures cannot be returned from `Drop`; call [`Self::restore`]
/// when the caller needs to observe the result directly.
pub struct DisplayAlertsGuard<'a> {
    application: &'a Application,
    previous: bool,
    active: bool,
}

impl Debug for DisplayAlertsGuard<'_> {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("DisplayAlertsGuard")
            .field("previous", &self.previous)
            .field("active", &self.active)
            .finish()
    }
}

impl DisplayAlertsGuard<'_> {
    /// Restores the prior `DisplayAlerts` value and disarms the guard.
    pub fn restore(mut self) -> Result<(), ExcelComError> {
        self.application.set_display_alerts(self.previous)?;
        self.active = false;
        Ok(())
    }
}

impl Drop for DisplayAlertsGuard<'_> {
    fn drop(&mut self) {
        if self.active {
            let _ = self.application.set_display_alerts(self.previous);
        }
    }
}

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
    /// Starts a fresh local Excel `Application` in the supplied STA apartment.
    pub fn new(apartment: &ComApartment) -> Result<Self, ExcelComError> {
        apartment.assert_current()?;
        Ok(Self {
            inner: DispatchObject {
                dispatch: activate_excel()?,
                kind: "Application",
            },
        })
    }
    /// Returns the server's Excel version string.
    pub fn version(&self) -> Result<String, ExcelComError> {
        property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.application.version"), false),
            vec![],
        )?
        .as_string()
    }
    /// Returns the current visibility of the crate-created Excel window.
    pub fn visible(&self) -> Result<bool, ExcelComError> {
        property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.application.visible"), false),
            vec![],
        )?
        .as_bool()
        .ok_or(ExcelComError::Conversion(
            ConversionError::UnsupportedVariantType { vartype: 0 },
        ))
    }
    /// Sets the visibility of the crate-created Excel window.
    pub fn set_visible(&self, value: bool) -> Result<(), ExcelComError> {
        let _ = property_put(
            &self.inner.dispatch,
            member(MemberId::new("excel.application.visible"), true),
            OwnedVariant::bool(value),
        )?;
        Ok(())
    }
    /// Returns Excel's `DisplayAlerts` setting.
    pub fn display_alerts(&self) -> Result<bool, ExcelComError> {
        let result = property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.application.displayalerts"), false),
            vec![],
        )?;
        result.as_bool().ok_or(ExcelComError::Conversion(
            ConversionError::UnsupportedVariantType {
                vartype: result.vt(),
            },
        ))
    }
    /// Sets Excel's `DisplayAlerts` setting.
    pub fn set_display_alerts(&self, value: bool) -> Result<(), ExcelComError> {
        let _ = property_put(
            &self.inner.dispatch,
            member(MemberId::new("excel.application.displayalerts"), true),
            OwnedVariant::bool(value),
        )?;
        Ok(())
    }
    /// Sets `DisplayAlerts` and returns a guard that restores its prior value.
    pub fn display_alerts_guard(
        &self,
        value: bool,
    ) -> Result<DisplayAlertsGuard<'_>, ExcelComError> {
        let previous = self.display_alerts()?;
        self.set_display_alerts(value)?;
        Ok(DisplayAlertsGuard {
            application: self,
            previous,
            active: true,
        })
    }
    /// Returns the application's `Workbooks` collection.
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
