//! Redacted ODBC connection metadata.

use std::fmt::{Debug, Formatter};

use crate::automation::{decode_variant, invoke, property_get};
use crate::excel::DispatchObject;
use crate::internal::{ComPtr, Dispatch};
use crate::object_model::{MemberId, member};
use crate::{AutomationValue, ConversionPolicy, ExcelComError};

use super::helpers::{bool_value, integer, string};
use super::oledb::{put_bool, put_i32};
use super::{CommandType, SecretStringValue};

/// Apartment-bound Excel ODBC connection metadata.
///
/// Connection strings remain redacted by `Debug`, `Display`, and normal
/// invocation diagnostics. The wrapper never provides a password setter.
pub struct OdbcConnection {
    inner: DispatchObject,
}
impl Debug for OdbcConnection {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("OdbcConnection").field(&self.inner).finish()
    }
}
impl Clone for OdbcConnection {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl OdbcConnection {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "ODBCConnection",
            },
        }
    }
    /// Returns the connection string in a redacting wrapper.
    pub fn connection_string(&self) -> Result<SecretStringValue, ExcelComError> {
        SecretStringValue::new(string(&self.inner, "excel.odbcconnection.connection")?)
    }
    /// Returns command text without coercing an array-shaped provider result.
    pub fn command_text(&self) -> Result<AutomationValue, ExcelComError> {
        let value = property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.odbcconnection.commandtext"), false),
            vec![],
        )?;
        decode_variant(&value, ConversionPolicy::default())
    }
    /// Returns the forward-compatible provider command type.
    pub fn command_type(&self) -> Result<CommandType, ExcelComError> {
        Ok(CommandType::from_raw(integer(
            &self.inner,
            "excel.odbcconnection.commandtype",
        )?))
    }
    /// Returns whether background refresh is enabled.
    pub fn background_query(&self) -> Result<bool, ExcelComError> {
        bool_value(&self.inner, "excel.odbcconnection.backgroundquery")
    }
    /// Enables or disables background refresh.
    pub fn set_background_query(&self, enabled: bool) -> Result<(), ExcelComError> {
        put_bool(&self.inner, "excel.odbcconnection.backgroundquery", enabled)
    }
    /// Returns whether a background refresh is active.
    pub fn refreshing(&self) -> Result<bool, ExcelComError> {
        bool_value(&self.inner, "excel.odbcconnection.refreshing")
    }
    /// Returns whether Excel refreshes this connection when opening its workbook.
    pub fn refresh_on_file_open(&self) -> Result<bool, ExcelComError> {
        bool_value(&self.inner, "excel.odbcconnection.refreshonfileopen")
    }
    /// Enables or disables refresh-on-open.
    pub fn set_refresh_on_file_open(&self, enabled: bool) -> Result<(), ExcelComError> {
        put_bool(
            &self.inner,
            "excel.odbcconnection.refreshonfileopen",
            enabled,
        )
    }
    /// Returns Excel's refresh period in minutes.
    pub fn refresh_period(&self) -> Result<u32, ExcelComError> {
        u32::try_from(integer(&self.inner, "excel.odbcconnection.refreshperiod")?).map_err(|_| {
            ExcelComError::Unsupported {
                detail: "ODBC refresh period was negative",
            }
        })
    }
    /// Sets Excel's refresh period in minutes.
    pub fn set_refresh_period(&self, minutes: u32) -> Result<(), ExcelComError> {
        let value = i32::try_from(minutes).map_err(|_| ExcelComError::Unsupported {
            detail: "refresh period exceeds i32",
        })?;
        put_i32(&self.inner, "excel.odbcconnection.refreshperiod", value)
    }
    /// Returns whether refresh is enabled.
    pub fn enable_refresh(&self) -> Result<bool, ExcelComError> {
        bool_value(&self.inner, "excel.odbcconnection.enablerefresh")
    }
    /// Enables or disables refresh.
    pub fn set_enable_refresh(&self, enabled: bool) -> Result<(), ExcelComError> {
        put_bool(&self.inner, "excel.odbcconnection.enablerefresh", enabled)
    }
    /// Requests cancellation; the provider may reject an idle cancellation.
    pub fn cancel_refresh(&self) -> Result<(), ExcelComError> {
        let _ = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.odbcconnection.cancelrefresh"), false),
            vec![],
            false,
        )?;
        Ok(())
    }
}
