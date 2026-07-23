//! Redacted OLE DB connection metadata.

use std::fmt::{Debug, Formatter};

use crate::automation::{OwnedVariant, decode_variant, invoke, property_get, property_put};
use crate::excel::DispatchObject;
use crate::internal::{ComPtr, Dispatch};
use crate::object_model::{MemberId, member};
use crate::{AutomationValue, ConversionPolicy, ExcelComError};

use super::helpers::{bool_value, integer, string};
use super::{CommandType, SecretStringValue};

/// Apartment-bound Excel OLE DB connection metadata.
///
/// Connection strings are returned as [`SecretStringValue`], so diagnostics
/// and tracing do not disclose passwords or endpoints by default.
pub struct OleDbConnection {
    inner: DispatchObject,
}
impl Debug for OleDbConnection {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("OleDbConnection").field(&self.inner).finish()
    }
}
impl Clone for OleDbConnection {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl OleDbConnection {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "OLEDBConnection",
            },
        }
    }
    /// Returns the connection string in a redacting wrapper.
    pub fn connection_string(&self) -> Result<SecretStringValue, ExcelComError> {
        SecretStringValue::new(string(&self.inner, "excel.oledbconnection.connection")?)
    }
    /// Returns the provider command text, including its actual scalar or array variant shape.
    pub fn command_text(&self) -> Result<AutomationValue, ExcelComError> {
        let value = property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.oledbconnection.commandtext"), false),
            vec![],
        )?;
        decode_variant(&value, ConversionPolicy::default())
    }
    /// Returns the forward-compatible provider command type.
    pub fn command_type(&self) -> Result<CommandType, ExcelComError> {
        Ok(CommandType::from_raw(integer(
            &self.inner,
            "excel.oledbconnection.commandtype",
        )?))
    }
    /// Returns whether Excel permits a background query.
    pub fn background_query(&self) -> Result<bool, ExcelComError> {
        bool_value(&self.inner, "excel.oledbconnection.backgroundquery")
    }
    /// Enables or disables background query execution.
    pub fn set_background_query(&self, enabled: bool) -> Result<(), ExcelComError> {
        put_bool(
            &self.inner,
            "excel.oledbconnection.backgroundquery",
            enabled,
        )
    }
    /// Returns whether Excel reports an active background refresh.
    pub fn refreshing(&self) -> Result<bool, ExcelComError> {
        bool_value(&self.inner, "excel.oledbconnection.refreshing")
    }
    /// Returns whether Excel refreshes this connection when opening the workbook.
    pub fn refresh_on_file_open(&self) -> Result<bool, ExcelComError> {
        bool_value(&self.inner, "excel.oledbconnection.refreshonfileopen")
    }
    /// Enables or disables refresh on file open.
    pub fn set_refresh_on_file_open(&self, enabled: bool) -> Result<(), ExcelComError> {
        put_bool(
            &self.inner,
            "excel.oledbconnection.refreshonfileopen",
            enabled,
        )
    }
    /// Returns Excel's refresh period in minutes.
    pub fn refresh_period(&self) -> Result<u32, ExcelComError> {
        u32::try_from(integer(&self.inner, "excel.oledbconnection.refreshperiod")?).map_err(|_| {
            ExcelComError::Unsupported {
                detail: "OLE DB refresh period was negative",
            }
        })
    }
    /// Sets Excel's refresh period in minutes.
    pub fn set_refresh_period(&self, minutes: u32) -> Result<(), ExcelComError> {
        let value = i32::try_from(minutes).map_err(|_| ExcelComError::Unsupported {
            detail: "refresh period exceeds i32",
        })?;
        put_i32(&self.inner, "excel.oledbconnection.refreshperiod", value)
    }
    /// Returns whether refresh is enabled.
    pub fn enable_refresh(&self) -> Result<bool, ExcelComError> {
        bool_value(&self.inner, "excel.oledbconnection.enablerefresh")
    }
    /// Enables or disables refresh.
    pub fn set_enable_refresh(&self, enabled: bool) -> Result<(), ExcelComError> {
        put_bool(&self.inner, "excel.oledbconnection.enablerefresh", enabled)
    }
    /// Requests cancellation; Excel may reject cancellation while idle or unsupported.
    pub fn cancel_refresh(&self) -> Result<(), ExcelComError> {
        let _ = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.oledbconnection.cancelrefresh"), false),
            vec![],
            false,
        )?;
        Ok(())
    }
}

pub(crate) fn put_bool(
    target: &DispatchObject,
    id: &'static str,
    value: bool,
) -> Result<(), ExcelComError> {
    let _ = property_put(
        &target.dispatch,
        member(MemberId::new(id), true),
        OwnedVariant::bool(value),
    )?;
    Ok(())
}
pub(crate) fn put_i32(
    target: &DispatchObject,
    id: &'static str,
    value: i32,
) -> Result<(), ExcelComError> {
    let _ = property_put(
        &target.dispatch,
        member(MemberId::new(id), true),
        OwnedVariant::i32(value),
    )?;
    Ok(())
}
