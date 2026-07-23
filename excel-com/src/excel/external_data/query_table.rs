//! One persistent, refreshable Excel QueryTable.

use std::fmt::{Debug, Formatter};

use crate::automation::{
    OwnedVariant, decode_variant, encode_variant, invoke, property_get, property_put,
};
use crate::excel::{DispatchObject, Range};
use crate::internal::{ComPtr, Dispatch};
use crate::object_model::{MemberId, member};
use crate::{AutomationValue, ConversionPolicy, ExcelComError};

use super::WorkbookConnection;
use super::helpers::{bool_value, optional_object, optional_range, range, string};

/// Apartment-bound persistent import. QueryTables retain refresh settings;
/// [`crate::Workbooks::open_text`] remains a one-time import operation.
pub struct QueryTable {
    inner: DispatchObject,
}
impl Debug for QueryTable {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("QueryTable").field(&self.inner).finish()
    }
}
impl Clone for QueryTable {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl QueryTable {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "QueryTable",
            },
        }
    }
    pub(crate) fn dispatch_object(&self) -> &DispatchObject {
        &self.inner
    }
    /// Returns the Excel-visible QueryTable name.
    pub fn name(&self) -> Result<String, ExcelComError> {
        string(&self.inner, "excel.querytable.name")
    }
    /// Changes the QueryTable name after rejecting embedded NUL.
    pub fn set_name(&self, value: &str) -> Result<(), ExcelComError> {
        self.put("excel.querytable.name", OwnedVariant::bstr(value)?)
    }
    /// Returns the designated destination before or after a refresh.
    pub fn destination(&self) -> Result<Range, ExcelComError> {
        range(&self.inner, "excel.querytable.destination")
    }
    /// Returns the imported result range, or `None` before Excel materializes one.
    pub fn result_range(&self) -> Result<Option<Range>, ExcelComError> {
        optional_range(&self.inner, "excel.querytable.resultrange")
    }
    /// Returns the backing workbook connection when this QueryTable has one.
    pub fn workbook_connection(&self) -> Result<Option<WorkbookConnection>, ExcelComError> {
        optional_object(
            &self.inner,
            "excel.querytable.workbookconnection",
            WorkbookConnection::from_dispatch,
        )
    }
    /// Returns command text without coercing its actual Automation variant shape.
    pub fn command_text(&self) -> Result<AutomationValue, ExcelComError> {
        let value = property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.querytable.commandtext"), false),
            vec![],
        )?;
        decode_variant(&value, ConversionPolicy::default())
    }
    /// Sets command text using the crate's checked Automation conversion layer.
    pub fn set_command_text(&self, value: &AutomationValue) -> Result<(), ExcelComError> {
        self.put(
            "excel.querytable.commandtext",
            encode_variant(value, ConversionPolicy::default())?,
        )
    }
    /// Returns whether Excel permits background refresh.
    pub fn background_query(&self) -> Result<bool, ExcelComError> {
        bool_value(&self.inner, "excel.querytable.backgroundquery")
    }
    /// Enables or disables background refresh.
    pub fn set_background_query(&self, enabled: bool) -> Result<(), ExcelComError> {
        self.put(
            "excel.querytable.backgroundquery",
            OwnedVariant::bool(enabled),
        )
    }
    /// Returns whether a background refresh is in progress.
    pub fn refreshing(&self) -> Result<bool, ExcelComError> {
        bool_value(&self.inner, "excel.querytable.refreshing")
    }
    /// Returns whether Excel refreshes this table while opening its workbook.
    pub fn refresh_on_file_open(&self) -> Result<bool, ExcelComError> {
        bool_value(&self.inner, "excel.querytable.refreshonfileopen")
    }
    /// Enables or disables refresh-on-open.
    pub fn set_refresh_on_file_open(&self, enabled: bool) -> Result<(), ExcelComError> {
        self.put(
            "excel.querytable.refreshonfileopen",
            OwnedVariant::bool(enabled),
        )
    }
    /// Returns whether Excel preserves cell formatting across a refresh.
    pub fn preserve_formatting(&self) -> Result<bool, ExcelComError> {
        bool_value(&self.inner, "excel.querytable.preserveformatting")
    }
    /// Enables or disables format preservation.
    pub fn set_preserve_formatting(&self, enabled: bool) -> Result<(), ExcelComError> {
        self.put(
            "excel.querytable.preserveformatting",
            OwnedVariant::bool(enabled),
        )
    }
    /// Returns whether Excel adjusts result columns after a refresh.
    pub fn adjust_column_width(&self) -> Result<bool, ExcelComError> {
        bool_value(&self.inner, "excel.querytable.adjustcolumnwidth")
    }
    /// Enables or disables result-column width adjustment.
    pub fn set_adjust_column_width(&self, enabled: bool) -> Result<(), ExcelComError> {
        self.put(
            "excel.querytable.adjustcolumnwidth",
            OwnedVariant::bool(enabled),
        )
    }
    /// Refreshes and preserves Excel's Boolean success result.
    pub fn refresh(&self, background: Option<bool>) -> Result<bool, ExcelComError> {
        let mut arguments = crate::automation::PositionalArguments::new();
        arguments.push_optional(background.map(OwnedVariant::bool));
        let value = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.querytable.refresh"), false),
            arguments.into_inner(),
            false,
        )?;
        value.as_bool().ok_or(ExcelComError::Unsupported {
            detail: "QueryTable.Refresh did not return VT_BOOL",
        })
    }
    /// Requests cancellation; an idle or unsupported QueryTable may return Excel's error.
    pub fn cancel_refresh(&self) -> Result<(), ExcelComError> {
        let _ = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.querytable.cancelrefresh"), false),
            vec![],
            false,
        )?;
        Ok(())
    }
    /// Deletes this persistent import and consumes its invalidated wrapper.
    pub fn delete(self) -> Result<(), ExcelComError> {
        let _ = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.querytable.delete"), false),
            vec![],
            false,
        )?;
        Ok(())
    }
    fn put(&self, id: &'static str, value: OwnedVariant) -> Result<(), ExcelComError> {
        let _ = property_put(&self.inner.dispatch, member(MemberId::new(id), true), value)?;
        Ok(())
    }
}
