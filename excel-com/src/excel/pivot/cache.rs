//! One PivotCache and PivotTable creation.

use super::helpers::{boolean, integer, put};
use super::{MissingItemsLimit, PivotSourceType, PivotTable, PivotTableCreateOptions};
use crate::SecretStringValue;
use crate::automation::{OwnedVariant, PositionalArguments, decode_variant, invoke, property_get};
use crate::excel::DispatchObject;
use crate::internal::{ComPtr, Dispatch};
use crate::object_model::{MemberId, member};
use crate::{AutomationValue, ConversionPolicy, ExcelComError};
use std::fmt::{Debug, Formatter};

/// Apartment-bound workbook data cache shared by one or more PivotTables.
pub struct PivotCache {
    inner: DispatchObject,
}
impl Debug for PivotCache {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("PivotCache").field(&self.inner).finish()
    }
}
impl Clone for PivotCache {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl PivotCache {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "PivotCache",
            },
        }
    }
    pub(crate) fn dispatch_object(&self) -> &DispatchObject {
        &self.inner
    }
    /// Returns the one-based cache index.
    pub fn index(&self) -> Result<usize, ExcelComError> {
        usize::try_from(integer(&self.inner, "excel.pivotcache.index")?).map_err(|_| {
            ExcelComError::Unsupported {
                detail: "PivotCache.Index was negative",
            }
        })
    }
    /// Returns the forward-compatible cache source type.
    pub fn source_type(&self) -> Result<PivotSourceType, ExcelComError> {
        Ok(PivotSourceType::from_raw(integer(
            &self.inner,
            "excel.pivotcache.sourcetype",
        )?))
    }
    /// Returns source data in its physical Excel Automation variant form.
    pub fn source_data(&self) -> Result<AutomationValue, ExcelComError> {
        let value = property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.pivotcache.sourcedata"), false),
            vec![],
        )?;
        decode_variant(&value, ConversionPolicy::default())
    }
    /// Returns a redacted external connection string when Excel supplies one.
    pub fn connection(&self) -> Result<Option<SecretStringValue>, ExcelComError> {
        let value = property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.pivotcache.connection"), false),
            vec![],
        )?;
        if matches!(
            value.vt(),
            windows_sys::Win32::System::Variant::VT_EMPTY
                | windows_sys::Win32::System::Variant::VT_NULL
        ) {
            return Ok(None);
        }
        Ok(Some(SecretStringValue::new(value.as_string()?)?))
    }
    /// Returns cache command text without coercing its variant form.
    pub fn command_text(&self) -> Result<AutomationValue, ExcelComError> {
        let value = property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.pivotcache.commandtext"), false),
            vec![],
        )?;
        decode_variant(&value, ConversionPolicy::default())
    }
    /// Returns whether Excel permits background refresh.
    pub fn background_query(&self) -> Result<bool, ExcelComError> {
        boolean(&self.inner, "excel.pivotcache.backgroundquery")
    }
    /// Enables or disables background refresh.
    pub fn set_background_query(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pivotcache.backgroundquery",
            OwnedVariant::bool(value),
        )
    }
    /// Returns whether Excel refreshes the cache on workbook open.
    pub fn refresh_on_file_open(&self) -> Result<bool, ExcelComError> {
        boolean(&self.inner, "excel.pivotcache.refreshonfileopen")
    }
    /// Enables or disables refresh on open.
    pub fn set_refresh_on_file_open(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pivotcache.refreshonfileopen",
            OwnedVariant::bool(value),
        )
    }
    /// Returns Excel's obsolete-item retention setting.
    pub fn missing_items_limit(&self) -> Result<MissingItemsLimit, ExcelComError> {
        Ok(MissingItemsLimit::from_raw(integer(
            &self.inner,
            "excel.pivotcache.missingitemslimit",
        )?))
    }
    /// Sets Excel's obsolete-item retention setting.
    pub fn set_missing_items_limit(&self, value: MissingItemsLimit) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pivotcache.missingitemslimit",
            OwnedVariant::i32(value.raw()),
        )
    }
    /// Requests an Excel-owned cache refresh.
    pub fn refresh(&self) -> Result<(), ExcelComError> {
        let _ = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.pivotcache.refresh"), false),
            vec![],
            false,
        )?;
        Ok(())
    }
    /// Creates a PivotTable from this cache using exact positional optional arguments.
    ///
    /// The installed type library exposes four arguments: destination, name,
    /// read-data, and `DefaultVersion`. `version` is accepted as a
    /// compatibility alias for `DefaultVersion`; when both are supplied they
    /// must agree rather than silently choosing one.
    pub fn create_pivot_table(
        &self,
        options: &PivotTableCreateOptions<'_>,
    ) -> Result<PivotTable, ExcelComError> {
        if options.name.contains('\0') {
            return Err(ExcelComError::Unsupported {
                detail: "PivotTable name cannot contain embedded NUL",
            });
        }
        let default_version = match (options.version, options.default_version) {
            (Some(version), Some(default_version)) if version != default_version => {
                return Err(ExcelComError::Unsupported {
                    detail: "PivotTable version and default_version must agree",
                });
            }
            (Some(version), _) => Some(version),
            (None, default_version) => default_version,
        };
        let mut args = PositionalArguments::new();
        args.push_object(options.destination.dispatch_object());
        args.push_required(OwnedVariant::bstr(options.name)?);
        args.push_optional(options.read_data.map(OwnedVariant::bool));
        args.push_optional(default_version.map(|v| OwnedVariant::i32(v.raw())));
        let mut value = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.pivotcache.createpivottable"), false),
            args.into_inner(),
            false,
        )?;
        Ok(PivotTable::from_dispatch(value.take_dispatch()?))
    }
}
