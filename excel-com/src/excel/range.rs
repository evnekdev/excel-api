use std::fmt::{Debug, Formatter};

use crate::ExcelComError;
use crate::automation::{
    AutomationValue, ConversionPolicy, DateWriteMode, decode_variant, invoke, property_get,
    property_put, validate_range_shape,
};
use crate::excel::DispatchObject;
use crate::internal::{ComPtr, Dispatch};
use crate::object_model::{MemberId, member};

/// Experimental wrapper for an Excel `Range`.
pub struct Range {
    inner: DispatchObject,
}

impl Debug for Range {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter.debug_tuple("Range").field(&self.inner).finish()
    }
}

impl Clone for Range {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}

impl Range {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "Range",
            },
        }
    }

    /// Returns the default A1-style address as reported by Excel.
    pub fn address(&self) -> Result<String, ExcelComError> {
        property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.range.address"), false),
            vec![],
        )?
        .as_string()
    }

    /// Returns the one-based first-row position.
    pub fn row(&self) -> Result<i32, ExcelComError> {
        self.i32_property("excel.range.row", "Range.Row")
    }

    /// Returns the one-based first-column position.
    pub fn column(&self) -> Result<i32, ExcelComError> {
        self.i32_property("excel.range.column", "Range.Column")
    }

    /// Returns the number of cells in the Range.
    pub fn cell_count(&self) -> Result<i32, ExcelComError> {
        self.i32_property("excel.range.count", "Range.Count")
    }

    /// Returns the number of rows in the Range.
    pub fn row_count(&self) -> Result<i32, ExcelComError> {
        self.related_count("excel.range.rows")
    }

    /// Returns the number of columns in the Range.
    pub fn column_count(&self) -> Result<i32, ExcelComError> {
        self.related_count("excel.range.columns")
    }

    /// Gets `Range.Value`, preserving scalar and rectangular Automation values.
    pub fn value(&self) -> Result<AutomationValue, ExcelComError> {
        self.value_get("excel.range.value")
    }

    /// Sets `Range.Value` after exact shape validation, before any COM setter call.
    pub fn set_value(&self, value: AutomationValue) -> Result<(), ExcelComError> {
        self.value_put("excel.range.value", value, DateWriteMode::Value)
    }

    /// Gets `Range.Value2`, which represents date serials as numbers.
    pub fn value2(&self) -> Result<AutomationValue, ExcelComError> {
        self.value_get("excel.range.value2")
    }

    /// Sets `Range.Value2` after exact shape validation, writing dates as serial numbers.
    pub fn set_value2(&self, value: AutomationValue) -> Result<(), ExcelComError> {
        self.value_put("excel.range.value2", value, DateWriteMode::Value2)
    }

    /// Gets `Range.Formula` as a semantic Automation value.
    pub fn formula(&self) -> Result<AutomationValue, ExcelComError> {
        self.value_get("excel.range.formula")
    }

    /// Sets `Range.Formula` after exact shape validation.
    pub fn set_formula(&self, value: AutomationValue) -> Result<(), ExcelComError> {
        self.value_put("excel.range.formula", value, DateWriteMode::Value)
    }

    /// Gets `Range.Formula2` as a semantic Automation value.
    pub fn formula2(&self) -> Result<AutomationValue, ExcelComError> {
        self.value_get("excel.range.formula2")
    }

    /// Sets `Range.Formula2` after exact shape validation.
    pub fn set_formula2(&self, value: AutomationValue) -> Result<(), ExcelComError> {
        self.value_put("excel.range.formula2", value, DateWriteMode::Value)
    }

    /// Removes formulas and values while leaving formatting untouched.
    pub fn clear_contents(&self) -> Result<(), ExcelComError> {
        let _ = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.range.clearcontents"), false),
            vec![],
            false,
        )?;
        Ok(())
    }

    fn i32_property(&self, id: &'static str, name: &'static str) -> Result<i32, ExcelComError> {
        property_get(
            &self.inner.dispatch,
            member(MemberId::new(id), false),
            vec![],
        )?
        .as_i32()
        .ok_or(ExcelComError::Unsupported { detail: name })
    }

    fn related_count(&self, id: &'static str) -> Result<i32, ExcelComError> {
        let mut result = property_get(
            &self.inner.dispatch,
            member(MemberId::new(id), false),
            vec![],
        )?;
        let related = Self::from_dispatch(result.take_dispatch()?);
        related.cell_count()
    }

    fn dimensions(&self) -> Result<(usize, usize), ExcelComError> {
        let rows = usize::try_from(self.row_count()?).map_err(|_| ExcelComError::Unsupported {
            detail: "negative Range row count",
        })?;
        let columns =
            usize::try_from(self.column_count()?).map_err(|_| ExcelComError::Unsupported {
                detail: "negative Range column count",
            })?;
        Ok((rows, columns))
    }

    fn value_get(&self, id: &'static str) -> Result<AutomationValue, ExcelComError> {
        let result = property_get(
            &self.inner.dispatch,
            member(MemberId::new(id), false),
            vec![],
        )?;
        decode_variant(&result, ConversionPolicy::default())
    }

    fn value_put(
        &self,
        id: &'static str,
        value: AutomationValue,
        date_write: DateWriteMode,
    ) -> Result<(), ExcelComError> {
        let (rows, columns) = self.dimensions()?;
        validate_range_shape(&value, rows, columns)?;
        let encoded = crate::automation::encode_variant(
            &value,
            ConversionPolicy {
                date_write,
                ..ConversionPolicy::default()
            },
        )?;
        let _ = property_put(
            &self.inner.dispatch,
            member(MemberId::new(id), true),
            encoded,
        )?;
        Ok(())
    }
}
