use std::fmt::{Debug, Formatter};

use crate::ExcelComError;
use crate::automation::{
    AutomationValue, ConversionPolicy, DateWriteMode, OwnedVariant, PositionalArguments,
    decode_variant, invoke, property_get, property_put, validate_range_shape,
};
use crate::excel::{Areas, DispatchObject, RangeAddressOptions, ReferenceStyle};
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
    pub(crate) fn dispatch_object(&self) -> &DispatchObject {
        &self.inner
    }

    /// Returns whether two Range wrappers denote the same COM object identity.
    pub fn is_same_object(&self, other: &Self) -> Result<bool, ExcelComError> {
        self.inner.same_object(&other.inner)
    }

    /// Returns the default absolute A1-style address as reported by Excel.
    ///
    /// This compatibility convenience is equivalent to [`Self::address_a1`].
    pub fn address(&self) -> Result<String, ExcelComError> {
        self.address_a1()
    }

    /// Returns Excel's absolute A1-style address with `External = false`.
    pub fn address_a1(&self) -> Result<String, ExcelComError> {
        self.address_with_options(&RangeAddressOptions::default())
    }

    /// Returns Excel's absolute R1C1-style address with `External = false`.
    ///
    /// ```no_run
    /// # fn example(range: &excel_com::Range) -> Result<(), excel_com::ExcelComError> {
    /// let r1c1 = range.address_r1c1()?;
    /// # assert!(!r1c1.is_empty());
    /// # Ok(())
    /// # }
    /// ```
    pub fn address_r1c1(&self) -> Result<String, ExcelComError> {
        self.address_with_options(&RangeAddressOptions {
            reference_style: ReferenceStyle::R1C1,
            ..RangeAddressOptions::default()
        })
    }

    /// Returns an address using explicit Excel `Range.Address` options.
    ///
    /// The options map directly to Excel's `RowAbsolute`, `ColumnAbsolute`,
    /// `ReferenceStyle`, `External`, and `RelativeTo` positions. Omitted
    /// booleans and bases are preserved as `Missing`; relative R1C1 offsets
    /// are calculated by Excel from `relative_to`, never by Rust arithmetic.
    pub fn address_with_options(
        &self,
        options: &RangeAddressOptions<'_>,
    ) -> Result<String, ExcelComError> {
        property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.range.address"), false),
            address_arguments(options),
        )?
        .as_string()
    }

    /// Returns Excel's external/qualified address in the requested notation.
    ///
    /// Unsaved workbooks commonly use a temporary workbook name rather than a
    /// filesystem path, and Excel controls all quoting. This textual context
    /// is not a COM identity claim.
    pub fn external_address(&self, style: ReferenceStyle) -> Result<String, ExcelComError> {
        self.address_with_options(&RangeAddressOptions {
            reference_style: style,
            external: Some(true),
            ..RangeAddressOptions::default()
        })
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

    /// Returns the Range containing this range's cells.
    pub fn cells(&self) -> Result<Range, ExcelComError> {
        self.range_property("excel.range.cells")
    }

    /// Returns the one-based cell addressed relative to this Range.
    pub fn item(&self, row: usize, column: Option<usize>) -> Result<Range, ExcelComError> {
        let row = one_based_i32(row, "Range.Item row")?;
        let mut arguments = vec![crate::automation::OwnedVariant::i32(row)];
        if let Some(column) = column {
            arguments.push(crate::automation::OwnedVariant::i32(one_based_i32(
                column,
                "Range.Item column",
            )?));
        }
        let mut result = property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.range.item"), false),
            arguments,
        )?;
        Ok(Self::from_dispatch(result.take_dispatch()?))
    }

    /// Returns the one-based cell at `row`, `column` relative to this Range.
    pub fn cell(&self, row: usize, column: usize) -> Result<Range, ExcelComError> {
        self.item(row, Some(column))
    }

    /// Returns a Range translated by signed row and column offsets.
    pub fn offset(&self, row_offset: isize, column_offset: isize) -> Result<Range, ExcelComError> {
        let row = i32::try_from(row_offset).map_err(|_| ExcelComError::Unsupported {
            detail: "Range.Offset row offset exceeds i32",
        })?;
        let column = i32::try_from(column_offset).map_err(|_| ExcelComError::Unsupported {
            detail: "Range.Offset column offset exceeds i32",
        })?;
        self.range_property_with(
            "excel.range.offset",
            vec![
                crate::automation::OwnedVariant::i32(row),
                crate::automation::OwnedVariant::i32(column),
            ],
        )
    }

    /// Returns this Range resized to the nonzero `rows` by `columns` dimensions.
    pub fn resize(&self, rows: usize, columns: usize) -> Result<Range, ExcelComError> {
        self.range_property_with(
            "excel.range.resize",
            vec![
                crate::automation::OwnedVariant::i32(one_based_i32(rows, "Range.Resize rows")?),
                crate::automation::OwnedVariant::i32(one_based_i32(
                    columns,
                    "Range.Resize columns",
                )?),
            ],
        )
    }

    /// Returns the Range representing this Range's rows.
    pub fn rows(&self) -> Result<Range, ExcelComError> {
        self.range_property("excel.range.rows")
    }

    /// Returns the Range representing this Range's columns.
    pub fn columns(&self) -> Result<Range, ExcelComError> {
        self.range_property("excel.range.columns")
    }

    /// Returns the collection of contiguous areas in this Range.
    pub fn areas(&self) -> Result<Areas, ExcelComError> {
        let mut result = property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.range.areas"), false),
            vec![],
        )?;
        Ok(Areas::from_dispatch(result.take_dispatch()?))
    }

    /// Returns the full worksheet rows intersecting this Range.
    pub fn entire_row(&self) -> Result<Range, ExcelComError> {
        self.range_property("excel.range.entirerow")
    }

    /// Returns the full worksheet columns intersecting this Range.
    pub fn entire_column(&self) -> Result<Range, ExcelComError> {
        self.range_property("excel.range.entirecolumn")
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
        let related = self.range_property(id)?;
        related.cell_count()
    }

    fn range_property(&self, id: &'static str) -> Result<Range, ExcelComError> {
        self.range_property_with(id, vec![])
    }
    fn range_property_with(
        &self,
        id: &'static str,
        arguments: Vec<crate::automation::OwnedVariant>,
    ) -> Result<Range, ExcelComError> {
        let mut result = property_get(
            &self.inner.dispatch,
            member(MemberId::new(id), false),
            arguments,
        )?;
        Ok(Self::from_dispatch(result.take_dispatch()?))
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

fn address_arguments(options: &RangeAddressOptions<'_>) -> Vec<OwnedVariant> {
    let mut arguments = PositionalArguments::new();
    arguments.push_optional(options.row_absolute.map(OwnedVariant::bool));
    arguments.push_optional(options.column_absolute.map(OwnedVariant::bool));
    arguments.push_required(OwnedVariant::i32(options.reference_style.raw()));
    arguments.push_optional(options.external.map(OwnedVariant::bool));
    arguments.push_optional_object(options.relative_to.map(Range::dispatch_object));
    arguments.into_inner()
}

fn one_based_i32(value: usize, detail: &'static str) -> Result<i32, ExcelComError> {
    if value == 0 {
        return Err(ExcelComError::Unsupported {
            detail: "Range indices and dimensions are one-based and nonzero",
        });
    }
    i32::try_from(value).map_err(|_| ExcelComError::Unsupported { detail })
}

#[cfg(test)]
mod tests {
    use super::*;
    use windows_sys::Win32::Foundation::DISP_E_PARAMNOTFOUND;

    #[test]
    fn address_arguments_are_logical_and_position_preserving() {
        let values = address_arguments(&RangeAddressOptions::default());
        assert_eq!(values.len(), 5);
        assert_eq!(values[0].as_bool(), Some(true));
        assert_eq!(values[1].as_bool(), Some(true));
        assert_eq!(values[2].as_i32(), Some(1));
        assert_eq!(values[3].as_bool(), Some(false));
        assert_eq!(values[4].as_scode(), Some(DISP_E_PARAMNOTFOUND));
    }

    #[test]
    fn address_arguments_keep_interior_missing_values() {
        let values = address_arguments(&RangeAddressOptions {
            row_absolute: None,
            column_absolute: Some(false),
            reference_style: ReferenceStyle::R1C1,
            external: Some(true),
            relative_to: None,
        });
        assert_eq!(values[0].as_scode(), Some(DISP_E_PARAMNOTFOUND));
        assert_eq!(values[1].as_bool(), Some(false));
        assert_eq!(values[2].as_i32(), Some(-4150));
        assert_eq!(values[3].as_bool(), Some(true));
        assert_eq!(values[4].as_scode(), Some(DISP_E_PARAMNOTFOUND));
    }
}
