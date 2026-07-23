//! PivotField collection, placement, aggregation, and subtotal operations.

use super::helpers::{Iter, integer, method_iter, object, put, text};
use super::{PivotFieldOrientation, PivotFilters, PivotItems};
use crate::AggregationFunction;
use crate::automation::{
    AutomationValue, OwnedVariant, decode_variant, encode_variant, invoke, property_get,
};
use crate::excel::DispatchObject;
use crate::excel::collection::{CollectionDescriptor, count};
use crate::internal::{ComPtr, Dispatch};
use crate::object_model::{MemberId, member};
use crate::{ConversionPolicy, ExcelComError};
use std::fmt::{Debug, Formatter};

const PIVOT_FIELDS: CollectionDescriptor = CollectionDescriptor {
    name: "PivotFields",
    count: MemberId::new("excel.pivotfields.count"),
    item: MemberId::new("excel.pivotfields.item"),
    new_enum: MemberId::new("excel.pivotfields.newenum"),
};
/// Apartment-bound PivotFields collection.
pub struct PivotFields {
    inner: DispatchObject,
}
impl Debug for PivotFields {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("PivotFields").field(&self.inner).finish()
    }
}
impl Clone for PivotFields {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl PivotFields {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "PivotFields",
            },
        }
    }
    /// Returns the number of PivotFields in this projection.
    pub fn count(&self) -> Result<usize, ExcelComError> {
        count(&self.inner, PIVOT_FIELDS)
    }
    /// Returns a one-based PivotField. Excel declares `Item` as a method.
    pub fn item_by_index(&self, index: usize) -> Result<PivotField, ExcelComError> {
        if index == 0 {
            return Err(ExcelComError::Unsupported {
                detail: "PivotField indexes are one-based",
            });
        }
        self.item(OwnedVariant::i32(i32::try_from(index).map_err(|_| {
            ExcelComError::Unsupported {
                detail: "PivotField index exceeds i32",
            }
        })?))
    }
    /// Returns a PivotField by source or report name.
    pub fn item_by_name(&self, name: &str) -> Result<PivotField, ExcelComError> {
        self.item(OwnedVariant::bstr(name)?)
    }
    fn item(&self, value: OwnedVariant) -> Result<PivotField, ExcelComError> {
        let mut value = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.pivotfields.item"), false),
            vec![value],
            false,
        )?;
        Ok(PivotField::from_dispatch(value.take_dispatch()?))
    }
    /// Iterates PivotFields in method-based `_NewEnum` order.
    pub fn iter(&self) -> Result<PivotFieldsIter, ExcelComError> {
        Ok(PivotFieldsIter {
            inner: method_iter(
                &self.inner,
                "excel.pivotfields.newenum",
                "PivotFields",
                PivotField::from_dispatch,
            )?,
        })
    }
}
/// Fallible, fused iterator over [`PivotFields`].
pub struct PivotFieldsIter {
    inner: Iter<PivotField>,
}
impl Iterator for PivotFieldsIter {
    type Item = Result<PivotField, ExcelComError>;
    fn next(&mut self) -> Option<Self::Item> {
        self.inner.next()
    }
}
impl std::iter::FusedIterator for PivotFieldsIter {}
/// Apartment-bound PivotField wrapper. Moving an orientation changes which
/// PivotTable collection reports the field.
pub struct PivotField {
    inner: DispatchObject,
}
impl Debug for PivotField {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("PivotField").field(&self.inner).finish()
    }
}
impl Clone for PivotField {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl PivotField {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "PivotField",
            },
        }
    }
    pub(crate) fn dispatch_object(&self) -> &DispatchObject {
        &self.inner
    }
    /// Returns the Excel-visible field name.
    pub fn name(&self) -> Result<String, ExcelComError> {
        text(&self.inner, "excel.pivotfield.name")
    }
    /// Returns the source field name.
    pub fn source_name(&self) -> Result<String, ExcelComError> {
        text(&self.inner, "excel.pivotfield.sourcename")
    }
    /// Returns the report caption.
    pub fn caption(&self) -> Result<String, ExcelComError> {
        text(&self.inner, "excel.pivotfield.caption")
    }
    /// Changes the report caption after rejecting embedded NUL.
    pub fn set_caption(&self, value: &str) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pivotfield.caption",
            OwnedVariant::bstr(value)?,
        )
    }
    /// Returns the current forward-compatible field orientation.
    pub fn orientation(&self) -> Result<PivotFieldOrientation, ExcelComError> {
        Ok(PivotFieldOrientation::from_raw(integer(
            &self.inner,
            "excel.pivotfield.orientation",
        )?))
    }
    /// Moves the field to a row, column, page, data, or hidden orientation.
    pub fn set_orientation(&self, orientation: PivotFieldOrientation) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pivotfield.orientation",
            OwnedVariant::i32(orientation.raw()),
        )
    }
    /// Returns the one-based field position in its current orientation.
    pub fn position(&self) -> Result<usize, ExcelComError> {
        usize::try_from(integer(&self.inner, "excel.pivotfield.position")?).map_err(|_| {
            ExcelComError::Unsupported {
                detail: "PivotField.Position was negative",
            }
        })
    }
    /// Sets a nonzero one-based position in the current orientation.
    pub fn set_position(&self, position: usize) -> Result<(), ExcelComError> {
        if position == 0 {
            return Err(ExcelComError::Unsupported {
                detail: "PivotField positions are one-based",
            });
        }
        put(
            &self.inner,
            "excel.pivotfield.position",
            OwnedVariant::i32(
                i32::try_from(position).map_err(|_| ExcelComError::Unsupported {
                    detail: "PivotField position exceeds i32",
                })?,
            ),
        )
    }
    /// Returns the data-field aggregation function.
    pub fn function(&self) -> Result<AggregationFunction, ExcelComError> {
        Ok(AggregationFunction::from_raw(integer(
            &self.inner,
            "excel.pivotfield.function",
        )?))
    }
    /// Sets the data-field aggregation function.
    pub fn set_function(&self, function: AggregationFunction) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pivotfield.function",
            OwnedVariant::i32(function.raw()),
        )
    }
    /// Returns the Excel number format string.
    pub fn number_format(&self) -> Result<String, ExcelComError> {
        text(&self.inner, "excel.pivotfield.numberformat")
    }
    /// Sets the Excel number format after rejecting embedded NUL.
    pub fn set_number_format(&self, format: &str) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.pivotfield.numberformat",
            OwnedVariant::bstr(format)?,
        )
    }
    /// Returns the subtotal flags in Excel's returned Automation array order.
    pub fn subtotals(&self) -> Result<Vec<bool>, ExcelComError> {
        let value = property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.pivotfield.subtotals"), false),
            vec![],
        )?;
        let AutomationValue::Array(array) = decode_variant(&value, ConversionPolicy::default())?
        else {
            return Err(ExcelComError::Unsupported {
                detail: "PivotField.Subtotals did not return an array",
            });
        };
        array
            .values()
            .iter()
            .map(|value| match value {
                AutomationValue::Bool(value) => Ok(*value),
                _ => Err(ExcelComError::Unsupported {
                    detail: "PivotField.Subtotals contained a non-Boolean value",
                }),
            })
            .collect()
    }
    /// Sets the full subtotal flag vector as an Automation SAFEARRAY.
    pub fn set_subtotals(&self, values: &[bool]) -> Result<(), ExcelComError> {
        if values.is_empty() {
            return Err(ExcelComError::Unsupported {
                detail: "PivotField.Subtotals cannot be empty",
            });
        }
        let values = values.iter().copied().map(AutomationValue::Bool).collect();
        put(
            &self.inner,
            "excel.pivotfield.subtotals",
            encode_variant(
                &AutomationValue::Array(crate::AutomationArray::row(values)?),
                ConversionPolicy::default(),
            )?,
        )
    }
    /// Returns the field's typed PivotItems collection.
    pub fn pivot_items(&self) -> Result<PivotItems, ExcelComError> {
        object(
            &self.inner,
            "excel.pivotfield.pivotitems",
            PivotItems::from_dispatch,
        )
    }
    /// Clears all field filters using Excel's own filter semantics.
    pub fn clear_all_filters(&self) -> Result<(), ExcelComError> {
        let _ = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.pivotfield.clearallfilters"), false),
            vec![],
            false,
        )?;
        Ok(())
    }
    /// Returns this field's PivotFilters collection.
    pub fn pivot_filters(&self) -> Result<PivotFilters, ExcelComError> {
        object(
            &self.inner,
            "excel.pivotfield.pivotfilters",
            PivotFilters::from_dispatch,
        )
    }
    /// Adds a typed label filter from the supported bounded subset.
    pub fn add_label_filter(
        &self,
        options: &super::PivotLabelFilterOptions<'_>,
    ) -> Result<super::PivotFilter, ExcelComError> {
        self.pivot_filters()?.add_label(options)
    }
    /// Adds a typed value filter from the supported bounded subset.
    pub fn add_value_filter(
        &self,
        options: &super::PivotValueFilterOptions<'_>,
    ) -> Result<super::PivotFilter, ExcelComError> {
        self.pivot_filters()?.add_value(options)
    }
}
