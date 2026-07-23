//! Typed bounded PivotFilters operations.

use super::helpers::{boolean, integer, text};
use super::{PivotFilterType, PivotLabelFilterOptions, PivotValueFilterOptions, valid_filter};
use crate::automation::{
    EnumVariant, OwnedVariant, PositionalArguments, encode_variant, enumerated_dispatch, invoke,
};
use crate::excel::DispatchObject;
use crate::excel::collection::{CollectionDescriptor, count, enumerator, item_by_index};
use crate::internal::{ComPtr, Dispatch};
use crate::object_model::{MemberId, member};
use crate::{ConversionPolicy, ExcelComError};
use std::fmt::{Debug, Formatter};

const PIVOT_FILTERS: CollectionDescriptor = CollectionDescriptor {
    name: "PivotFilters",
    count: MemberId::new("excel.pivotfilters.count"),
    item: MemberId::new("excel.pivotfilters.item"),
    new_enum: MemberId::new("excel.pivotfilters.newenum"),
};
/// Apartment-bound PivotFilters collection for a PivotField.
pub struct PivotFilters {
    inner: DispatchObject,
}
impl Debug for PivotFilters {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("PivotFilters").field(&self.inner).finish()
    }
}
impl Clone for PivotFilters {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl PivotFilters {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "PivotFilters",
            },
        }
    }
    /// Returns the active and inactive filter count Excel exposes.
    pub fn count(&self) -> Result<usize, ExcelComError> {
        count(&self.inner, PIVOT_FILTERS)
    }
    /// Returns a filter at a one-based index.
    pub fn item_by_index(&self, index: usize) -> Result<PivotFilter, ExcelComError> {
        Ok(PivotFilter::from_dispatch(item_by_index(
            &self.inner,
            PIVOT_FILTERS,
            index,
        )?))
    }
    /// Iterates filters in Excel's `_NewEnum` order on the owning apartment.
    pub fn iter(&self) -> Result<PivotFiltersIter, ExcelComError> {
        Ok(PivotFiltersIter {
            enumerator: enumerator(&self.inner, PIVOT_FILTERS)?,
            index: 0,
            terminal: false,
        })
    }
    pub(crate) fn add_label(
        &self,
        options: &PivotLabelFilterOptions<'_>,
    ) -> Result<PivotFilter, ExcelComError> {
        valid_filter(options.filter_type)?;
        self.add(options.filter_type, None, options.value1, options.value2)
    }
    pub(crate) fn add_value(
        &self,
        options: &PivotValueFilterOptions<'_>,
    ) -> Result<PivotFilter, ExcelComError> {
        valid_filter(options.filter_type)?;
        self.add(
            options.filter_type,
            Some(options.data_field.dispatch_object()),
            options.value1,
            options.value2,
        )
    }
    fn add(
        &self,
        kind: PivotFilterType,
        data_field: Option<&DispatchObject>,
        value1: &crate::AutomationValue,
        value2: Option<&crate::AutomationValue>,
    ) -> Result<PivotFilter, ExcelComError> {
        let mut args = PositionalArguments::new();
        args.push_required(OwnedVariant::i32(kind.raw()));
        args.push_optional_object(data_field);
        args.push_required(encode_variant(value1, ConversionPolicy::default())?);
        args.push_optional(
            value2
                .map(|value| encode_variant(value, ConversionPolicy::default()))
                .transpose()?,
        );
        for _ in 0..4 {
            args.push_optional(None)
        }
        let mut value = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.pivotfilters.add"), false),
            args.into_inner(),
            false,
        )?;
        Ok(PivotFilter::from_dispatch(value.take_dispatch()?))
    }
}
/// Fallible, fused, apartment-bound iterator over [`PivotFilters`].
pub struct PivotFiltersIter {
    enumerator: EnumVariant,
    index: usize,
    terminal: bool,
}
impl Iterator for PivotFiltersIter {
    type Item = Result<PivotFilter, ExcelComError>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.terminal {
            return None;
        }
        match self.enumerator.next() {
            Ok(Some(mut value)) => {
                let index = self.index;
                self.index += 1;
                Some(
                    enumerated_dispatch(&mut value, "PivotFilters", index)
                        .map(PivotFilter::from_dispatch),
                )
            }
            Ok(None) => {
                self.terminal = true;
                None
            }
            Err(error) => {
                self.terminal = true;
                Some(Err(error))
            }
        }
    }
}
impl std::iter::FusedIterator for PivotFiltersIter {}
/// One typed PivotFilter created through a bounded label or value filter API.
pub struct PivotFilter {
    inner: DispatchObject,
}
impl Debug for PivotFilter {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("PivotFilter").field(&self.inner).finish()
    }
}
impl Clone for PivotFilter {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl PivotFilter {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "PivotFilter",
            },
        }
    }
    /// Returns Excel's filter name.
    pub fn name(&self) -> Result<String, ExcelComError> {
        text(&self.inner, "excel.pivotfilter.name")
    }
    /// Returns Excel's filter description.
    pub fn description(&self) -> Result<String, ExcelComError> {
        text(&self.inner, "excel.pivotfilter.description")
    }
    /// Returns whether the filter is active.
    pub fn active(&self) -> Result<bool, ExcelComError> {
        boolean(&self.inner, "excel.pivotfilter.active")
    }
    /// Returns the forward-compatible Excel filter type.
    pub fn filter_type(&self) -> Result<PivotFilterType, ExcelComError> {
        Ok(PivotFilterType::from_raw(integer(
            &self.inner,
            "excel.pivotfilter.filtertype",
        )?))
    }
    /// Deletes this filter.
    pub fn delete(self) -> Result<(), ExcelComError> {
        let _ = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.pivotfilter.delete"), false),
            vec![],
            false,
        )?;
        Ok(())
    }
}
#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn bounded_filter_rejects_unlisted_type() {
        assert!(valid_filter(PivotFilterType::from_raw(999)).is_err());
    }
}
