//! Stateful AutoFilter wrappers and typed filter criteria.

use std::fmt::{Debug, Formatter};
use std::iter::FusedIterator;

use windows_sys::Win32::System::Com::SAFEARRAYBOUND;
use windows_sys::Win32::System::Variant::{VT_EMPTY, VT_NULL};

use crate::ExcelComError;
use crate::automation::{
    AutomationValue, ConversionPolicy, EnumVariant, OwnedVariant, PositionalArguments, SafeArray,
    decode_variant, encode_variant, enumerated_dispatch, invoke, property_get,
};
use crate::excel::collection::{
    CollectionDescriptor, count as collection_count, enumerator, item_by_index,
};
use crate::excel::table::{bool_get, object_get, one_based, range_get};
use crate::excel::{DispatchObject, Range, Worksheet};
use crate::internal::{ComPtr, Dispatch};
use crate::object_model::{MemberId, member};

const FILTERS: CollectionDescriptor = CollectionDescriptor {
    name: "Filters",
    count: MemberId::new("excel.filters.count"),
    item: MemberId::new("excel.filters.item"),
    new_enum: MemberId::new("excel.filters.newenum"),
};

macro_rules! raw_filter_type {
    ($(#[$docs:meta])* $name:ident { $($constant:ident = $value:expr => $constant_docs:literal;)* }) => {
        $(#[$docs])*
        #[derive(Clone, Copy, Debug, Eq, PartialEq, Hash)]
        #[repr(transparent)]
        pub struct $name(i32);
        impl $name {
            $(#[doc = $constant_docs] pub const $constant: Self = Self($value);)*
            /// Creates this value from an Excel type-library integer.
            pub const fn from_raw(value: i32) -> Self { Self(value) }
            /// Returns the raw Excel integer, preserving unknown values.
            pub const fn raw(self) -> i32 { self.0 }
        }
    };
}

raw_filter_type! {
    /// A forward-compatible `XlAutoFilterOperator` value.
    AutoFilterOperator {
        AND = 1 => "`xlAnd`.";
        OR = 2 => "`xlOr`.";
        TOP_ITEMS = 3 => "`xlTop10Items`.";
        BOTTOM_ITEMS = 4 => "`xlBottom10Items`.";
        TOP_PERCENT = 5 => "`xlTop10Percent`.";
        BOTTOM_PERCENT = 6 => "`xlBottom10Percent`.";
        VALUES = 7 => "`xlFilterValues`.";
        DYNAMIC = 11 => "`xlFilterDynamic`.";
    }
}

raw_filter_type! {
    /// A forward-compatible `XlDynamicFilterCriteria` value.
    DynamicFilterCriteria {
        TODAY = 1 => "`xlFilterToday`.";
        YESTERDAY = 2 => "`xlFilterYesterday`.";
        TOMORROW = 3 => "`xlFilterTomorrow`.";
        THIS_WEEK = 4 => "`xlFilterThisWeek`.";
        LAST_WEEK = 5 => "`xlFilterLastWeek`.";
        NEXT_WEEK = 6 => "`xlFilterNextWeek`.";
        THIS_MONTH = 7 => "`xlFilterThisMonth`.";
        LAST_MONTH = 8 => "`xlFilterLastMonth`.";
        NEXT_MONTH = 9 => "`xlFilterNextMonth`.";
        THIS_QUARTER = 10 => "`xlFilterThisQuarter`.";
        THIS_YEAR = 13 => "`xlFilterThisYear`.";
        YEAR_TO_DATE = 16 => "`xlFilterYearToDate`.";
        ABOVE_AVERAGE = 33 => "`xlFilterAboveAverage`.";
        BELOW_AVERAGE = 34 => "`xlFilterBelowAverage`.";
    }
}

/// One bounded criterion accepted by [`Range::apply_auto_filter`].
///
/// Values are carried as semantic Automation values.  The `Values` form is
/// encoded as a one-dimensional `SAFEARRAY(VARIANT)` because that is the
/// Automation shape Excel uses for `xlFilterValues`; nested arrays remain
/// rejected before COM.
#[derive(Clone, Debug)]
pub enum FilterCriterion {
    /// A scalar equality/comparison criterion interpreted by Excel.
    Value(AutomationValue),
    /// A vector of selected values used with `xlFilterValues`.
    Values(Vec<AutomationValue>),
    /// A dynamic date or aggregate criterion used with `xlFilterDynamic`.
    Dynamic(DynamicFilterCriteria),
    /// A top-item count used with `xlTop10Items`.
    TopItems(i32),
    /// A top percentage used with `xlTop10Percent`.
    TopPercent(i32),
    /// A bottom-item count used with `xlBottom10Items`.
    BottomItems(i32),
    /// A bottom percentage used with `xlBottom10Percent`.
    BottomPercent(i32),
}

/// Positional input for Excel `Range.AutoFilter`.
///
/// `field` is one-based relative to the receiver Range (or the containing
/// table). Filtering is Excel state: it changes row visibility and can leave
/// filter arrows in place after criteria are cleared.
#[derive(Clone, Debug)]
pub struct AutoFilterOptions {
    /// One-based field relative to the filtered Range.
    pub field: usize,
    /// The first criterion, or `None` for Excel's omitted position.
    pub criterion1: Option<FilterCriterion>,
    /// An explicit operator; specialized criteria infer a matching operator when absent.
    pub operator: Option<AutoFilterOperator>,
    /// The optional second criterion.
    pub criterion2: Option<FilterCriterion>,
    /// Whether Excel should show this field's filter dropdown.
    pub visible_dropdown: Option<bool>,
}

/// Apartment-bound wrapper for one Excel AutoFilter state object.
pub struct AutoFilter {
    inner: DispatchObject,
}
impl Debug for AutoFilter {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("AutoFilter").field(&self.inner).finish()
    }
}
impl Clone for AutoFilter {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl AutoFilter {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "AutoFilter",
            },
        }
    }
    /// Returns the Range covered by this stateful AutoFilter.
    pub fn range(&self) -> Result<Range, ExcelComError> {
        range_get(&self.inner, "excel.autofilter.range")
    }
    /// Returns the typed per-field Filters collection.
    pub fn filters(&self) -> Result<Filters, ExcelComError> {
        object_get(
            &self.inner,
            "excel.autofilter.filters",
            Filters::from_dispatch,
        )
    }
}

/// Apartment-bound collection of per-field Excel filters.
pub struct Filters {
    inner: DispatchObject,
}
impl Debug for Filters {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("Filters").field(&self.inner).finish()
    }
}
impl Clone for Filters {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl Filters {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "Filters",
            },
        }
    }
    /// Returns the number of filter fields.
    pub fn count(&self) -> Result<usize, ExcelComError> {
        collection_count(&self.inner, FILTERS)
    }
    /// Returns the one-based filter for a field.
    pub fn item(&self, index: usize) -> Result<Filter, ExcelComError> {
        Ok(Filter::from_dispatch(item_by_index(
            &self.inner,
            FILTERS,
            index,
        )?))
    }
    /// Iterates fields in Excel's `_NewEnum` order.
    pub fn iter(&self) -> Result<FiltersIter, ExcelComError> {
        Ok(FiltersIter {
            enumerator: enumerator(&self.inner, FILTERS)?,
            next_index: 0,
            terminal: false,
        })
    }
}

/// Fallible, single-pass, apartment-bound iterator over [`Filters`].
pub struct FiltersIter {
    enumerator: EnumVariant,
    next_index: usize,
    terminal: bool,
}
impl Iterator for FiltersIter {
    type Item = Result<Filter, ExcelComError>;
    fn next(&mut self) -> Option<Self::Item> {
        if self.terminal {
            return None;
        }
        match self.enumerator.next() {
            Ok(Some(mut value)) => {
                let index = self.next_index;
                self.next_index += 1;
                Some(enumerated_dispatch(&mut value, "Filters", index).map(Filter::from_dispatch))
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
impl FusedIterator for FiltersIter {}

/// Apartment-bound read-only state for a single filtered field.
pub struct Filter {
    inner: DispatchObject,
}
impl Debug for Filter {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("Filter").field(&self.inner).finish()
    }
}
impl Clone for Filter {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl Filter {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "Filter",
            },
        }
    }
    /// Returns whether this field currently has criteria applied.
    pub fn is_on(&self) -> Result<bool, ExcelComError> {
        bool_get(&self.inner, "excel.filter.on")
    }
    /// Returns the field operator, preserving unknown Excel values.
    pub fn operator(&self) -> Result<AutoFilterOperator, ExcelComError> {
        let value = property_get(
            &self.inner.dispatch,
            member(MemberId::new("excel.filter.operator-2641"), false),
            vec![],
        )?
        .as_i32()
        .ok_or(ExcelComError::Unsupported {
            detail: "Filter.Operator did not return VT_I4",
        })?;
        Ok(AutoFilterOperator::from_raw(value))
    }
    /// Returns the first raw Automation criterion, or `None` when Excel has no criterion.
    pub fn criteria1(&self) -> Result<Option<AutomationValue>, ExcelComError> {
        criterion_get(&self.inner, "excel.filter.criteria1")
    }
    /// Returns the second raw Automation criterion, or `None` when Excel has no criterion.
    pub fn criteria2(&self) -> Result<Option<AutomationValue>, ExcelComError> {
        criterion_get(&self.inner, "excel.filter.criteria2")
    }
}

impl Range {
    /// Applies Excel AutoFilter criteria to this Range.
    ///
    /// Specialized criteria infer their matching `XlAutoFilterOperator` only
    /// when `options.operator` is absent. Excel validates criterion syntax,
    /// filter field containment, and state interactions.
    pub fn apply_auto_filter(&self, options: &AutoFilterOptions) -> Result<(), ExcelComError> {
        let mut args = PositionalArguments::new();
        args.push_optional(Some(OwnedVariant::i32(one_based(
            options.field,
            "Range.AutoFilter field",
        )?)));
        push_criterion(&mut args, options.criterion1.as_ref())?;
        args.push_optional(
            options
                .operator
                .or_else(|| options.criterion1.as_ref().and_then(inferred_operator))
                .map(|value| OwnedVariant::i32(value.raw())),
        );
        push_criterion(&mut args, options.criterion2.as_ref())?;
        args.push_optional(None); // SubField is intentionally outside this bounded wrapper.
        args.push_optional(options.visible_dropdown.map(OwnedVariant::bool));
        let _ = invoke(
            &self.dispatch_object().dispatch,
            member(MemberId::new("excel.range.autofilter-3289"), false),
            args.into_inner(),
            false,
        )?;
        Ok(())
    }

    /// Clears active criteria through the containing worksheet's `ShowAllData`.
    ///
    /// This deliberately does not toggle AutoFilter arrows. Excel returns its
    /// own invocation failure when there is no active filter state to clear.
    pub fn clear_auto_filter(&self) -> Result<(), ExcelComError> {
        let mut value = property_get(
            &self.dispatch_object().dispatch,
            member(MemberId::new("excel.range.worksheet"), false),
            vec![],
        )?;
        let sheet = Worksheet::from_dispatch(value.take_dispatch()?);
        sheet.show_all_data()
    }
}

impl Worksheet {
    /// Returns this worksheet's AutoFilter state, or `None` when no filter object exists.
    pub fn auto_filter(&self) -> Result<Option<AutoFilter>, ExcelComError> {
        let mut value = property_get(
            &self.dispatch_object().dispatch,
            member(MemberId::new("excel.worksheet.autofilter-3289"), false),
            vec![],
        )?;
        value
            .take_optional_dispatch()
            .map(|value| value.map(AutoFilter::from_dispatch))
    }
    /// Returns whether AutoFilter arrows are displayed on this worksheet.
    pub fn auto_filter_mode(&self) -> Result<bool, ExcelComError> {
        bool_get(self.dispatch_object(), "excel.worksheet.autofiltermode")
    }
    /// Returns whether the worksheet currently has a filtered-out row.
    pub fn filter_mode(&self) -> Result<bool, ExcelComError> {
        bool_get(self.dispatch_object(), "excel.worksheet.filtermode")
    }
    /// Clears active worksheet filtering while retaining AutoFilter arrows.
    pub fn show_all_data(&self) -> Result<(), ExcelComError> {
        let _ = invoke(
            &self.dispatch_object().dispatch,
            member(MemberId::new("excel.worksheet.showalldata"), false),
            vec![],
            false,
        )?;
        Ok(())
    }
}

fn criterion_get(
    target: &DispatchObject,
    id: &'static str,
) -> Result<Option<AutomationValue>, ExcelComError> {
    let value = property_get(&target.dispatch, member(MemberId::new(id), false), vec![])?;
    if matches!(value.vt(), VT_EMPTY | VT_NULL) {
        return Ok(None);
    }
    decode_variant(&value, ConversionPolicy::default()).map(Some)
}

fn inferred_operator(value: &FilterCriterion) -> Option<AutoFilterOperator> {
    match value {
        FilterCriterion::Values(_) => Some(AutoFilterOperator::VALUES),
        FilterCriterion::Dynamic(_) => Some(AutoFilterOperator::DYNAMIC),
        FilterCriterion::TopItems(_) => Some(AutoFilterOperator::TOP_ITEMS),
        FilterCriterion::TopPercent(_) => Some(AutoFilterOperator::TOP_PERCENT),
        FilterCriterion::BottomItems(_) => Some(AutoFilterOperator::BOTTOM_ITEMS),
        FilterCriterion::BottomPercent(_) => Some(AutoFilterOperator::BOTTOM_PERCENT),
        FilterCriterion::Value(_) => None,
    }
}

fn push_criterion(
    arguments: &mut PositionalArguments,
    criterion: Option<&FilterCriterion>,
) -> Result<(), ExcelComError> {
    match criterion {
        None => arguments.push_optional(None),
        Some(value) => arguments.push_result(encode_criterion(value))?,
    }
    Ok(())
}

fn encode_criterion(criterion: &FilterCriterion) -> Result<OwnedVariant, ExcelComError> {
    match criterion {
        FilterCriterion::Value(value) => encode_variant(value, ConversionPolicy::default()),
        FilterCriterion::Values(values) => encode_filter_values(values),
        FilterCriterion::Dynamic(value) => Ok(OwnedVariant::i32(value.raw())),
        FilterCriterion::TopItems(value)
        | FilterCriterion::TopPercent(value)
        | FilterCriterion::BottomItems(value)
        | FilterCriterion::BottomPercent(value) => Ok(OwnedVariant::i32(*value)),
    }
}

fn encode_filter_values(values: &[AutomationValue]) -> Result<OwnedVariant, ExcelComError> {
    if values.is_empty() {
        return Err(ExcelComError::Unsupported {
            detail: "xlFilterValues requires at least one value",
        });
    }
    let count = u32::try_from(values.len()).map_err(|_| ExcelComError::Unsupported {
        detail: "too many xlFilterValues criteria",
    })?;
    let array = SafeArray::create_variant(&[SAFEARRAYBOUND {
        cElements: count,
        lLbound: 0,
    }])
    .ok_or(ExcelComError::Unsupported {
        detail: "could not allocate xlFilterValues SAFEARRAY",
    })?;
    for (index, value) in values.iter().enumerate() {
        if matches!(value, AutomationValue::Array(_)) {
            return Err(ExcelComError::Unsupported {
                detail: "nested arrays are not AutoFilter values",
            });
        }
        let encoded = encode_variant(value, ConversionPolicy::default())?;
        let index = i32::try_from(index).map_err(|_| ExcelComError::Unsupported {
            detail: "too many xlFilterValues criteria",
        })?;
        if !array.put_variant(&[index], &encoded) {
            return Err(ExcelComError::Unsupported {
                detail: "could not populate xlFilterValues SAFEARRAY",
            });
        }
    }
    Ok(OwnedVariant::array(array))
}

#[cfg(test)]
mod tests {
    use super::*;
    use windows_sys::Win32::Foundation::DISP_E_PARAMNOTFOUND;

    #[test]
    fn auto_filter_positions_preserve_subfield_gap() {
        let mut args = PositionalArguments::new();
        args.push_optional(Some(OwnedVariant::i32(1)));
        push_criterion(
            &mut args,
            Some(&FilterCriterion::Value(AutomationValue::Text(
                "x".to_owned(),
            ))),
        )
        .expect("criterion");
        args.push_optional(Some(OwnedVariant::i32(AutoFilterOperator::AND.raw())));
        push_criterion(&mut args, None).expect("criterion");
        args.push_optional(None);
        args.push_optional(Some(OwnedVariant::bool(false)));
        let values = args.into_inner();
        assert_eq!(values.len(), 6);
        assert_eq!(values[4].as_scode(), Some(DISP_E_PARAMNOTFOUND));
        assert_eq!(values[5].as_bool(), Some(false));
    }
}
