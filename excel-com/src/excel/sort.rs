//! Range and persistent Excel Sort wrappers.

use std::fmt::{Debug, Formatter};

use crate::ExcelComError;
use crate::automation::{OwnedVariant, PositionalArguments, invoke};
use crate::excel::collection::{CollectionDescriptor, count as collection_count};
use crate::excel::table::{bool_get, bool_put, i32_get, i32_put, object_get};
use crate::excel::{DispatchObject, Range, TableHeaderMode};
use crate::internal::{ComPtr, Dispatch};
use crate::object_model::{MemberId, member};

const SORT_FIELDS: CollectionDescriptor = CollectionDescriptor {
    name: "SortFields",
    count: MemberId::new("excel.sortfields.count"),
    item: MemberId::new("excel.sortfields.item"),
    new_enum: MemberId::new("excel.sortfields.newenum"),
};

macro_rules! raw_sort_type {
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

raw_sort_type! {
    /// A forward-compatible `XlSortOrder` value.
    SortOrder {
        ASCENDING = 1 => "`xlAscending`.";
        DESCENDING = 2 => "`xlDescending`.";
    }
}
raw_sort_type! {
    /// A forward-compatible `XlSortOrientation` value.
    SortOrientation {
        COLUMNS = 1 => "`xlSortColumns`.";
        ROWS = 2 => "`xlSortRows` (top-to-bottom), Excel's normal orientation.";
    }
}
raw_sort_type! {
    /// A forward-compatible `XlSortMethod` value.
    SortMethod {
        PINYIN = 1 => "`xlPinYin`.";
        STROKE = 2 => "`xlStroke`.";
    }
}
raw_sort_type! {
    /// A forward-compatible `XlSortDataOption` value.
    SortDataOption {
        NORMAL = 0 => "`xlSortNormal`.";
        TEXT_AS_NUMBERS = 1 => "`xlSortTextAsNumbers`.";
    }
}

/// Typed positional input for simple Excel `Range.Sort`.
///
/// Sorting changes the receiver's worksheet cells in place. Key ranges are
/// passed as dispatch objects; Excel, not Rust identity comparison, validates
/// whether each key belongs to a compatible sort range.
#[derive(Debug)]
pub struct RangeSortOptions<'a> {
    /// Required first key Range.
    pub key1: &'a Range,
    /// Sort order for the first key.
    pub order1: SortOrder,
    /// Optional second key Range.
    pub key2: Option<&'a Range>,
    /// Optional second-key order; it requires `key2`.
    pub order2: Option<SortOrder>,
    /// Optional third key Range.
    pub key3: Option<&'a Range>,
    /// Optional third-key order; it requires `key3`.
    pub order3: Option<SortOrder>,
    /// Header declaration sent to Excel.
    pub header: TableHeaderMode,
    /// Optional orientation.
    pub orientation: Option<SortOrientation>,
    /// Optional case-sensitive comparison setting.
    pub match_case: Option<bool>,
    /// Optional sorting method.
    pub sort_method: Option<SortMethod>,
    /// Optional first-key data treatment.
    pub data_option1: Option<SortDataOption>,
    /// Optional second-key data treatment.
    pub data_option2: Option<SortDataOption>,
    /// Optional third-key data treatment.
    pub data_option3: Option<SortDataOption>,
}

/// Apartment-bound wrapper for Excel's persistent Sort state.
pub struct Sort {
    inner: DispatchObject,
}
impl Debug for Sort {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("Sort").field(&self.inner).finish()
    }
}
impl Clone for Sort {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl Sort {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "Sort",
            },
        }
    }
    /// Returns the collection of configured sort fields.
    pub fn sort_fields(&self) -> Result<SortFields, ExcelComError> {
        object_get(
            &self.inner,
            "excel.sort.sortfields",
            SortFields::from_dispatch,
        )
    }
    /// Sets the Range to be changed when this persistent sort is applied.
    pub fn set_range(&self, range: &Range) -> Result<(), ExcelComError> {
        let _ = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.sort.setrange"), false),
            vec![OwnedVariant::dispatch_borrowed(
                &range.dispatch_object().dispatch,
            )],
            false,
        )?;
        Ok(())
    }
    /// Returns the current header interpretation.
    pub fn header(&self) -> Result<TableHeaderMode, ExcelComError> {
        Ok(TableHeaderMode::from_raw(i32_get(
            &self.inner,
            "excel.sort.header",
            "Sort.Header",
        )?))
    }
    /// Changes the header interpretation.
    pub fn set_header(&self, mode: TableHeaderMode) -> Result<(), ExcelComError> {
        i32_put(&self.inner, "excel.sort.header", mode.raw())
    }
    /// Returns whether comparison is case-sensitive.
    pub fn match_case(&self) -> Result<bool, ExcelComError> {
        bool_get(&self.inner, "excel.sort.matchcase")
    }
    /// Changes case-sensitive comparison.
    pub fn set_match_case(&self, enabled: bool) -> Result<(), ExcelComError> {
        bool_put(&self.inner, "excel.sort.matchcase", enabled)
    }
    /// Returns the sort orientation.
    pub fn orientation(&self) -> Result<SortOrientation, ExcelComError> {
        Ok(SortOrientation::from_raw(i32_get(
            &self.inner,
            "excel.sort.orientation",
            "Sort.Orientation",
        )?))
    }
    /// Changes the sort orientation.
    pub fn set_orientation(&self, orientation: SortOrientation) -> Result<(), ExcelComError> {
        i32_put(&self.inner, "excel.sort.orientation", orientation.raw())
    }
    /// Applies the configured persistent sort, modifying its configured Range in place.
    pub fn apply(&self) -> Result<(), ExcelComError> {
        let _ = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.sort.apply"), false),
            vec![],
            false,
        )?;
        Ok(())
    }
}

/// Apartment-bound collection of persistent Excel sort fields.
pub struct SortFields {
    inner: DispatchObject,
}
impl Debug for SortFields {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("SortFields").field(&self.inner).finish()
    }
}
impl Clone for SortFields {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl SortFields {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "SortFields",
            },
        }
    }
    /// Returns the number of configured fields.
    pub fn count(&self) -> Result<usize, ExcelComError> {
        collection_count(&self.inner, SORT_FIELDS)
    }
    /// Removes every configured sort field.
    pub fn clear(&self) -> Result<(), ExcelComError> {
        let _ = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.sortfields.clear"), false),
            vec![],
            false,
        )?;
        Ok(())
    }
    /// Adds a value-based sort field and returns its Excel wrapper.
    ///
    /// Color, icon, custom-list, and rich-data sort options remain outside this
    /// bounded API; their positional slots are explicitly Missing.
    pub fn add(
        &self,
        key: &Range,
        order: SortOrder,
        data_option: Option<SortDataOption>,
    ) -> Result<SortField, ExcelComError> {
        let mut args = PositionalArguments::new();
        args.push_object(key.dispatch_object());
        args.push_optional(None); // SortOn: Excel's default values sort.
        args.push_optional(Some(OwnedVariant::i32(order.raw())));
        args.push_optional(None); // CustomOrder
        args.push_optional(data_option.map(|value| OwnedVariant::i32(value.raw())));
        let mut value = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.sortfields.add"), false),
            args.into_inner(),
            false,
        )?;
        Ok(SortField::from_dispatch(value.take_dispatch()?))
    }
}

/// Apartment-bound wrapper returned for one configured Excel sort field.
///
/// The wrapper intentionally exposes no color, icon, custom-list, or rich
/// data configuration in this Prompt 14 slice.
pub struct SortField {
    inner: DispatchObject,
}
impl Debug for SortField {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("SortField").field(&self.inner).finish()
    }
}
impl Clone for SortField {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl SortField {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "SortField",
            },
        }
    }
}

impl Range {
    /// Sorts this Range in place using the documented positional Range.Sort surface.
    pub fn sort(&self, options: &RangeSortOptions<'_>) -> Result<(), ExcelComError> {
        if options.order2.is_some() && options.key2.is_none() {
            return Err(ExcelComError::Unsupported {
                detail: "Range.Sort order2 requires key2",
            });
        }
        if options.order3.is_some() && options.key3.is_none() {
            return Err(ExcelComError::Unsupported {
                detail: "Range.Sort order3 requires key3",
            });
        }
        let mut args = PositionalArguments::new();
        args.push_object(options.key1.dispatch_object());
        args.push_optional(Some(OwnedVariant::i32(options.order1.raw())));
        args.push_optional_object(options.key2.map(Range::dispatch_object));
        args.push_optional(None); // Type is meaningful for PivotTable sorting and is deferred.
        args.push_optional(options.order2.map(|value| OwnedVariant::i32(value.raw())));
        args.push_optional_object(options.key3.map(Range::dispatch_object));
        args.push_optional(options.order3.map(|value| OwnedVariant::i32(value.raw())));
        args.push_optional(Some(OwnedVariant::i32(options.header.raw())));
        args.push_optional(None); // OrderCustom
        args.push_optional(options.match_case.map(OwnedVariant::bool));
        args.push_optional(
            options
                .orientation
                .map(|value| OwnedVariant::i32(value.raw())),
        );
        args.push_optional(
            options
                .sort_method
                .map(|value| OwnedVariant::i32(value.raw())),
        );
        args.push_optional(
            options
                .data_option1
                .map(|value| OwnedVariant::i32(value.raw())),
        );
        args.push_optional(
            options
                .data_option2
                .map(|value| OwnedVariant::i32(value.raw())),
        );
        args.push_optional(
            options
                .data_option3
                .map(|value| OwnedVariant::i32(value.raw())),
        );
        let _ = invoke(
            &self.dispatch_object().dispatch,
            member(MemberId::new("excel.range.sort"), false),
            args.into_inner(),
            false,
        )?;
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use windows_sys::Win32::Foundation::DISP_E_PARAMNOTFOUND;

    #[test]
    fn range_sort_preserves_legacy_positional_holes() {
        let mut args = PositionalArguments::new();
        args.push_optional(Some(OwnedVariant::i32(1)));
        for _ in 0..8 {
            args.push_optional(None);
        }
        args.push_optional(Some(OwnedVariant::bool(false)));
        for _ in 0..5 {
            args.push_optional(None);
        }
        let values = args.into_inner();
        assert_eq!(values.len(), 15);
        assert_eq!(values[8].as_scode(), Some(DISP_E_PARAMNOTFOUND));
        assert_eq!(values[9].as_bool(), Some(false));
    }
}
