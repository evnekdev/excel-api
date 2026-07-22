//! Special-cell discovery and Excel-backed Range search operations.

use std::collections::BTreeSet;
use std::ops::{BitOr, BitOrAssign};

use crate::ExcelComError;
use crate::automation::{
    AutomationValue, ConversionPolicy, OwnedVariant, PositionalArguments, encode_variant, invoke,
};
use crate::excel::{Range, ReferenceStyle};
use crate::object_model::{MemberId, member};

/// An Excel `XlCellType` selector for [`Range::special_cells`].
#[derive(Clone, Copy, Debug, Eq, PartialEq, Hash)]
#[repr(transparent)]
pub struct SpecialCellType(i32);

impl SpecialCellType {
    /// `xlCellTypeConstants`.
    pub const CONSTANTS: Self = Self(2);
    /// `xlCellTypeFormulas`.
    pub const FORMULAS: Self = Self(-4123);
    /// `xlCellTypeBlanks`.
    pub const BLANKS: Self = Self(4);
    /// `xlCellTypeLastCell`.
    pub const LAST_CELL: Self = Self(11);
    /// `xlCellTypeVisible`.
    pub const VISIBLE: Self = Self(12);
    /// `xlCellTypeSameFormatConditions`.
    pub const SAME_FORMAT_CONDITIONS: Self = Self(-4173);
    /// `xlCellTypeAllFormatConditions`.
    pub const ALL_FORMAT_CONDITIONS: Self = Self(-4172);
    /// `xlCellTypeSameValidation`.
    pub const SAME_VALIDATION: Self = Self(-4175);
    /// `xlCellTypeAllValidation`.
    pub const ALL_VALIDATION: Self = Self(-4174);

    /// Preserves an Excel `XlCellType` value without validation.
    pub const fn from_raw(value: i32) -> Self {
        Self(value)
    }

    /// Returns the raw `XlCellType` value.
    pub const fn raw(self) -> i32 {
        self.0
    }
}

/// A bitmask of Excel `XlSpecialCellsValue` categories.
///
/// The mask is meaningful for `CONSTANTS` and `FORMULAS`; Excel controls how
/// it treats a mask supplied for other cell types.
#[derive(Clone, Copy, Debug, Eq, PartialEq, Hash)]
#[repr(transparent)]
pub struct SpecialCellValueMask(i32);

impl SpecialCellValueMask {
    /// `xlNumbers`.
    pub const NUMBERS: Self = Self(1);
    /// `xlTextValues`.
    pub const TEXT: Self = Self(2);
    /// `xlLogical`.
    pub const LOGICAL: Self = Self(4);
    /// `xlErrors`.
    pub const ERRORS: Self = Self(16);

    /// Preserves an Excel `XlSpecialCellsValue` mask without validation.
    pub const fn from_raw(value: i32) -> Self {
        Self(value)
    }

    /// Returns the raw `XlSpecialCellsValue` mask.
    pub const fn raw(self) -> i32 {
        self.0
    }
}

impl BitOr for SpecialCellValueMask {
    type Output = Self;

    fn bitor(self, rhs: Self) -> Self::Output {
        Self(self.0 | rhs.0)
    }
}

impl BitOrAssign for SpecialCellValueMask {
    fn bitor_assign(&mut self, rhs: Self) {
        self.0 |= rhs.0;
    }
}

/// An Excel `XlFindLookIn` value.
#[derive(Clone, Copy, Debug, Eq, PartialEq, Hash)]
#[repr(transparent)]
pub struct FindLookIn(i32);

impl FindLookIn {
    /// `xlFormulas`.
    pub const FORMULAS: Self = Self(-4123);
    /// `xlValues`.
    pub const VALUES: Self = Self(-4163);
    /// `xlComments`.
    pub const COMMENTS: Self = Self(-4144);
    /// `xlCommentsThreaded`.
    pub const COMMENTS_THREADED: Self = Self(-4184);
    /// `xlFormulas2`.
    pub const FORMULAS2: Self = Self(-4185);

    /// Preserves an Excel `XlFindLookIn` value without validation.
    pub const fn from_raw(value: i32) -> Self {
        Self(value)
    }

    /// Returns the raw `XlFindLookIn` value.
    pub const fn raw(self) -> i32 {
        self.0
    }
}

/// An Excel `XlLookAt` value used by Find and Replace.
#[derive(Clone, Copy, Debug, Eq, PartialEq, Hash)]
#[repr(transparent)]
pub struct FindMatchMode(i32);

impl FindMatchMode {
    /// `xlWhole`.
    pub const WHOLE: Self = Self(1);
    /// `xlPart`.
    pub const PART: Self = Self(2);

    /// Preserves an Excel `XlLookAt` value without validation.
    pub const fn from_raw(value: i32) -> Self {
        Self(value)
    }

    /// Returns the raw `XlLookAt` value.
    pub const fn raw(self) -> i32 {
        self.0
    }
}

/// An Excel `XlSearchOrder` value.
#[derive(Clone, Copy, Debug, Eq, PartialEq, Hash)]
#[repr(transparent)]
pub struct SearchOrder(i32);

impl SearchOrder {
    /// `xlByRows`.
    pub const BY_ROWS: Self = Self(1);
    /// `xlByColumns`.
    pub const BY_COLUMNS: Self = Self(2);

    /// Preserves an Excel `XlSearchOrder` value without validation.
    pub const fn from_raw(value: i32) -> Self {
        Self(value)
    }

    /// Returns the raw `XlSearchOrder` value.
    pub const fn raw(self) -> i32 {
        self.0
    }
}

/// An Excel `XlSearchDirection` value.
#[derive(Clone, Copy, Debug, Eq, PartialEq, Hash)]
#[repr(transparent)]
pub struct SearchDirection(i32);

impl SearchDirection {
    /// `xlNext`.
    pub const NEXT: Self = Self(1);
    /// `xlPrevious`.
    pub const PREVIOUS: Self = Self(2);

    /// Preserves an Excel `XlSearchDirection` value without validation.
    pub const fn from_raw(value: i32) -> Self {
        Self(value)
    }

    /// Returns the raw `XlSearchDirection` value.
    pub const fn raw(self) -> i32 {
        self.0
    }
}

/// Explicit optional arguments for [`Range::find`].
///
/// Excel remembers several Find dialog settings. The default deliberately
/// sends concrete values for every stateful setting, avoiding accidental reuse
/// of UI or earlier Automation state. Set a field to `None` only when the
/// caller intentionally wants to send Excel `Missing`.
#[derive(Debug)]
pub struct FindOptions<'a> {
    /// The single cell after which searching begins, or `None` for `Missing`.
    pub after: Option<&'a Range>,
    /// Formula/value/comment context, or `None` for `Missing`.
    pub look_in: Option<FindLookIn>,
    /// Whole-cell or partial match mode, or `None` for `Missing`.
    pub look_at: Option<FindMatchMode>,
    /// By-row or by-column search order, or `None` for `Missing`.
    pub search_order: Option<SearchOrder>,
    /// Search direction, or `None` for `Missing`.
    pub search_direction: Option<SearchDirection>,
    /// Case sensitivity, or `None` for `Missing`.
    pub match_case: Option<bool>,
    /// Double-byte matching behavior, or `None` for `Missing`.
    pub match_byte: Option<bool>,
    /// Whether to search format criteria, or `None` for `Missing`.
    pub search_format: Option<bool>,
}

impl Default for FindOptions<'_> {
    fn default() -> Self {
        Self {
            after: None,
            look_in: Some(FindLookIn::VALUES),
            look_at: Some(FindMatchMode::PART),
            search_order: Some(SearchOrder::BY_ROWS),
            search_direction: Some(SearchDirection::NEXT),
            match_case: Some(false),
            match_byte: Some(false),
            search_format: Some(false),
        }
    }
}

/// Explicit optional arguments for [`Range::replace`].
///
/// Defaults are concrete for all remembered Find/Replace settings. Formula
/// version remains omitted because this bounded API deliberately avoids a
/// second replacement-mode abstraction.
#[derive(Debug)]
pub struct ReplaceOptions {
    /// Whole-cell or partial match mode, or `None` for `Missing`.
    pub look_at: Option<FindMatchMode>,
    /// By-row or by-column search order, or `None` for `Missing`.
    pub search_order: Option<SearchOrder>,
    /// Case sensitivity, or `None` for `Missing`.
    pub match_case: Option<bool>,
    /// Double-byte matching behavior, or `None` for `Missing`.
    pub match_byte: Option<bool>,
    /// Whether to search format criteria, or `None` for `Missing`.
    pub search_format: Option<bool>,
    /// Whether to apply replacement-format criteria, or `None` for `Missing`.
    pub replace_format: Option<bool>,
}

impl Default for ReplaceOptions {
    fn default() -> Self {
        Self {
            look_at: Some(FindMatchMode::PART),
            search_order: Some(SearchOrder::BY_ROWS),
            match_case: Some(false),
            match_byte: Some(false),
            search_format: Some(false),
            replace_format: Some(false),
        }
    }
}

/// A fallible, apartment-bound iterator over all matches from [`Range::find_all`].
///
/// The iterator starts one Excel Find operation, then follows `FindNext` and
/// stops before Excel's wraparound would emit an already seen external address.
/// A terminal COM failure is yielded once and then fuses the iterator.
pub struct RangeFindIter {
    range: Range,
    pending: Option<Range>,
    after: Option<Range>,
    addresses: FindAddressTracker,
    done: bool,
}

impl Iterator for RangeFindIter {
    type Item = Result<Range, ExcelComError>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.done {
            return None;
        }
        if let Some(found) = self.pending.take() {
            self.after = Some(found.clone());
            return Some(Ok(found));
        }
        let Some(after) = self.after.as_ref() else {
            self.done = true;
            return None;
        };
        let found = match self.range.find_next(Some(after)) {
            Ok(Some(found)) => found,
            Ok(None) => {
                self.done = true;
                return None;
            }
            Err(error) => {
                self.done = true;
                return Some(Err(error));
            }
        };
        let address = match normalized_search_address(&found) {
            Ok(address) => address,
            Err(error) => {
                self.done = true;
                return Some(Err(error));
            }
        };
        if !self.addresses.observe(address) {
            self.done = true;
            return None;
        }
        self.after = Some(found.clone());
        Some(Ok(found))
    }
}

impl Range {
    /// Returns Excel's matching cells for a `SpecialCells` selector.
    ///
    /// The returned Range may be multi-area. When no matching cells exist,
    /// Excel's structured invocation error is preserved rather than converted
    /// to an empty Range.
    pub fn special_cells(
        &self,
        cell_type: SpecialCellType,
        value_mask: Option<SpecialCellValueMask>,
    ) -> Result<Range, ExcelComError> {
        let mut arguments = PositionalArguments::new();
        arguments.push_required(OwnedVariant::i32(cell_type.raw()));
        arguments.push_optional(value_mask.map(|value| OwnedVariant::i32(value.raw())));
        self.required_search_range("excel.range.specialcells", arguments.into_inner())
    }

    /// Starts an Excel-backed search with explicit typed options.
    ///
    /// Text and finite numeric `AutomationValue` inputs are accepted. A
    /// no-match result becomes `Ok(None)`; a matching result remains an
    /// apartment-bound Range.
    pub fn find(
        &self,
        what: &AutomationValue,
        options: &FindOptions<'_>,
    ) -> Result<Option<Range>, ExcelComError> {
        let mut value = invoke(
            &self.dispatch_object().dispatch,
            member(MemberId::new("excel.range.find"), false),
            find_arguments(what, options)?,
            false,
        )?;
        value
            .take_optional_dispatch()
            .map(|value| value.map(Range::from_dispatch))
    }

    /// Continues the current Excel Find state in the forward direction.
    ///
    /// Excel state comes from a preceding [`Self::find`] call. `after`, when
    /// supplied, must be a single cell in this search Range.
    pub fn find_next(&self, after: Option<&Range>) -> Result<Option<Range>, ExcelComError> {
        self.find_relative("excel.range.findnext", after)
    }

    /// Continues the current Excel Find state in the backward direction.
    ///
    /// Excel state comes from a preceding [`Self::find`] call. `after`, when
    /// supplied, must be a single cell in this search Range.
    pub fn find_previous(&self, after: Option<&Range>) -> Result<Option<Range>, ExcelComError> {
        self.find_relative("excel.range.findprevious", after)
    }

    /// Starts a wrap-safe iterator over all matching cells.
    ///
    /// The iterator stores normalized external addresses, not COM identities,
    /// because Excel may materialize equivalent Range objects separately.
    pub fn find_all(
        &self,
        what: &AutomationValue,
        options: &FindOptions<'_>,
    ) -> Result<RangeFindIter, ExcelComError> {
        let initial = self.find(what, options)?;
        let mut addresses = FindAddressTracker::default();
        if let Some(found) = initial.as_ref() {
            let _ = addresses.observe(normalized_search_address(found)?);
        }
        Ok(RangeFindIter {
            range: self.clone(),
            pending: initial,
            after: None,
            addresses,
            done: false,
        })
    }

    /// Replaces supported scalar values according to explicit typed options.
    ///
    /// The returned Boolean is Excel's replacement-result flag. Text, finite
    /// numbers, Booleans, and Excel errors are accepted; arrays and other
    /// Automation values fail before COM.
    pub fn replace(
        &self,
        what: &AutomationValue,
        replacement: &AutomationValue,
        options: &ReplaceOptions,
    ) -> Result<bool, ExcelComError> {
        let result = invoke(
            &self.dispatch_object().dispatch,
            member(MemberId::new("excel.range.replace-3305"), false),
            replace_arguments(what, replacement, options)?,
            false,
        )?;
        result.as_bool().ok_or(ExcelComError::Unsupported {
            detail: "Range.Replace did not return VT_BOOL",
        })
    }

    fn required_search_range(
        &self,
        id: &'static str,
        arguments: Vec<OwnedVariant>,
    ) -> Result<Range, ExcelComError> {
        let mut value = invoke(
            &self.dispatch_object().dispatch,
            member(MemberId::new(id), false),
            arguments,
            false,
        )?;
        Ok(Range::from_dispatch(value.take_dispatch()?))
    }

    fn find_relative(
        &self,
        id: &'static str,
        after: Option<&Range>,
    ) -> Result<Option<Range>, ExcelComError> {
        let mut arguments = PositionalArguments::new();
        arguments.push_optional_object(after.map(Range::dispatch_object));
        let mut value = invoke(
            &self.dispatch_object().dispatch,
            member(MemberId::new(id), false),
            arguments.into_inner(),
            false,
        )?;
        value
            .take_optional_dispatch()
            .map(|value| value.map(Range::from_dispatch))
    }
}

#[derive(Default)]
struct FindAddressTracker {
    seen: BTreeSet<String>,
}

impl FindAddressTracker {
    fn observe(&mut self, address: String) -> bool {
        self.seen.insert(address)
    }
}

fn normalized_search_address(range: &Range) -> Result<String, ExcelComError> {
    range
        .external_address(ReferenceStyle::A1)
        .map(|address| address.replace('$', "").to_ascii_lowercase())
}

fn find_arguments(
    what: &AutomationValue,
    options: &FindOptions<'_>,
) -> Result<Vec<OwnedVariant>, ExcelComError> {
    let mut arguments = PositionalArguments::new();
    arguments.push_result(find_value(what))?;
    arguments.push_optional_object(options.after.map(Range::dispatch_object));
    arguments.push_optional(options.look_in.map(|value| OwnedVariant::i32(value.raw())));
    arguments.push_optional(options.look_at.map(|value| OwnedVariant::i32(value.raw())));
    arguments.push_optional(
        options
            .search_order
            .map(|value| OwnedVariant::i32(value.raw())),
    );
    arguments.push_optional(
        options
            .search_direction
            .map(|value| OwnedVariant::i32(value.raw())),
    );
    arguments.push_optional(options.match_case.map(OwnedVariant::bool));
    arguments.push_optional(options.match_byte.map(OwnedVariant::bool));
    arguments.push_optional(options.search_format.map(OwnedVariant::bool));
    Ok(arguments.into_inner())
}

fn replace_arguments(
    what: &AutomationValue,
    replacement: &AutomationValue,
    options: &ReplaceOptions,
) -> Result<Vec<OwnedVariant>, ExcelComError> {
    let mut arguments = PositionalArguments::new();
    arguments.push_result(replace_value(what))?;
    arguments.push_result(replace_value(replacement))?;
    arguments.push_optional(options.look_at.map(|value| OwnedVariant::i32(value.raw())));
    arguments.push_optional(
        options
            .search_order
            .map(|value| OwnedVariant::i32(value.raw())),
    );
    arguments.push_optional(options.match_case.map(OwnedVariant::bool));
    arguments.push_optional(options.match_byte.map(OwnedVariant::bool));
    arguments.push_optional(options.search_format.map(OwnedVariant::bool));
    arguments.push_optional(options.replace_format.map(OwnedVariant::bool));
    arguments.push_optional(None);
    Ok(arguments.into_inner())
}

fn find_value(value: &AutomationValue) -> Result<OwnedVariant, ExcelComError> {
    match value {
        AutomationValue::Text(_) | AutomationValue::Number(_) => {
            encode_variant(value, ConversionPolicy::default())
        }
        _ => Err(ExcelComError::Unsupported {
            detail: "Range.Find accepts only text or numeric Automation values",
        }),
    }
}

fn replace_value(value: &AutomationValue) -> Result<OwnedVariant, ExcelComError> {
    match value {
        AutomationValue::Text(_)
        | AutomationValue::Number(_)
        | AutomationValue::Bool(_)
        | AutomationValue::Error(_) => encode_variant(value, ConversionPolicy::default()),
        _ => Err(ExcelComError::Unsupported {
            detail: "Range.Replace accepts only text, numeric, Boolean, or error Automation values",
        }),
    }
}

#[cfg(test)]
mod tests {
    use super::{
        FindAddressTracker, FindLookIn, FindMatchMode, FindOptions, ReplaceOptions,
        SearchDirection, SearchOrder, SpecialCellType, SpecialCellValueMask, find_arguments,
        replace_arguments,
    };
    use crate::AutomationValue;
    use windows_sys::Win32::Foundation::DISP_E_PARAMNOTFOUND;

    #[test]
    fn special_cell_masks_combine_with_excel_values() {
        let mask = SpecialCellValueMask::NUMBERS | SpecialCellValueMask::TEXT;
        assert_eq!(mask.raw(), 3);
        assert_eq!(SpecialCellType::FORMULAS.raw(), -4123);
    }

    #[test]
    fn find_defaults_make_remembered_state_explicit() {
        let values = find_arguments(
            &AutomationValue::Text("target".to_owned()),
            &FindOptions::default(),
        )
        .expect("arguments");
        assert_eq!(values.len(), 9);
        assert_eq!(values[0].as_string().expect("what"), "target");
        assert_eq!(values[1].as_scode(), Some(DISP_E_PARAMNOTFOUND));
        assert_eq!(values[2].as_i32(), Some(FindLookIn::VALUES.raw()));
        assert_eq!(values[3].as_i32(), Some(FindMatchMode::PART.raw()));
        assert_eq!(values[4].as_i32(), Some(SearchOrder::BY_ROWS.raw()));
        assert_eq!(values[5].as_i32(), Some(SearchDirection::NEXT.raw()));
        assert_eq!(values[6].as_bool(), Some(false));
        assert_eq!(values[7].as_bool(), Some(false));
        assert_eq!(values[8].as_bool(), Some(false));
        assert!(find_arguments(&AutomationValue::Bool(true), &FindOptions::default()).is_err());
    }

    #[test]
    fn replace_preserves_every_optional_position() {
        let values = replace_arguments(
            &AutomationValue::Text("old".to_owned()),
            &AutomationValue::Text("new".to_owned()),
            &ReplaceOptions::default(),
        )
        .expect("arguments");
        assert_eq!(values.len(), 9);
        assert_eq!(values[0].as_string().expect("what"), "old");
        assert_eq!(values[1].as_string().expect("replacement"), "new");
        assert_eq!(values[2].as_i32(), Some(FindMatchMode::PART.raw()));
        assert_eq!(values[3].as_i32(), Some(SearchOrder::BY_ROWS.raw()));
        assert_eq!(values[4].as_bool(), Some(false));
        assert_eq!(values[5].as_bool(), Some(false));
        assert_eq!(values[6].as_bool(), Some(false));
        assert_eq!(values[7].as_bool(), Some(false));
        assert_eq!(values[8].as_scode(), Some(DISP_E_PARAMNOTFOUND));
        assert!(
            replace_arguments(
                &AutomationValue::Array(
                    crate::AutomationArray::row(vec![AutomationValue::Text("old".to_owned())])
                        .expect("array"),
                ),
                &AutomationValue::Text("new".to_owned()),
                &ReplaceOptions::default(),
            )
            .is_err()
        );
    }

    #[test]
    fn address_tracker_stops_wraparound_and_duplicates() {
        let mut tracker = FindAddressTracker::default();
        assert!(tracker.observe("[book]sheet!a1".to_owned()));
        assert!(tracker.observe("[book]sheet!a2".to_owned()));
        assert!(!tracker.observe("[book]sheet!a1".to_owned()));
        assert!(!tracker.observe("[book]sheet!a2".to_owned()));
    }
}
