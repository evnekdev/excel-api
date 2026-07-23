//! Public, forward-compatible PivotTable value types and layout descriptors.

use crate::excel::Range;
use crate::{AutomationValue, ExcelComError};

macro_rules! raw_type {
    ($(#[$docs:meta])* $name:ident { $($constant:ident = $value:expr => $constant_docs:literal;)* }) => {
        $(#[$docs])*
        #[derive(Clone, Copy, Debug, Eq, PartialEq, Hash)]
        #[repr(transparent)]
        pub struct $name(i32);
        impl $name {
            $(#[doc = $constant_docs] pub const $constant: Self = Self($value);)*
            /// Preserves a future Excel value.
            pub const fn from_raw(value: i32) -> Self { Self(value) }
            /// Returns Excel's raw value.
            pub const fn raw(self) -> i32 { self.0 }
        }
    };
}
raw_type! {
    /// A forward-compatible `XlPivotTableSourceType`.
    PivotSourceType { DATABASE = 1 => "`xlDatabase`."; EXTERNAL = 2 => "`xlExternal`."; CONSOLIDATION = 3 => "`xlConsolidation`."; SCENARIO = 4 => "`xlScenario`."; PIVOT_TABLE = -4148 => "`xlPivotTable`."; }
}
raw_type! {
    /// A forward-compatible `XlPivotTableVersionList`.
    PivotTableVersion { VERSION_2000 = 0 => "`xlPivotTableVersion2000`."; VERSION_10 = 1 => "`xlPivotTableVersion10`."; VERSION_11 = 2 => "`xlPivotTableVersion11`."; VERSION_12 = 3 => "`xlPivotTableVersion12`."; VERSION_14 = 4 => "`xlPivotTableVersion14`."; VERSION_15 = 5 => "`xlPivotTableVersion15`."; CURRENT = -1 => "`xlPivotTableVersionCurrent`."; }
}
raw_type! {
    /// A forward-compatible `XlPivotFieldOrientation`.
    PivotFieldOrientation { HIDDEN = 0 => "`xlHidden`."; ROW = 1 => "`xlRowField`."; COLUMN = 2 => "`xlColumnField`."; PAGE = 3 => "`xlPageField`."; DATA = 4 => "`xlDataField`."; }
}
raw_type! {
    /// A forward-compatible `XlPivotTableMissingItems` value.
    MissingItemsLimit { NONE = 0 => "No obsolete items retained."; DEFAULT = -1 => "Excel default retention."; MAXIMUM = 32500 => "Excel maximum retention."; }
}
raw_type! {
    /// A useful, typed subset of `XlPivotFilterType`.
    PivotFilterType { LABEL_EQUALS = 15 => "Caption equals."; LABEL_NOT_EQUALS = 16 => "Caption does not equal."; LABEL_BEGINS_WITH = 17 => "Caption begins with."; LABEL_CONTAINS = 21 => "Caption contains."; LABEL_GREATER_THAN = 23 => "Caption greater than."; LABEL_LESS_THAN = 25 => "Caption less than."; LABEL_BETWEEN = 27 => "Caption is between."; TOP_COUNT = 1 => "Top count."; TOP_PERCENT = 3 => "Top percent."; VALUE_EQUALS = 7 => "Value equals."; VALUE_GREATER_THAN = 9 => "Value greater than."; VALUE_LESS_THAN = 11 => "Value less than."; VALUE_BETWEEN = 13 => "Value is between."; }
}

/// Exact optional arguments for `PivotCache.CreatePivotTable`.
#[derive(Debug)]
pub struct PivotTableCreateOptions<'a> {
    /// Destination cell in the target workbook.
    pub destination: &'a Range,
    /// Excel-visible PivotTable name.
    pub name: &'a str,
    /// Recorded requested cache version; Excel may restrict supported values.
    pub version: Option<PivotTableVersion>,
    /// Whether Excel should read all data before creating the report.
    pub read_data: Option<bool>,
    /// Optional default PivotTable version argument.
    pub default_version: Option<PivotTableVersion>,
}

/// One field placement applied by [`super::PivotTable::apply_layout`].
#[derive(Debug)]
pub struct PivotFieldPlacement<'a> {
    /// Source field name.
    pub field_name: &'a str,
    /// Excel report orientation.
    pub orientation: PivotFieldOrientation,
    /// Optional one-based position in that orientation collection.
    pub position: Option<usize>,
}

/// One data-field aggregation applied by [`super::PivotTable::apply_layout`].
#[derive(Debug)]
pub struct PivotDataField<'a> {
    /// Source field name.
    pub field_name: &'a str,
    /// Optional report caption.
    pub caption: Option<&'a str>,
    /// Excel aggregation function.
    pub function: crate::AggregationFunction,
    /// Optional Excel number format.
    pub number_format: Option<&'a str>,
}

/// A bounded, declarative PivotTable field layout.
#[derive(Debug, Default)]
pub struct PivotLayoutOptions<'a> {
    /// Row, column, page, hidden, or existing data field placements.
    pub fields: Vec<PivotFieldPlacement<'a>>,
    /// Data field aggregations.
    pub data_fields: Vec<PivotDataField<'a>>,
}

/// Typed label-filter arguments. Only the documented label subset is accepted.
#[derive(Debug)]
pub struct PivotLabelFilterOptions<'a> {
    /// One of the supported label filter kinds.
    pub filter_type: PivotFilterType,
    /// First comparison value.
    pub value1: &'a AutomationValue,
    /// Second comparison value for `between` filters.
    pub value2: Option<&'a AutomationValue>,
}

/// Typed value-filter arguments. A data field is required by Excel.
#[derive(Debug)]
pub struct PivotValueFilterOptions<'a> {
    /// One of the supported value filter kinds.
    pub filter_type: PivotFilterType,
    /// The data field the filter measures.
    pub data_field: &'a super::PivotField,
    /// First comparison value.
    pub value1: &'a AutomationValue,
    /// Second comparison value for `between` filters.
    pub value2: Option<&'a AutomationValue>,
}

pub(crate) fn valid_filter(value: PivotFilterType) -> Result<(), ExcelComError> {
    match value {
        PivotFilterType::LABEL_EQUALS
        | PivotFilterType::LABEL_NOT_EQUALS
        | PivotFilterType::LABEL_BEGINS_WITH
        | PivotFilterType::LABEL_CONTAINS
        | PivotFilterType::LABEL_GREATER_THAN
        | PivotFilterType::LABEL_LESS_THAN
        | PivotFilterType::LABEL_BETWEEN
        | PivotFilterType::TOP_COUNT
        | PivotFilterType::TOP_PERCENT
        | PivotFilterType::VALUE_EQUALS
        | PivotFilterType::VALUE_GREATER_THAN
        | PivotFilterType::VALUE_LESS_THAN
        | PivotFilterType::VALUE_BETWEEN => Ok(()),
        _ => Err(ExcelComError::Unsupported {
            detail: "PivotFilterType is outside this crate's bounded filter subset",
        }),
    }
}
