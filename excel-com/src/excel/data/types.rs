use std::path::Path;

use crate::automation::AutomationArray;
use crate::excel::Range;

macro_rules! raw_data_type {
    ($(#[$docs:meta])* $name:ident { $($constant:ident = $value:expr => $constant_docs:literal;)* }) => {
        $(#[$docs])*
        #[derive(Clone, Copy, Debug, Eq, PartialEq, Hash)]
        #[repr(transparent)]
        pub struct $name(i32);
        impl $name {
            $(#[doc = $constant_docs] pub const $constant: Self = Self($value);)*
            /// Preserves an Excel value not yet named by this wrapper.
            pub const fn from_raw(value: i32) -> Self { Self(value) }
            /// Returns the exact Excel integer value.
            pub const fn raw(self) -> i32 { self.0 }
        }
    };
}

raw_data_type! {
    /// A forward-compatible `XlTextParsingType` value.
    TextParsingType {
        DELIMITED = 1 => "`xlDelimited`.";
        FIXED_WIDTH = 2 => "`xlFixedWidth`.";
    }
}
raw_data_type! {
    /// A forward-compatible `XlTextQualifier` value.
    TextQualifier {
        DOUBLE_QUOTE = 1 => "`xlTextQualifierDoubleQuote`.";
        SINGLE_QUOTE = 2 => "`xlTextQualifierSingleQuote`.";
        NONE = -4142 => "`xlTextQualifierNone`.";
    }
}
raw_data_type! {
    /// A forward-compatible `XlColumnDataType` value used by Excel field info.
    TextColumnType {
        GENERAL = 1 => "`xlGeneralFormat`.";
        TEXT = 2 => "`xlTextFormat`.";
        MDY = 3 => "`xlMDYFormat`.";
        DMY = 4 => "`xlDMYFormat`.";
        YMD = 5 => "`xlYMDFormat`.";
        MYD = 6 => "`xlMYDFormat`.";
        DYM = 7 => "`xlDYMFormat`.";
        YDM = 8 => "`xlYDMFormat`.";
        SKIP = 9 => "`xlSkipColumn`.";
    }
}
raw_data_type! {
    /// A forward-compatible source-platform value for `Workbooks.OpenText`.
    TextPlatform {
        MACINTOSH = 1 => "`xlMacintosh`.";
        WINDOWS = 2 => "`xlWindows`.";
        MSDOS = 3 => "`xlMSDOS`.";
    }
}
raw_data_type! {
    /// A forward-compatible `XlAutoFillType` value.
    AutoFillType {
        DEFAULT = 0 => "`xlFillDefault`.";
        COPY = 1 => "`xlFillCopy`.";
        SERIES = 2 => "`xlFillSeries`.";
        FORMATS = 3 => "`xlFillFormats`.";
        VALUES = 4 => "`xlFillValues`.";
        DAYS = 5 => "`xlFillDays`.";
        WEEKDAYS = 6 => "`xlFillWeekdays`.";
        MONTHS = 7 => "`xlFillMonths`.";
        YEARS = 8 => "`xlFillYears`.";
        LINEAR_TREND = 9 => "`xlFillLinear`.";
        GROWTH_TREND = 10 => "`xlFillGrowth`.";
        FLASH_FILL = 11 => "`xlFlashFill`.";
    }
}
raw_data_type! {
    /// A forward-compatible `XlRowCol` orientation for `Range.DataSeries`.
    SeriesOrientation {
        ROWS = 1 => "`xlRows`.";
        COLUMNS = 2 => "`xlColumns`.";
    }
}
raw_data_type! {
    /// A forward-compatible `XlDataSeriesType` value.
    DataSeriesType {
        LINEAR = -4132 => "`xlLinear`.";
        GROWTH = -4133 => "`xlGrowth`.";
        DATE = 2 => "`xlDataSeriesDate`.";
        AUTO_FILL = 4 => "`xlAutoFill`.";
    }
}
raw_data_type! {
    /// A forward-compatible `XlDataSeriesDate` value.
    DataSeriesDateUnit {
        DAY = 1 => "`xlDay`.";
        WEEKDAY = 2 => "`xlWeekday`.";
        MONTH = 3 => "`xlMonth`.";
        YEAR = 4 => "`xlYear`.";
    }
}
raw_data_type! {
    /// A forward-compatible `XlFilterAction` value.
    AdvancedFilterAction {
        FILTER_IN_PLACE = 1 => "`xlFilterInPlace`.";
        COPY = 2 => "`xlFilterCopy`.";
    }
}
raw_data_type! {
    /// A forward-compatible `XlConsolidationFunction` value.
    AggregationFunction {
        SUM = -4157 => "`xlSum`.";
        COUNT = -4112 => "`xlCount`.";
        AVERAGE = -4106 => "`xlAverage`.";
        MAX = -4136 => "`xlMax`.";
        MIN = -4139 => "`xlMin`.";
        PRODUCT = -4149 => "`xlProduct`.";
        COUNT_NUMS = -4113 => "`xlCountNums`.";
        STD_DEV = -4155 => "`xlStdDev`.";
        STD_DEV_P = -4156 => "`xlStdDevP`.";
        VAR = -4164 => "`xlVar`.";
        VAR_P = -4165 => "`xlVarP`.";
    }
}
raw_data_type! {
    /// A forward-compatible `XlLinkType` value.
    LinkType {
        EXCEL_LINK = 1 => "`xlLinkTypeExcelLinks`.";
        OLE_LINK = 2 => "`xlLinkTypeOLELinks`.";
    }
}
raw_data_type! {
    /// A forward-compatible `XlLinkStatus` value.
    LinkStatus {
        OK = 0 => "`xlLinkStatusOK`.";
        MISSING_FILE = 1 => "`xlLinkStatusMissingFile`.";
        MISSING_SHEET = 2 => "`xlLinkStatusMissingSheet`.";
        MISSING_RANGE = 3 => "`xlLinkStatusMissingRange`.";
        NOT_STARTED = 4 => "`xlLinkStatusNotStarted`.";
        INVALID_NAME = 7 => "`xlLinkStatusInvalidName`.";
        SOURCE_NOT_CALCULATED = 8 => "`xlLinkStatusSourceNotCalculated`.";
        IND_SOURCE_NOT_OPENED = 9 => "`xlLinkStatusIndSourceNotOpened`.";
        SOURCE_NOT_OPENED = 10 => "`xlLinkStatusSourceNotOpened`.";
        SOURCE_OPENED = 11 => "`xlLinkStatusSourceOpened`.";
        SOURCE_OPENED_CHANGED = 12 => "`xlLinkStatusSourceOpenAndChanged`.";
    }
}
raw_data_type! {
    /// A forward-compatible scenario-summary report type.
    ScenarioReportType {
        SUMMARY = 1 => "`xlSummary`.";
        PIVOT_TABLE = 2 => "`xlPivotTable`.";
    }
}
raw_data_type! {
    /// An Excel text file format accepted by `Workbook.SaveAs`.
    TextFileFormat {
        CSV = 6 => "`xlCSV`.";
        CSV_WINDOWS = 23 => "`xlCSVWindows`.";
        CSV_MACINTOSH = 22 => "`xlCSVMac`.";
        CSV_MSDOS = 24 => "`xlCSVMSDOS`.";
        CSV_UTF8 = 62 => "`xlCSVUTF8`.";
        TAB_DELIMITED = -4158 => "`xlText`.";
        TEXT_WINDOWS = 20 => "`xlTextWindows`.";
        TEXT_MACINTOSH = 19 => "`xlTextMac`.";
        TEXT_MSDOS = 21 => "`xlTextMSDOS`.";
        UNICODE_TEXT = 42 => "`xlUnicodeText`.";
        FORMATTED_TEXT = 36 => "`xlTextPrinter`.";
    }
}

/// Delimiter switches supplied to Excel text import operations.
#[derive(Clone, Debug)]
pub enum TextDelimiter {
    /// Enables tab delimiters.
    Tab,
    /// Enables semicolon delimiters.
    Semicolon,
    /// Enables comma delimiters.
    Comma,
    /// Enables space delimiters.
    Space,
    /// Enables one other delimiter character.
    Other(char),
    /// Supplies multiple Excel delimiter switches explicitly.
    Custom {
        /// Enables tabs.
        tab: bool,
        /// Enables semicolons.
        semicolon: bool,
        /// Enables commas.
        comma: bool,
        /// Enables spaces.
        space: bool,
        /// Supplies Excel's one-character `OtherChar` when present.
        other: Option<char>,
    },
}

/// One `FieldInfo` entry for an Excel text-import operation.
#[derive(Clone, Debug)]
pub struct TextColumnSpec {
    /// One-based field number for delimited input or zero-based character start for fixed width.
    pub start: Option<usize>,
    /// Excel's requested output type for this field.
    pub column_type: TextColumnType,
}

/// Explicit options for `Workbooks.OpenText`; the operation creates a workbook.
#[derive(Debug)]
pub struct OpenTextOptions<'a> {
    /// Text-file path converted directly from its Windows UTF-16 representation.
    pub path: &'a Path,
    /// Optional text-file platform origin.
    pub origin: Option<TextPlatform>,
    /// One-based first row to import.
    pub start_row: Option<usize>,
    /// Delimited or fixed-width parser selection.
    pub parsing_type: TextParsingType,
    /// Optional quote qualifier.
    pub text_qualifier: Option<TextQualifier>,
    /// Whether consecutive delimiters are treated as one.
    pub consecutive_delimiters: Option<bool>,
    /// Optional Excel delimiter switches.
    pub delimiter: Option<TextDelimiter>,
    /// Excel FieldInfo entries.
    pub columns: Vec<TextColumnSpec>,
    /// Optional one-character decimal separator.
    pub decimal_separator: Option<char>,
    /// Optional one-character thousands separator.
    pub thousands_separator: Option<char>,
    /// Whether a trailing minus denotes a negative number.
    pub trailing_minus_numbers: Option<bool>,
    /// Whether Excel should apply local language settings.
    pub local: Option<bool>,
}

/// Explicit options for `Range.TextToColumns`, which overwrites worksheet cells.
#[derive(Debug)]
pub struct TextToColumnsOptions<'a> {
    /// Optional top-left destination Range.
    pub destination: Option<&'a Range>,
    /// Delimited or fixed-width parser selection.
    pub parsing_type: TextParsingType,
    /// Optional quote qualifier.
    pub text_qualifier: Option<TextQualifier>,
    /// Whether consecutive delimiters are treated as one.
    pub consecutive_delimiters: Option<bool>,
    /// Optional Excel delimiter switches.
    pub delimiter: Option<TextDelimiter>,
    /// Excel FieldInfo entries.
    pub columns: Vec<TextColumnSpec>,
    /// Optional one-character decimal separator.
    pub decimal_separator: Option<char>,
    /// Optional one-character thousands separator.
    pub thousands_separator: Option<char>,
    /// Whether a trailing minus denotes a negative number.
    pub trailing_minus_numbers: Option<bool>,
}

impl Default for TextToColumnsOptions<'_> {
    fn default() -> Self {
        Self {
            destination: None,
            parsing_type: TextParsingType::DELIMITED,
            text_qualifier: None,
            consecutive_delimiters: None,
            delimiter: None,
            columns: Vec::new(),
            decimal_separator: None,
            thousands_separator: None,
            trailing_minus_numbers: None,
        }
    }
}

/// Explicit options for Excel text-file export through `Workbook.SaveAs`.
#[derive(Debug)]
pub struct TextExportOptions<'a> {
    /// Output path passed directly to Excel's `SaveAs` member.
    pub path: &'a Path,
    /// Exact Excel text-file format.
    pub format: TextFileFormat,
    /// Whether Excel should use local language settings.
    pub local: Option<bool>,
    /// Whether Excel should create a backup.
    pub create_backup: Option<bool>,
}

/// Explicit options for Excel `Range.DataSeries`.
#[derive(Debug, Default)]
pub struct DataSeriesOptions {
    /// Optional row/column orientation.
    pub orientation: Option<SeriesOrientation>,
    /// Optional series calculation type.
    pub series_type: Option<DataSeriesType>,
    /// Optional unit for date series.
    pub date_unit: Option<DataSeriesDateUnit>,
    /// Optional finite increment.
    pub step_value: Option<f64>,
    /// Optional finite upper bound.
    pub stop_value: Option<f64>,
    /// Optional trend extrapolation request.
    pub trend: Option<bool>,
}

/// Explicit options for `Range.AdvancedFilter`.
#[derive(Debug)]
pub struct AdvancedFilterOptions<'a> {
    /// In-place or copy action.
    pub action: AdvancedFilterAction,
    /// Optional criteria range including headers where Excel requires them.
    pub criteria_range: Option<&'a Range>,
    /// Required only for a copy action.
    pub copy_to_range: Option<&'a Range>,
    /// Requests unique rows when `true`.
    pub unique: Option<bool>,
}

/// Explicit options for Excel's subtotal operation.
#[derive(Debug)]
pub struct SubtotalOptions {
    /// One-based receiver-relative group column.
    pub group_by: usize,
    /// Excel aggregation selector.
    pub function: AggregationFunction,
    /// One-based receiver-relative columns to aggregate.
    pub total_columns: Vec<usize>,
    /// Whether to replace existing subtotals.
    pub replace: Option<bool>,
    /// Whether Excel should insert page breaks between groups.
    pub page_breaks: Option<bool>,
    /// Whether summary rows appear below detail.
    pub summary_below_data: Option<bool>,
}

/// A controlled source for Excel `Range.Consolidate`.
pub enum ConsolidationSource<'a> {
    /// A local Range converted through Excel to its external-qualified address.
    Range(&'a Range),
    /// An Excel reference string; it must not contain NUL.
    Reference(&'a str),
}

/// Explicit options for Excel `Range.Consolidate`.
pub struct ConsolidateOptions<'a> {
    /// Source Ranges or controlled Excel references.
    pub sources: Vec<ConsolidationSource<'a>>,
    /// Excel aggregation selector.
    pub function: AggregationFunction,
    /// Whether source top rows are labels.
    pub top_row_labels: Option<bool>,
    /// Whether source left columns are labels.
    pub left_column_labels: Option<bool>,
    /// Whether Excel should create links to the sources.
    pub create_links: Option<bool>,
}

/// Explicit options for Excel's numerical Goal Seek solver.
#[derive(Debug)]
pub struct GoalSeekOptions<'a> {
    /// Finite formula result requested from Excel.
    pub goal: f64,
    /// The one-cell changing input passed to Excel.
    pub changing_cell: &'a Range,
}

/// Inputs for Excel's what-if Data Table feature, not a `ListObject` table.
pub enum DataTableInputs<'a> {
    /// A one-variable row-oriented table.
    Row {
        /// Formula input cell.
        row_input: &'a Range,
    },
    /// A one-variable column-oriented table.
    Column {
        /// Formula input cell.
        column_input: &'a Range,
    },
    /// A two-variable table.
    TwoVariable {
        /// Row input cell.
        row_input: &'a Range,
        /// Column input cell.
        column_input: &'a Range,
    },
}

/// Data supplied when adding a worksheet Scenario.
#[derive(Debug)]
pub struct ScenarioAddOptions<'a> {
    /// Scenario name.
    pub name: &'a str,
    /// One or more changing cells selected by Excel.
    pub changing_cells: &'a Range,
    /// Optional one-dimensional values matching changing cells where representable.
    pub values: Option<&'a AutomationArray>,
    /// Optional comment.
    pub comment: Option<&'a str>,
    /// Optional protection state.
    pub locked: Option<bool>,
    /// Optional hidden state.
    pub hidden: Option<bool>,
}

/// Explicit options for creation of an Excel Scenario report worksheet.
#[derive(Debug)]
pub struct ScenarioSummaryOptions<'a> {
    /// Summary or pivot-style report selection.
    pub report_type: ScenarioReportType,
    /// Optional result cells selected by Excel.
    pub result_cells: Option<&'a Range>,
}

/// An external source name returned unchanged by Excel.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct ExternalLinkSource {
    /// Exact source string supplied by Excel; it is never normalized by this crate.
    pub name: String,
    /// Link category used for later controlled operations.
    pub link_type: LinkType,
}
