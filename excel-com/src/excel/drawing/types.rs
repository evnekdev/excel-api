//! Typed, apartment-bound wrappers for Excel-native charts and drawing objects.
//!
//! The public surface deliberately models Excel and shared Office objects by
//! their concrete role. It does not expose a generic `IDispatch` escape hatch.
//! Coordinates are worksheet points, and Excel remains responsible for source
//! data interpretation, chart calculation, installed export filters, and
//! picture rendering.
#![allow(missing_docs)]

pub(super) use std::fmt::{Debug, Formatter};
pub(super) use std::iter::FusedIterator;
pub(super) use std::path::Path;

pub(super) use crate::automation::{
    AutomationArray, AutomationValue, ConversionPolicy, EnumVariant, OwnedVariant,
    PositionalArguments, decode_variant, encode_variant, enumerated_dispatch, invoke, property_get,
    property_put,
};
pub(super) use crate::excel::text::text_bstr;
pub(super) use crate::excel::{
    Application, DispatchObject, Font, Range, Sheet, SheetDestination, Workbook, Worksheet,
};
pub(super) use crate::internal::{ComPtr, Dispatch, path_bstr};
pub(super) use crate::object_model::{MemberId, member};
pub(super) use crate::{ConversionError, ExcelComError};

macro_rules! drawing_value {
    ($(#[$meta:meta])* $name:ident { $($(#[$constant_meta:meta])* $constant:ident = $value:expr;)* }) => {
        $(#[$meta])*
        #[repr(transparent)]
        #[allow(dead_code)]
        #[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
        pub struct $name(i32);
        #[allow(dead_code)]
        impl $name {
            $($(#[$constant_meta])* pub const $constant: Self = Self($value);)*
            /// Preserves a raw value returned by the installed Excel or Office version.
            pub const fn from_raw(value: i32) -> Self { Self(value) }
            /// Returns the raw COM enum value without narrowing it to known constants.
            pub const fn raw(self) -> i32 { self.0 }
        }
    };
}

drawing_value! {
/// A forward-compatible `XlChartType` value.
ChartType {
    /// `xlColumnClustered`.
    COLUMN_CLUSTERED = 51;
    /// `xlLine`.
    LINE = 4;
    /// `xlLineMarkers`.
    LINE_MARKERS = 65;
    /// `xlXYScatter`.
    XY_SCATTER = -4169;
} }
drawing_value! {
/// Excel's `XlRowCol` orientation for source data.
PlotBy {
    /// Treat each source row as a series.
    ROWS = 1;
    /// Treat each source column as a series.
    COLUMNS = 2;
} }
drawing_value! {
/// A forward-compatible `XlAxisType` selector.
AxisType {
    /// Category axis.
    CATEGORY = 1;
    /// Value axis.
    VALUE = 2;
    /// Series axis for applicable 3-D charts.
    SERIES = 3;
} }
drawing_value! {
/// A forward-compatible `XlAxisGroup` selector.
AxisGroup {
    /// Primary chart axis group.
    PRIMARY = 1;
    /// Secondary chart axis group.
    SECONDARY = 2;
} }
drawing_value! {
/// A forward-compatible `XlScaleType` value.
AxisScaleType {
    /// Linear scale.
    LINEAR = -4132;
    /// Logarithmic scale.
    LOGARITHMIC = -4133;
} }
drawing_value! {
/// A forward-compatible `XlTickMark` value.
TickMark {
    /// No tick mark.
    NONE = -4142;
    /// Tick marks inside the plot area.
    INSIDE = 2;
    /// Tick marks outside the plot area.
    OUTSIDE = 3;
    /// Tick marks crossing the axis.
    CROSS = 4;
} }
drawing_value! {
/// A forward-compatible `XlTickLabelPosition` value.
TickLabelPosition {
    /// No labels.
    NONE = -4142;
    /// Labels next to the axis.
    NEXT_TO_AXIS = 4;
    /// Labels at the high end.
    HIGH = -4127;
    /// Labels at the low end.
    LOW = -4134;
} }
drawing_value! {
/// A forward-compatible `XlLegendPosition` value.
LegendPosition {
    /// Bottom legend.
    BOTTOM = -4107;
    /// Corner legend.
    CORNER = 2;
    /// Left legend.
    LEFT = -4131;
    /// Right legend.
    RIGHT = -4152;
    /// Top legend.
    TOP = -4160;
} }
drawing_value! {
/// A forward-compatible `XlDataLabelsType` value.
DataLabelType {
    /// Do not show labels.
    NONE = -4142;
    /// Show values.
    VALUE = 2;
    /// Show category labels.
    LABEL = 4;
    /// Show percentages.
    PERCENT = 3;
    /// Show bubble sizes.
    BUBBLE_SIZE = 6;
} }
drawing_value! {
/// A forward-compatible `XlMarkerStyle` value.
MarkerStyle {
    /// Let Excel choose a marker.
    AUTOMATIC = -4105;
    /// No marker.
    NONE = -4142;
    /// Circular marker.
    CIRCLE = 8;
    /// Diamond marker.
    DIAMOND = 2;
    /// Square marker.
    SQUARE = 1;
    /// Triangle marker.
    TRIANGLE = 3;
} }
drawing_value! {
/// A forward-compatible `XlTrendlineType` value.
TrendlineType {
    /// Linear trendline.
    LINEAR = -4132;
    /// Exponential trendline.
    EXPONENTIAL = 5;
    /// Logarithmic trendline.
    LOGARITHMIC = -4133;
    /// Moving-average trendline.
    MOVING_AVERAGE = 6;
    /// Polynomial trendline.
    POLYNOMIAL = 3;
    /// Power trendline.
    POWER = 4;
} }
drawing_value! {
/// A forward-compatible error-bar direction.
ErrorBarDirection {
    /// X direction.
    X = -4168;
    /// Y direction.
    Y = 1;
} }
drawing_value! {
/// A forward-compatible error-bar inclusion value.
ErrorBarInclude {
    /// Plus and minus bars.
    BOTH = 1;
    /// Minus bars only.
    MINUS = 3;
    /// No bars.
    NONE = -4142;
    /// Plus bars only.
    PLUS = 2;
} }
drawing_value! {
/// A forward-compatible error-bar type.
ErrorBarType {
    /// Excel reads custom values from supplied series data.
    CUSTOM = -4114;
    /// Fixed-value bars.
    FIXED_VALUE = 1;
    /// Percentage bars.
    PERCENT = 2;
    /// Standard-deviation bars.
    STANDARD_DEVIATION = -4155;
    /// Standard-error bars.
    STANDARD_ERROR = 4;
} }
drawing_value! {
/// Excel `CopyPicture` appearance.
PictureAppearance {
    /// Screen appearance.
    SCREEN = 1;
    /// Printer appearance.
    PRINTER = 2;
} }
drawing_value! {
/// Excel `CopyPicture` rendering format.
CopyPictureFormat {
    /// Bitmap result.
    BITMAP = 2;
    /// Metafile picture result.
    PICTURE = -4147;
} }
drawing_value! {
/// Excel's cell-placement policy for a drawing object.
ShapePlacement {
    /// Object floats independently of cells.
    FREE_FLOATING = 3;
    /// Object moves with cells.
    MOVE = 2;
    /// Object moves and sizes with cells.
    MOVE_AND_SIZE = 1;
} }
drawing_value! {
/// A forward-compatible `XlSparkType` value.
SparklineType {
    /// Line sparkline.
    LINE = 1;
    /// Column sparkline.
    COLUMN = 2;
    /// Win/loss sparkline.
    WIN_LOSS = 3;
} }
drawing_value! {
/// A forward-compatible `XlSparkScale` value.
SparkScale {
    /// Scale each sparkline independently.
    SINGLE = 2;
    /// Share a group scale.
    GROUP = 1;
    /// Excel custom scale.
    CUSTOM = 3;
} }
drawing_value! {
/// A forward-compatible Office `MsoShapeType` value.
ShapeType {
    /// AutoShape.
    AUTO_SHAPE = 1;
    /// Grouped drawing.
    GROUP = 6;
    /// Picture.
    PICTURE = 13;
    /// Text box.
    TEXT_BOX = 17;
} }
drawing_value! {
/// A curated Office `MsoAutoShapeType` value.
AutoShapeType {
    /// Rectangle AutoShape.
    RECTANGLE = 1;
    /// Rounded rectangle AutoShape.
    ROUNDED_RECTANGLE = 5;
    /// Ellipse AutoShape.
    OVAL = 9;
} }
drawing_value! {
/// A forward-compatible Office text orientation.
TextOrientation {
    /// Horizontal Office text.
    HORIZONTAL = 1;
    /// Vertical Office text.
    VERTICAL = 5;
} }
drawing_value! {
/// A forward-compatible Office z-order command.
ZOrderCommand {
    /// Bring to front.
    BRING_TO_FRONT = 0;
    /// Send to back.
    SEND_TO_BACK = 1;
    /// Bring forward one level.
    BRING_FORWARD = 2;
    /// Send backward one level.
    SEND_BACKWARD = 3;
} }

/// Excel's enum-like `Application.CutCopyMode` result.
#[repr(transparent)]
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
pub struct CutCopyMode(i32);
impl CutCopyMode {
    /// No Excel copy or cut operation is active.
    pub const NONE: Self = Self(0);
    /// A copy operation is active.
    pub const COPY: Self = Self(1);
    /// A cut operation is active.
    pub const CUT: Self = Self(2);
    /// Preserves an unrecognized Excel value.
    pub const fn from_raw(value: i32) -> Self {
        Self(value)
    }
    /// Returns the raw Excel value.
    pub const fn raw(self) -> i32 {
        self.0
    }
}

/// Point-based bounds for an embedded chart object.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct ChartBounds {
    pub left: f64,
    pub top: f64,
    pub width: f64,
    pub height: f64,
}
/// A point used for a shape endpoint.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct ShapePoint {
    pub x: f64,
    pub y: f64,
}
/// Point-based bounds for a worksheet or chart shape.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct ShapeBounds {
    pub left: f64,
    pub top: f64,
    pub width: f64,
    pub height: f64,
}

/// Options for creating an embedded chart from an Excel-owned Range.
#[derive(Debug)]
pub struct ChartCreateOptions<'a> {
    pub source: &'a Range,
    pub chart_type: ChartType,
    pub bounds: ChartBounds,
    pub plot_by: Option<PlotBy>,
    pub title: Option<&'a str>,
    pub has_legend: Option<bool>,
}
/// Options for `Chart.Export`; installed Excel controls supported filters.
#[derive(Debug)]
pub struct ChartExportOptions<'a> {
    pub path: &'a Path,
    pub filter_name: Option<&'a str>,
    pub interactive: Option<bool>,
}
/// Options for Excel `CopyPicture`, which changes Excel's own cut/copy state.
#[derive(Debug)]
pub struct CopyPictureOptions {
    pub appearance: PictureAppearance,
    pub format: CopyPictureFormat,
}
/// Options used to create a Series from one Excel Range.
#[derive(Debug, Default)]
pub struct SeriesAddOptions {
    pub row_col: Option<PlotBy>,
    pub series_labels: Option<bool>,
    pub category_labels: Option<bool>,
    pub replace: Option<bool>,
}
/// A bounded series-data source. Arrays must be rank-two `AutomationArray` values.
#[derive(Debug)]
pub enum SeriesData<'a> {
    Range(&'a Range),
    Array(&'a AutomationArray),
    Formula(&'a str),
}
/// Data-label positions accepted by `Series.ApplyDataLabels`.
#[derive(Debug, Default)]
pub struct DataLabelOptions {
    pub label_type: Option<DataLabelType>,
    pub show_series_name: Option<bool>,
    pub show_category_name: Option<bool>,
    pub show_value: Option<bool>,
    pub show_percentage: Option<bool>,
    pub show_legend_key: Option<bool>,
    pub separator: Option<String>,
}
/// Trendline construction options delegated to Excel.
#[derive(Debug)]
pub struct TrendlineAddOptions<'a> {
    pub trendline_type: TrendlineType,
    pub order: Option<usize>,
    pub period: Option<usize>,
    pub forward: Option<f64>,
    pub backward: Option<f64>,
    pub intercept: Option<f64>,
    pub display_equation: Option<bool>,
    pub display_r_squared: Option<bool>,
    pub name: Option<&'a str>,
}
impl Default for TrendlineAddOptions<'_> {
    fn default() -> Self {
        Self {
            trendline_type: TrendlineType::LINEAR,
            order: None,
            period: None,
            forward: None,
            backward: None,
            intercept: None,
            display_equation: None,
            display_r_squared: None,
            name: None,
        }
    }
}
/// Error-bar construction options delegated to Excel.
#[derive(Debug)]
pub struct ErrorBarOptions<'a> {
    pub direction: ErrorBarDirection,
    pub include: ErrorBarInclude,
    pub error_type: ErrorBarType,
    pub amount: Option<SeriesData<'a>>,
    pub minus_values: Option<SeriesData<'a>>,
}
/// Local-picture insertion options. The path must already exist.
#[derive(Debug)]
pub struct PictureAddOptions<'a> {
    pub path: &'a Path,
    pub link_to_file: bool,
    pub save_with_document: bool,
    pub bounds: ShapeBounds,
}
/// Text-box construction options for Office `TextFrame2` text.
#[derive(Debug)]
pub struct TextBoxAddOptions<'a> {
    pub orientation: TextOrientation,
    pub bounds: ShapeBounds,
    pub text: &'a str,
}
/// Destination used when creating a chart sheet.
#[derive(Debug)]
pub enum ChartSheetDestination<'a> {
    Before(&'a Sheet),
    After(&'a Sheet),
    End,
}
