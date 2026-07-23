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
    Application, Border, DispatchObject, ExcelColor, Font, Range, Sheet, SheetDestination,
    ThemeColor, Workbook, Worksheet,
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
    /// `xl3DArea`.
    AREA_3D = -4098;
    /// `xl3DAreaStacked`.
    AREA_3D_STACKED = 78;
    /// `xl3DAreaStacked100`.
    AREA_3D_STACKED_100 = 79;
    /// `xlArea`.
    AREA = 1;
    /// `xlAreaStacked`.
    AREA_STACKED = 76;
    /// `xlAreaStacked100`.
    AREA_STACKED_100 = 77;
    /// `xl3DBarClustered`.
    BAR_3D_CLUSTERED = 60;
    /// `xl3DBarStacked`.
    BAR_3D_STACKED = 61;
    /// `xl3DBarStacked100`.
    BAR_3D_STACKED_100 = 62;
    /// `xlBarClustered`.
    BAR_CLUSTERED = 57;
    /// `xlBarStacked`.
    BAR_STACKED = 58;
    /// `xlBarStacked100`.
    BAR_STACKED_100 = 59;
    /// `xlBubble`.
    BUBBLE = 15;
    /// `xlBubble3DEffect`.
    BUBBLE_3D_EFFECT = 87;
    /// `xl3DColumn`.
    COLUMN_3D = -4100;
    /// `xl3DColumnClustered`.
    COLUMN_3D_CLUSTERED = 54;
    /// `xl3DColumnStacked`.
    COLUMN_3D_STACKED = 55;
    /// `xl3DColumnStacked100`.
    COLUMN_3D_STACKED_100 = 56;
    /// `xl3DLine`.
    LINE_3D = -4101;
    /// `xl3DPie`.
    PIE_3D = -4102;
    /// `xl3DPieExploded`.
    PIE_3D_EXPLODED = 70;
    /// `xlColumnClustered`.
    COLUMN_CLUSTERED = 51;
    /// `xlColumnStacked`.
    COLUMN_STACKED = 52;
    /// `xlColumnStacked100`.
    COLUMN_STACKED_100 = 53;
    /// `xlConeBarClustered`.
    CONE_BAR_CLUSTERED = 102;
    /// `xlConeBarStacked`.
    CONE_BAR_STACKED = 103;
    /// `xlConeBarStacked100`.
    CONE_BAR_STACKED_100 = 104;
    /// `xlConeCol`.
    CONE_COLUMN_3D = 105;
    /// `xlConeColClustered`.
    CONE_COLUMN_CLUSTERED = 99;
    /// `xlConeColStacked`.
    CONE_COLUMN_STACKED = 100;
    /// `xlConeColStacked100`.
    CONE_COLUMN_STACKED_100 = 101;
    /// `xlCylinderBarClustered`.
    CYLINDER_BAR_CLUSTERED = 95;
    /// `xlCylinderBarStacked`.
    CYLINDER_BAR_STACKED = 96;
    /// `xlCylinderBarStacked100`.
    CYLINDER_BAR_STACKED_100 = 97;
    /// `xlCylinderCol`.
    CYLINDER_COLUMN_3D = 98;
    /// `xlCylinderColClustered`.
    CYLINDER_COLUMN_CLUSTERED = 92;
    /// `xlCylinderColStacked`.
    CYLINDER_COLUMN_STACKED = 93;
    /// `xlCylinderColStacked100`.
    CYLINDER_COLUMN_STACKED_100 = 94;
    /// `xlDoughnut`.
    DOUGHNUT = -4120;
    /// `xlDoughnutExploded`.
    DOUGHNUT_EXPLODED = 80;
    /// `xlLine`.
    LINE = 4;
    /// `xlLineMarkers`.
    LINE_MARKERS = 65;
    /// `xlLineMarkersStacked`.
    LINE_MARKERS_STACKED = 66;
    /// `xlLineMarkersStacked100`.
    LINE_MARKERS_STACKED_100 = 67;
    /// `xlLineStacked`.
    LINE_STACKED = 63;
    /// `xlLineStacked100`.
    LINE_STACKED_100 = 64;
    /// `xlPie`.
    PIE = 5;
    /// `xlPieExploded`.
    PIE_EXPLODED = 69;
    /// `xlPieOfPie`.
    PIE_OF_PIE = 68;
    /// `xlBarOfPie`.
    BAR_OF_PIE = 71;
    /// `xlPyramidBarClustered`.
    PYRAMID_BAR_CLUSTERED = 109;
    /// `xlPyramidBarStacked`.
    PYRAMID_BAR_STACKED = 110;
    /// `xlPyramidBarStacked100`.
    PYRAMID_BAR_STACKED_100 = 111;
    /// `xlPyramidCol`.
    PYRAMID_COLUMN_3D = 112;
    /// `xlPyramidColClustered`.
    PYRAMID_COLUMN_CLUSTERED = 106;
    /// `xlPyramidColStacked`.
    PYRAMID_COLUMN_STACKED = 107;
    /// `xlPyramidColStacked100`.
    PYRAMID_COLUMN_STACKED_100 = 108;
    /// `xlRadar`.
    RADAR = -4151;
    /// `xlRadarFilled`.
    RADAR_FILLED = 82;
    /// `xlRadarMarkers`.
    RADAR_MARKERS = 81;
    /// `xlStockHLC`.
    STOCK_HLC = 88;
    /// `xlStockOHLC`.
    STOCK_OHLC = 89;
    /// `xlStockVHLC`.
    STOCK_VHLC = 90;
    /// `xlStockVOHLC`.
    STOCK_VOHLC = 91;
    /// `xlSurface`.
    SURFACE = 83;
    /// `xlSurfaceWireframe`.
    SURFACE_WIREFRAME = 84;
    /// `xlSurfaceTopView`.
    SURFACE_TOP_VIEW = 85;
    /// `xlSurfaceTopViewWireframe`.
    SURFACE_TOP_VIEW_WIREFRAME = 86;
    /// `xlXYScatter`.
    XY_SCATTER = -4169;
    /// `xlXYScatterLines`.
    XY_SCATTER_LINES = 74;
    /// `xlXYScatterLinesNoMarkers`.
    XY_SCATTER_LINES_NO_MARKERS = 75;
    /// `xlXYScatterSmooth`.
    XY_SCATTER_SMOOTH = 72;
    /// `xlXYScatterSmoothNoMarkers`.
    XY_SCATTER_SMOOTH_NO_MARKERS = 73;
    /// `xlTreemap`.
    TREEMAP = 117;
    /// `xlHistogram`.
    HISTOGRAM = 118;
    /// `xlWaterfall`.
    WATERFALL = 119;
    /// `xlSunburst`.
    SUNBURST = 120;
    /// `xlBoxwhisker`.
    BOX_AND_WHISKER = 121;
    /// `xlPareto`.
    PARETO = 122;
    /// `xlFunnel`.
    FUNNEL = 123;
    /// `xlRegionMap`.
    REGION_MAP = 140;
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
/// A forward-compatible `XlAxisCrosses` value.
AxisCrosses {
    /// Excel chooses the crossing automatically.
    AUTOMATIC = -4105;
    /// Cross at the minimum axis value.
    MINIMUM = 4;
    /// Cross at the maximum axis value.
    MAXIMUM = 2;
    /// Cross at a custom numeric value.
    CUSTOM = -4114;
} }
drawing_value! {
/// A forward-compatible `XlCategoryType` value.
CategoryType {
    /// Excel chooses category behavior automatically.
    AUTOMATIC = -4105;
    /// Treat categories as text.
    TEXT = 2;
    /// Treat categories as date values.
    TIME_SCALE = 3;
} }
drawing_value! {
/// A forward-compatible `XlTimeUnit` value.
TimeUnit {
    /// Days.
    DAYS = 0;
    /// Months.
    MONTHS = 1;
    /// Years.
    YEARS = 2;
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
    /// Plus marker.
    PLUS = 9;
    /// Star marker.
    STAR = 5;
    /// X marker.
    X = -4168;
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
/// Point-based geometry reported by Excel for a chart element.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct ChartElementBounds {
    /// Left edge in points.
    pub left: f64,
    /// Top edge in points.
    pub top: f64,
    /// Width in points.
    pub width: f64,
    /// Height in points.
    pub height: f64,
}

/// A chart colour that preserves direct RGB, theme, and automatic modes.
#[derive(Clone, Copy, Debug, PartialEq)]
pub enum ChartColor {
    /// Direct Excel RGB colour.
    Rgb(crate::ExcelColor),
    /// Theme colour with optional Excel tint/shade adjustment.
    Theme {
        /// Excel theme colour index.
        color: crate::ThemeColor,
        /// Excel tint-and-shade value.
        tint_and_shade: Option<f64>,
    },
    /// Excel's automatic colour choice.
    Automatic,
}

/// Optional legacy Series marker properties.
///
/// Excel exposes marker foreground and background colours as legacy colour
/// properties; this is intentionally separate from Office `FillFormat`.
#[derive(Clone, Copy, Debug, Default, PartialEq)]
pub struct MarkerFormat {
    /// Marker shape.
    pub style: Option<MarkerStyle>,
    /// Marker size between 2 and 72 points.
    pub size: Option<i32>,
    /// Legacy marker foreground colour.
    pub foreground_color: Option<ChartColor>,
    /// Legacy marker background colour.
    pub background_color: Option<ChartColor>,
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
