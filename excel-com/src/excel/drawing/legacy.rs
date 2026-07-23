//! Typed, apartment-bound wrappers for Excel-native charts and drawing objects.
//!
//! The public surface deliberately models Excel and shared Office objects by
//! their concrete role. It does not expose a generic `IDispatch` escape hatch.
//! Coordinates are worksheet points, and Excel remains responsible for source
//! data interpretation, chart calculation, installed export filters, and
//! picture rendering.
#![allow(missing_docs)]

use std::fmt::{Debug, Formatter};
use std::iter::FusedIterator;
use std::path::Path;

use crate::automation::{
    AutomationArray, AutomationValue, ConversionPolicy, EnumVariant, OwnedVariant,
    PositionalArguments, decode_variant, encode_variant, enumerated_dispatch, invoke, property_get,
    property_put,
};
use crate::excel::text::text_bstr;
use crate::excel::{
    Application, DispatchObject, Font, Range, Sheet, SheetDestination, Workbook, Worksheet,
};
use crate::internal::{ComPtr, Dispatch, path_bstr};
use crate::object_model::{MemberId, member};
use crate::{ConversionError, ExcelComError};

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

fn dispatch(kind: &'static str, value: ComPtr<Dispatch>) -> DispatchObject {
    DispatchObject {
        dispatch: value,
        kind,
    }
}
fn get_dispatch<T>(
    target: &DispatchObject,
    id: &'static str,
    from: impl FnOnce(ComPtr<Dispatch>) -> T,
) -> Result<T, ExcelComError> {
    let mut result = property_get(&target.dispatch, member(MemberId::new(id), false), vec![])?;
    Ok(from(result.take_dispatch()?))
}
fn method_dispatch<T>(
    target: &DispatchObject,
    id: &'static str,
    arguments: Vec<OwnedVariant>,
    from: impl FnOnce(ComPtr<Dispatch>) -> T,
) -> Result<T, ExcelComError> {
    let mut result = invoke(
        &target.dispatch,
        member(MemberId::new(id), false),
        arguments,
        false,
    )?;
    Ok(from(result.take_dispatch()?))
}
fn optional_dispatch<T>(
    target: &DispatchObject,
    id: &'static str,
    from: impl FnOnce(ComPtr<Dispatch>) -> T,
) -> Result<Option<T>, ExcelComError> {
    let mut result = property_get(&target.dispatch, member(MemberId::new(id), false), vec![])?;
    result.take_optional_dispatch()?.map(from).pipe(Ok)
}
fn get_bool(target: &DispatchObject, id: &'static str) -> Result<bool, ExcelComError> {
    let value = property_get(&target.dispatch, member(MemberId::new(id), false), vec![])?;
    value.as_bool().ok_or(ExcelComError::Conversion(
        ConversionError::UnsupportedVariantType {
            vartype: value.vt(),
        },
    ))
}
fn get_i32(
    target: &DispatchObject,
    id: &'static str,
    detail: &'static str,
) -> Result<i32, ExcelComError> {
    let value = property_get(&target.dispatch, member(MemberId::new(id), false), vec![])?;
    value
        .as_i32()
        .or_else(|| {
            value
                .as_f64()
                .filter(|number| {
                    number.is_finite()
                        && number.fract() == 0.0
                        && *number >= i32::MIN as f64
                        && *number <= i32::MAX as f64
                })
                .map(|number| number as i32)
        })
        .ok_or(ExcelComError::Unsupported { detail })
}
fn get_f64(
    target: &DispatchObject,
    id: &'static str,
    detail: &'static str,
) -> Result<f64, ExcelComError> {
    let value = property_get(&target.dispatch, member(MemberId::new(id), false), vec![])?;
    value
        .as_f64()
        .or_else(|| value.as_i32().map(f64::from))
        .ok_or(ExcelComError::Unsupported { detail })
}
fn get_text(target: &DispatchObject, id: &'static str) -> Result<String, ExcelComError> {
    property_get(&target.dispatch, member(MemberId::new(id), false), vec![])?.as_string()
}
fn put(
    target: &DispatchObject,
    id: &'static str,
    value: OwnedVariant,
) -> Result<(), ExcelComError> {
    property_put(&target.dispatch, member(MemberId::new(id), true), value).map(|_| ())
}
fn call(
    target: &DispatchObject,
    id: &'static str,
    arguments: Vec<OwnedVariant>,
) -> Result<(), ExcelComError> {
    invoke(
        &target.dispatch,
        member(MemberId::new(id), false),
        arguments,
        false,
    )
    .map(|_| ())
}
fn finite(value: f64, detail: &'static str) -> Result<(), ExcelComError> {
    value
        .is_finite()
        .then_some(())
        .ok_or(ExcelComError::Unsupported { detail })
}
fn positive(value: f64, detail: &'static str) -> Result<(), ExcelComError> {
    finite(value, detail)?;
    (value > 0.0)
        .then_some(())
        .ok_or(ExcelComError::Unsupported { detail })
}
fn one_based(index: usize, detail: &'static str) -> Result<OwnedVariant, ExcelComError> {
    if index == 0 {
        return Err(ExcelComError::Unsupported { detail });
    }
    i32::try_from(index)
        .map(OwnedVariant::i32)
        .map_err(|_| ExcelComError::Unsupported { detail })
}
fn chart_bounds(bounds: ChartBounds) -> Result<(), ExcelComError> {
    for value in [bounds.left, bounds.top] {
        finite(value, "chart bounds must be finite")?;
    }
    positive(bounds.width, "chart width must be positive")?;
    positive(bounds.height, "chart height must be positive")
}
fn shape_bounds(bounds: ShapeBounds) -> Result<(), ExcelComError> {
    for value in [bounds.left, bounds.top, bounds.width, bounds.height] {
        finite(value, "shape geometry must be finite")?;
    }
    positive(bounds.width, "shape width must be positive")?;
    positive(bounds.height, "shape height must be positive")
}
fn shape_f32(value: f64) -> Result<OwnedVariant, ExcelComError> {
    finite(value, "shape geometry must be finite")?;
    let value = value as f32;
    value
        .is_finite()
        .then_some(OwnedVariant::f32(value))
        .ok_or(ExcelComError::Unsupported {
            detail: "shape geometry exceeds Office single precision",
        })
}
fn series_data(value: SeriesData<'_>) -> Result<OwnedVariant, ExcelComError> {
    match value {
        SeriesData::Range(range) => Ok(OwnedVariant::dispatch_borrowed(
            &range.dispatch_object().dispatch,
        )),
        SeriesData::Array(array) => encode_variant(
            &AutomationValue::Array(array.clone()),
            ConversionPolicy::default(),
        ),
        SeriesData::Formula(value) => text_bstr(value),
    }
}

trait Pipe: Sized {
    fn pipe<T>(self, value: impl FnOnce(Self) -> T) -> T {
        value(self)
    }
}
impl<T> Pipe for T {}

const CHART_OBJECTS: Collection = Collection {
    name: "ChartObjects",
    count: "excel.chartobjects.count",
    item: "excel.chartobjects.item",
    new_enum: "excel.chartobjects.newenum",
};
const CHARTS: Collection = Collection {
    name: "Charts",
    count: "excel.charts.count",
    item: "excel.charts.item",
    new_enum: "excel.charts.newenum",
};
const SERIES: Collection = Collection {
    name: "SeriesCollection",
    count: "excel.seriescollection.count",
    item: "excel.seriescollection.item",
    new_enum: "excel.seriescollection.newenum",
};
const SHAPES: Collection = Collection {
    name: "Shapes",
    count: "excel.shapes.count",
    item: "excel.shapes.item",
    new_enum: "excel.shapes.newenum",
};
const TRENDLINES: Collection = Collection {
    name: "Trendlines",
    count: "excel.trendlines.count",
    item: "excel.trendlines.item",
    new_enum: "excel.trendlines.newenum",
};
const SPARKLINE_GROUPS: Collection = Collection {
    name: "SparklineGroups",
    count: "excel.sparklinegroups.count",
    item: "excel.sparklinegroups.item",
    new_enum: "excel.sparklinegroups.newenum",
};
struct Collection {
    name: &'static str,
    count: &'static str,
    item: &'static str,
    new_enum: &'static str,
}
fn collection_count(
    target: &DispatchObject,
    descriptor: Collection,
) -> Result<usize, ExcelComError> {
    usize::try_from(get_i32(
        target,
        descriptor.count,
        "collection Count must be nonnegative",
    )?)
    .map_err(|_| ExcelComError::Unsupported {
        detail: "collection Count was negative",
    })
}
fn collection_item<T>(
    target: &DispatchObject,
    descriptor: Collection,
    index: usize,
    from: impl FnOnce(ComPtr<Dispatch>) -> T,
) -> Result<T, ExcelComError> {
    let mut value = property_get(
        &target.dispatch,
        member(MemberId::new(descriptor.item), false),
        vec![one_based(index, "collection index is one-based")?],
    )?;
    Ok(from(value.take_dispatch()?))
}
fn collection_named<T>(
    target: &DispatchObject,
    descriptor: Collection,
    name: &str,
    from: impl FnOnce(ComPtr<Dispatch>) -> T,
) -> Result<T, ExcelComError> {
    let mut value = property_get(
        &target.dispatch,
        member(MemberId::new(descriptor.item), false),
        vec![text_bstr(name)?],
    )?;
    Ok(from(value.take_dispatch()?))
}

/// An embedded-chart collection on one worksheet.
pub struct ChartObjects {
    inner: DispatchObject,
}
impl Debug for ChartObjects {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("ChartObjects").field(&self.inner).finish()
    }
}
impl Clone for ChartObjects {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl ChartObjects {
    pub(crate) fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
        Self {
            inner: dispatch("ChartObjects", value),
        }
    }
    pub fn count(&self) -> Result<usize, ExcelComError> {
        collection_count(&self.inner, CHART_OBJECTS)
    }
    pub fn item_by_index(&self, index: usize) -> Result<ChartObject, ExcelComError> {
        collection_item(
            &self.inner,
            CHART_OBJECTS,
            index,
            ChartObject::from_dispatch,
        )
    }
    pub fn item_by_name(&self, name: &str) -> Result<ChartObject, ExcelComError> {
        collection_named(&self.inner, CHART_OBJECTS, name, ChartObject::from_dispatch)
    }
    pub fn iter(&self) -> Result<ChartObjectsIter, ExcelComError> {
        Ok(ChartObjectsIter {
            inner: EnumVariant::from_new_enum(
                &self.inner,
                MemberId::new(CHART_OBJECTS.new_enum),
                CHART_OBJECTS.name,
            )?,
            index: 0,
            done: false,
        })
    }
    pub fn add(&self, bounds: ChartBounds) -> Result<ChartObject, ExcelComError> {
        chart_bounds(bounds)?;
        let mut args = PositionalArguments::new();
        for value in [bounds.left, bounds.top, bounds.width, bounds.height] {
            args.push_required(OwnedVariant::f64(value));
        }
        let mut value = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.chartobjects.add"), false),
            args.into_inner(),
            false,
        )?;
        Ok(ChartObject::from_dispatch(value.take_dispatch()?))
    }
}
pub struct ChartObjectsIter {
    inner: EnumVariant,
    index: usize,
    done: bool,
}
impl Iterator for ChartObjectsIter {
    type Item = Result<ChartObject, ExcelComError>;
    fn next(&mut self) -> Option<Self::Item> {
        if self.done {
            return None;
        }
        match self.inner.next() {
            Ok(Some(mut value)) => {
                let index = self.index;
                self.index += 1;
                Some(
                    enumerated_dispatch(&mut value, "ChartObjects", index)
                        .map(ChartObject::from_dispatch),
                )
            }
            Ok(None) => {
                self.done = true;
                None
            }
            Err(error) => {
                self.done = true;
                Some(Err(error))
            }
        }
    }
}
impl FusedIterator for ChartObjectsIter {}

/// A point-positioned embedded chart object.
pub struct ChartObject {
    inner: DispatchObject,
}
impl Debug for ChartObject {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("ChartObject").field(&self.inner).finish()
    }
}
impl Clone for ChartObject {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl ChartObject {
    pub(crate) fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
        Self {
            inner: dispatch("ChartObject", value),
        }
    }
    pub fn name(&self) -> Result<String, ExcelComError> {
        get_text(&self.inner, "excel.chartobject.name")
    }
    pub fn set_name(&self, name: &str) -> Result<(), ExcelComError> {
        put(&self.inner, "excel.chartobject.name", text_bstr(name)?)
    }
    pub fn chart(&self) -> Result<Chart, ExcelComError> {
        get_dispatch(&self.inner, "excel.chartobject.chart", Chart::from_dispatch)
    }
    pub fn left(&self) -> Result<f64, ExcelComError> {
        get_f64(
            &self.inner,
            "excel.chartobject.left",
            "ChartObject.Left was not numeric",
        )
    }
    pub fn set_left(&self, value: f64) -> Result<(), ExcelComError> {
        finite(value, "chart position must be finite")?;
        put(
            &self.inner,
            "excel.chartobject.left",
            OwnedVariant::f64(value),
        )
    }
    pub fn top(&self) -> Result<f64, ExcelComError> {
        get_f64(
            &self.inner,
            "excel.chartobject.top",
            "ChartObject.Top was not numeric",
        )
    }
    pub fn set_top(&self, value: f64) -> Result<(), ExcelComError> {
        finite(value, "chart position must be finite")?;
        put(
            &self.inner,
            "excel.chartobject.top",
            OwnedVariant::f64(value),
        )
    }
    pub fn width(&self) -> Result<f64, ExcelComError> {
        get_f64(
            &self.inner,
            "excel.chartobject.width",
            "ChartObject.Width was not numeric",
        )
    }
    pub fn set_width(&self, value: f64) -> Result<(), ExcelComError> {
        positive(value, "chart width must be positive")?;
        put(
            &self.inner,
            "excel.chartobject.width",
            OwnedVariant::f64(value),
        )
    }
    pub fn height(&self) -> Result<f64, ExcelComError> {
        get_f64(
            &self.inner,
            "excel.chartobject.height",
            "ChartObject.Height was not numeric",
        )
    }
    pub fn set_height(&self, value: f64) -> Result<(), ExcelComError> {
        positive(value, "chart height must be positive")?;
        put(
            &self.inner,
            "excel.chartobject.height",
            OwnedVariant::f64(value),
        )
    }
    pub fn visible(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.chartobject.visible")
    }
    pub fn set_visible(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.chartobject.visible",
            OwnedVariant::bool(value),
        )
    }
    pub fn placement(&self) -> Result<ShapePlacement, ExcelComError> {
        Ok(ShapePlacement::from_raw(get_i32(
            &self.inner,
            "excel.chartobject.placement",
            "ChartObject.Placement was not an integer",
        )?))
    }
    pub fn set_placement(&self, value: ShapePlacement) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.chartobject.placement",
            OwnedVariant::i32(value.raw()),
        )
    }
    pub fn activate(&self) -> Result<(), ExcelComError> {
        call(&self.inner, "excel.chartobject.activate", vec![])
    }
    pub fn copy(&self) -> Result<(), ExcelComError> {
        call(&self.inner, "excel.chartobject.copy", vec![])
    }
    pub fn copy_picture(&self, options: &CopyPictureOptions) -> Result<(), ExcelComError> {
        copy_picture(&self.inner, "excel.chartobject.copypicture", options)
    }
    pub fn delete(self) -> Result<(), ExcelComError> {
        call(&self.inner, "excel.chartobject.delete", vec![])
    }
}

impl Worksheet {
    /// Returns the worksheet's embedded chart objects.
    pub fn chart_objects(&self) -> Result<ChartObjects, ExcelComError> {
        let mut arguments = PositionalArguments::new();
        // Worksheet.ChartObjects is a method with an optional index, not a
        // property. Retaining its missing position matches the typelib.
        arguments.push_optional(None);
        method_dispatch(
            self.dispatch_object(),
            "excel.worksheet.chartobjects",
            arguments.into_inner(),
            ChartObjects::from_dispatch,
        )
    }
    /// Creates an embedded Range-backed chart and configures its optional title and legend.
    pub fn add_chart(
        &self,
        options: &ChartCreateOptions<'_>,
    ) -> Result<ChartObject, ExcelComError> {
        chart_bounds(options.bounds)?;
        if let Some(title) = options.title {
            let _ = text_bstr(title)?;
        }
        let object = self.chart_objects()?.add(options.bounds)?;
        let result = (|| {
            let chart = object.chart()?;
            chart.set_source_data(options.source, options.plot_by)?;
            chart.set_chart_type(options.chart_type)?;
            if let Some(title) = options.title {
                chart.set_has_title(true)?;
                chart
                    .chart_title()?
                    .ok_or(ExcelComError::Unsupported {
                        detail: "Excel did not create ChartTitle",
                    })?
                    .set_text(title)?;
            }
            if let Some(value) = options.has_legend {
                chart.set_has_legend(value)?;
            }
            Ok(())
        })();
        if let Err(error) = result {
            let _ = object.clone().delete();
            return Err(error);
        }
        Ok(object)
    }
    /// Returns worksheet Shapes.
    pub fn shapes(&self) -> Result<Shapes, ExcelComError> {
        get_dispatch(
            self.dispatch_object(),
            "excel.worksheet.shapes",
            Shapes::from_dispatch,
        )
    }
}

/// A chart's native Excel object.
pub struct Chart {
    inner: DispatchObject,
}
impl Debug for Chart {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("Chart").field(&self.inner).finish()
    }
}
impl Clone for Chart {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl Chart {
    pub(crate) fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
        Self {
            inner: dispatch("Chart", value),
        }
    }
    pub fn chart_type(&self) -> Result<ChartType, ExcelComError> {
        Ok(ChartType::from_raw(get_i32(
            &self.inner,
            "excel.chart.charttype",
            "Chart.ChartType was not an integer",
        )?))
    }
    pub fn set_chart_type(&self, value: ChartType) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.chart.charttype",
            OwnedVariant::i32(value.raw()),
        )
    }
    pub fn set_source_data(
        &self,
        source: &Range,
        plot_by: Option<PlotBy>,
    ) -> Result<(), ExcelComError> {
        let mut args = PositionalArguments::new();
        args.push_object(source.dispatch_object());
        args.push_optional(plot_by.map(|value| OwnedVariant::i32(value.raw())));
        call(&self.inner, "excel.chart.setsourcedata", args.into_inner())
    }
    pub fn plot_by(&self) -> Result<PlotBy, ExcelComError> {
        Ok(PlotBy::from_raw(get_i32(
            &self.inner,
            "excel.chart.plotby",
            "Chart.PlotBy was not an integer",
        )?))
    }
    pub fn set_plot_by(&self, value: PlotBy) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.chart.plotby",
            OwnedVariant::i32(value.raw()),
        )
    }
    pub fn has_title(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.chart.hastitle")
    }
    pub fn set_has_title(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.chart.hastitle",
            OwnedVariant::bool(value),
        )
    }
    pub fn chart_title(&self) -> Result<Option<ChartTitle>, ExcelComError> {
        if !self.has_title()? {
            return Ok(None);
        }
        optional_dispatch(
            &self.inner,
            "excel.chart.charttitle",
            ChartTitle::from_dispatch,
        )
    }
    pub fn has_legend(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.chart.haslegend")
    }
    pub fn set_has_legend(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.chart.haslegend",
            OwnedVariant::bool(value),
        )
    }
    pub fn legend(&self) -> Result<Option<Legend>, ExcelComError> {
        if !self.has_legend()? {
            return Ok(None);
        }
        optional_dispatch(&self.inner, "excel.chart.legend", Legend::from_dispatch)
    }
    pub fn series_collection(&self) -> Result<SeriesCollection, ExcelComError> {
        let mut value = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.chart.seriescollection"), false),
            vec![],
            false,
        )?;
        Ok(SeriesCollection::from_dispatch(value.take_dispatch()?))
    }
    /// Returns an Axis selector that supports primary and secondary axis lookup.
    pub fn axes(&self) -> Result<Axes, ExcelComError> {
        Ok(Axes {
            chart: self.inner.clone(),
        })
    }
    pub fn chart_area(&self) -> Result<ChartArea, ExcelComError> {
        get_dispatch(
            &self.inner,
            "excel.chart.chartarea",
            ChartArea::from_dispatch,
        )
    }
    pub fn plot_area(&self) -> Result<PlotArea, ExcelComError> {
        get_dispatch(&self.inner, "excel.chart.plotarea", PlotArea::from_dispatch)
    }
    pub fn refresh(&self) -> Result<(), ExcelComError> {
        call(&self.inner, "excel.chart.refresh", vec![])
    }
    pub fn export(&self, options: &ChartExportOptions<'_>) -> Result<bool, ExcelComError> {
        let mut args = PositionalArguments::new();
        args.push_result(path_bstr(options.path))?;
        match options.filter_name {
            Some(value) => args.push_result(text_bstr(value))?,
            None => args.push_optional(None),
        };
        args.push_optional(options.interactive.map(OwnedVariant::bool));
        let value = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.chart.export"), false),
            args.into_inner(),
            false,
        )?;
        value.as_bool().ok_or(ExcelComError::Conversion(
            ConversionError::UnsupportedVariantType {
                vartype: value.vt(),
            },
        ))
    }
    pub fn copy_picture(&self, options: &CopyPictureOptions) -> Result<(), ExcelComError> {
        copy_picture(&self.inner, "excel.chart.copypicture", options)
    }
    pub fn shapes(&self) -> Result<Shapes, ExcelComError> {
        get_dispatch(&self.inner, "excel.chart.shapes", Shapes::from_dispatch)
    }
}

fn copy_picture(
    target: &DispatchObject,
    id: &'static str,
    options: &CopyPictureOptions,
) -> Result<(), ExcelComError> {
    let mut args = PositionalArguments::new();
    args.push_required(OwnedVariant::i32(options.appearance.raw()));
    args.push_required(OwnedVariant::i32(options.format.raw()));
    call(target, id, args.into_inner())
}

pub struct ChartTitle {
    inner: DispatchObject,
}
impl Debug for ChartTitle {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("ChartTitle").field(&self.inner).finish()
    }
}
impl ChartTitle {
    fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
        Self {
            inner: dispatch("ChartTitle", value),
        }
    }
    pub fn text(&self) -> Result<String, ExcelComError> {
        get_text(&self.inner, "excel.charttitle.text")
    }
    pub fn set_text(&self, value: &str) -> Result<(), ExcelComError> {
        put(&self.inner, "excel.charttitle.text", text_bstr(value)?)
    }
    pub fn font(&self) -> Result<Font, ExcelComError> {
        get_dispatch(&self.inner, "excel.charttitle.font", Font::from_dispatch)
    }
}
pub struct AxisTitle {
    inner: DispatchObject,
}
impl Debug for AxisTitle {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("AxisTitle").field(&self.inner).finish()
    }
}
impl AxisTitle {
    fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
        Self {
            inner: dispatch("AxisTitle", value),
        }
    }
    pub fn text(&self) -> Result<String, ExcelComError> {
        get_text(&self.inner, "excel.axistitle.text")
    }
    pub fn set_text(&self, value: &str) -> Result<(), ExcelComError> {
        put(&self.inner, "excel.axistitle.text", text_bstr(value)?)
    }
    pub fn font(&self) -> Result<Font, ExcelComError> {
        get_dispatch(&self.inner, "excel.axistitle.font", Font::from_dispatch)
    }
}
pub struct Legend {
    inner: DispatchObject,
}
impl Debug for Legend {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("Legend").field(&self.inner).finish()
    }
}
impl Legend {
    fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
        Self {
            inner: dispatch("Legend", value),
        }
    }
    pub fn position(&self) -> Result<LegendPosition, ExcelComError> {
        Ok(LegendPosition::from_raw(get_i32(
            &self.inner,
            "excel.legend.position",
            "Legend.Position was not an integer",
        )?))
    }
    pub fn set_position(&self, value: LegendPosition) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.legend.position",
            OwnedVariant::i32(value.raw()),
        )
    }
    pub fn include_in_layout(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.legend.includeinlayout")
    }
    pub fn set_include_in_layout(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.legend.includeinlayout",
            OwnedVariant::bool(value),
        )
    }
    pub fn font(&self) -> Result<Font, ExcelComError> {
        get_dispatch(&self.inner, "excel.legend.font", Font::from_dispatch)
    }
    pub fn format(&self) -> Result<ChartFormat, ExcelComError> {
        get_dispatch(
            &self.inner,
            "excel.legend.format",
            ChartFormat::from_dispatch,
        )
    }
}
pub struct ChartArea {
    inner: DispatchObject,
}
impl Debug for ChartArea {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("ChartArea").field(&self.inner).finish()
    }
}
impl ChartArea {
    fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
        Self {
            inner: dispatch("ChartArea", value),
        }
    }
    pub fn format(&self) -> Result<ChartFormat, ExcelComError> {
        get_dispatch(
            &self.inner,
            "excel.chartarea.format",
            ChartFormat::from_dispatch,
        )
    }
}
pub struct PlotArea {
    inner: DispatchObject,
}
impl Debug for PlotArea {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("PlotArea").field(&self.inner).finish()
    }
}
impl PlotArea {
    fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
        Self {
            inner: dispatch("PlotArea", value),
        }
    }
    pub fn format(&self) -> Result<ChartFormat, ExcelComError> {
        get_dispatch(
            &self.inner,
            "excel.plotarea.format",
            ChartFormat::from_dispatch,
        )
    }
}
pub struct ChartFormat {
    inner: DispatchObject,
}
impl Debug for ChartFormat {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("ChartFormat").field(&self.inner).finish()
    }
}
impl ChartFormat {
    fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
        Self {
            inner: dispatch("ChartFormat", value),
        }
    }
    pub fn fill(&self) -> Result<FillFormat, ExcelComError> {
        get_dispatch(
            &self.inner,
            "excel.chartformat.fill",
            FillFormat::from_dispatch,
        )
    }
    pub fn line(&self) -> Result<LineFormat, ExcelComError> {
        get_dispatch(
            &self.inner,
            "excel.chartformat.line",
            LineFormat::from_dispatch,
        )
    }
}
/// Typed Office fill object; its richer Office-only member surface remains deferred.
pub struct FillFormat {
    inner: DispatchObject,
}
impl Debug for FillFormat {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("FillFormat").field(&self.inner).finish()
    }
}
impl FillFormat {
    fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
        Self {
            inner: dispatch("FillFormat", value),
        }
    }
}
/// Typed Office line object; its richer Office-only member surface remains deferred.
pub struct LineFormat {
    inner: DispatchObject,
}
impl Debug for LineFormat {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("LineFormat").field(&self.inner).finish()
    }
}
impl LineFormat {
    fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
        Self {
            inner: dispatch("LineFormat", value),
        }
    }
}

/// A chart's typed Series collection.
pub struct SeriesCollection {
    inner: DispatchObject,
}
impl Debug for SeriesCollection {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("SeriesCollection")
            .field(&self.inner)
            .finish()
    }
}
impl Clone for SeriesCollection {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl SeriesCollection {
    fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
        Self {
            inner: dispatch("SeriesCollection", value),
        }
    }
    pub fn count(&self) -> Result<usize, ExcelComError> {
        collection_count(&self.inner, SERIES)
    }
    pub fn item(&self, index: usize) -> Result<Series, ExcelComError> {
        collection_item(&self.inner, SERIES, index, Series::from_dispatch)
    }
    pub fn iter(&self) -> Result<SeriesCollectionIter, ExcelComError> {
        Ok(SeriesCollectionIter {
            inner: EnumVariant::from_new_enum(
                &self.inner,
                MemberId::new(SERIES.new_enum),
                SERIES.name,
            )?,
            index: 0,
            done: false,
        })
    }
    pub fn new_series(&self) -> Result<Series, ExcelComError> {
        let mut value = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.seriescollection.newseries"), false),
            vec![],
            false,
        )?;
        Ok(Series::from_dispatch(value.take_dispatch()?))
    }
    pub fn add_from_range(
        &self,
        source: &Range,
        options: &SeriesAddOptions,
    ) -> Result<Series, ExcelComError> {
        let mut args = PositionalArguments::new();
        args.push_object(source.dispatch_object());
        args.push_optional(options.row_col.map(|value| OwnedVariant::i32(value.raw())));
        args.push_optional(options.series_labels.map(OwnedVariant::bool));
        args.push_optional(options.category_labels.map(OwnedVariant::bool));
        args.push_optional(options.replace.map(OwnedVariant::bool));
        let mut value = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.seriescollection.add"), false),
            args.into_inner(),
            false,
        )?;
        Ok(Series::from_dispatch(value.take_dispatch()?))
    }
}
pub struct SeriesCollectionIter {
    inner: EnumVariant,
    index: usize,
    done: bool,
}
impl Iterator for SeriesCollectionIter {
    type Item = Result<Series, ExcelComError>;
    fn next(&mut self) -> Option<Self::Item> {
        if self.done {
            return None;
        }
        match self.inner.next() {
            Ok(Some(mut value)) => {
                let index = self.index;
                self.index += 1;
                Some(
                    enumerated_dispatch(&mut value, "SeriesCollection", index)
                        .map(Series::from_dispatch),
                )
            }
            Ok(None) => {
                self.done = true;
                None
            }
            Err(error) => {
                self.done = true;
                Some(Err(error))
            }
        }
    }
}
impl FusedIterator for SeriesCollectionIter {}

/// One Excel-native chart Series.
pub struct Series {
    inner: DispatchObject,
}
impl Debug for Series {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("Series").field(&self.inner).finish()
    }
}
impl Clone for Series {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl Series {
    fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
        Self {
            inner: dispatch("Series", value),
        }
    }
    pub fn name(&self) -> Result<String, ExcelComError> {
        get_text(&self.inner, "excel.series.name")
    }
    pub fn set_name(&self, value: &str) -> Result<(), ExcelComError> {
        put(&self.inner, "excel.series.name", text_bstr(value)?)
    }
    pub fn formula(&self) -> Result<String, ExcelComError> {
        get_text(&self.inner, "excel.series.formula")
    }
    pub fn set_formula(&self, value: &str) -> Result<(), ExcelComError> {
        put(&self.inner, "excel.series.formula", text_bstr(value)?)
    }
    pub fn values(&self) -> Result<AutomationValue, ExcelComError> {
        decode_variant(
            &property_get(
                &self.inner.dispatch,
                member(MemberId::new("excel.series.values"), false),
                vec![],
            )?,
            ConversionPolicy::default(),
        )
    }
    pub fn set_values(&self, value: SeriesData<'_>) -> Result<(), ExcelComError> {
        put(&self.inner, "excel.series.values", series_data(value)?)
    }
    pub fn x_values(&self) -> Result<AutomationValue, ExcelComError> {
        decode_variant(
            &property_get(
                &self.inner.dispatch,
                member(MemberId::new("excel.series.xvalues"), false),
                vec![],
            )?,
            ConversionPolicy::default(),
        )
    }
    pub fn set_x_values(&self, value: SeriesData<'_>) -> Result<(), ExcelComError> {
        put(&self.inner, "excel.series.xvalues", series_data(value)?)
    }
    pub fn axis_group(&self) -> Result<AxisGroup, ExcelComError> {
        Ok(AxisGroup::from_raw(get_i32(
            &self.inner,
            "excel.series.axisgroup",
            "Series.AxisGroup was not an integer",
        )?))
    }
    pub fn set_axis_group(&self, value: AxisGroup) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.series.axisgroup",
            OwnedVariant::i32(value.raw()),
        )
    }
    pub fn chart_type(&self) -> Result<ChartType, ExcelComError> {
        Ok(ChartType::from_raw(get_i32(
            &self.inner,
            "excel.series.charttype",
            "Series.ChartType was not an integer",
        )?))
    }
    pub fn set_chart_type(&self, value: ChartType) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.series.charttype",
            OwnedVariant::i32(value.raw()),
        )
    }
    pub fn marker_style(&self) -> Result<MarkerStyle, ExcelComError> {
        Ok(MarkerStyle::from_raw(get_i32(
            &self.inner,
            "excel.series.markerstyle",
            "Series.MarkerStyle was not an integer",
        )?))
    }
    pub fn set_marker_style(&self, value: MarkerStyle) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.series.markerstyle",
            OwnedVariant::i32(value.raw()),
        )
    }
    pub fn marker_size(&self) -> Result<i32, ExcelComError> {
        get_i32(
            &self.inner,
            "excel.series.markersize",
            "Series.MarkerSize was not an integer",
        )
    }
    pub fn set_marker_size(&self, value: i32) -> Result<(), ExcelComError> {
        if !(2..=72).contains(&value) {
            return Err(ExcelComError::Unsupported {
                detail: "Series.MarkerSize must be between 2 and 72",
            });
        }
        put(
            &self.inner,
            "excel.series.markersize",
            OwnedVariant::i32(value),
        )
    }
    pub fn format(&self) -> Result<ChartFormat, ExcelComError> {
        get_dispatch(
            &self.inner,
            "excel.series.format",
            ChartFormat::from_dispatch,
        )
    }
    pub fn apply_data_labels(&self, options: &DataLabelOptions) -> Result<(), ExcelComError> {
        if let Some(value) = &options.separator {
            let _ = text_bstr(value)?;
        }
        let mut args = PositionalArguments::new();
        args.push_optional(
            options
                .label_type
                .map(|value| OwnedVariant::i32(value.raw())),
        );
        args.push_optional(options.show_legend_key.map(OwnedVariant::bool));
        args.push_optional(None);
        args.push_optional(None);
        args.push_optional(options.show_series_name.map(OwnedVariant::bool));
        args.push_optional(options.show_category_name.map(OwnedVariant::bool));
        args.push_optional(options.show_value.map(OwnedVariant::bool));
        args.push_optional(options.show_percentage.map(OwnedVariant::bool));
        args.push_optional(None);
        match &options.separator {
            Some(value) => args.push_result(text_bstr(value))?,
            None => args.push_optional(None),
        };
        call(
            &self.inner,
            "excel.series.applydatalabels",
            args.into_inner(),
        )
    }
    pub fn data_labels(&self) -> Result<Option<DataLabels>, ExcelComError> {
        optional_dispatch(
            &self.inner,
            "excel.series.datalabels",
            DataLabels::from_dispatch,
        )
    }
    pub fn trendlines(&self) -> Result<Trendlines, ExcelComError> {
        get_dispatch(
            &self.inner,
            "excel.series.trendlines",
            Trendlines::from_dispatch,
        )
    }
    pub fn set_error_bars(&self, options: &ErrorBarOptions<'_>) -> Result<(), ExcelComError> {
        let mut args = PositionalArguments::new();
        args.push_required(OwnedVariant::i32(options.direction.raw()));
        args.push_required(OwnedVariant::i32(options.include.raw()));
        args.push_required(OwnedVariant::i32(options.error_type.raw()));
        args.push_optional(
            options
                .amount
                .as_ref()
                .map(|value| series_data_ref(value))
                .transpose()?,
        );
        args.push_optional(
            options
                .minus_values
                .as_ref()
                .map(|value| series_data_ref(value))
                .transpose()?,
        );
        call(&self.inner, "excel.series.errorbar", args.into_inner())
    }
    pub fn clear_error_bars(&self) -> Result<(), ExcelComError> {
        call(
            &self.inner,
            "excel.series.errorbar",
            vec![
                OwnedVariant::i32(ErrorBarDirection::Y.raw()),
                OwnedVariant::i32(ErrorBarInclude::NONE.raw()),
                OwnedVariant::i32(ErrorBarType::FIXED_VALUE.raw()),
            ],
        )
    }
    pub fn has_error_bars(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.series.haserrorbars")
    }
    pub fn delete(self) -> Result<(), ExcelComError> {
        call(&self.inner, "excel.series.delete", vec![])
    }
}
fn series_data_ref(value: &SeriesData<'_>) -> Result<OwnedVariant, ExcelComError> {
    match value {
        SeriesData::Range(range) => Ok(OwnedVariant::dispatch_borrowed(
            &range.dispatch_object().dispatch,
        )),
        SeriesData::Array(array) => encode_variant(
            &AutomationValue::Array((*array).clone()),
            ConversionPolicy::default(),
        ),
        SeriesData::Formula(value) => text_bstr(value),
    }
}

/// A collection of data labels returned by Excel for a Series.
pub struct DataLabels {
    inner: DispatchObject,
}
impl Debug for DataLabels {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("DataLabels").field(&self.inner).finish()
    }
}
impl DataLabels {
    fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
        Self {
            inner: dispatch("DataLabels", value),
        }
    }
    pub fn count(&self) -> Result<usize, ExcelComError> {
        collection_count(
            &self.inner,
            Collection {
                name: "DataLabels",
                count: "excel.datalabels.count",
                item: "excel.datalabels.item",
                new_enum: "excel.datalabels.newenum",
            },
        )
    }
    pub fn item(&self, index: usize) -> Result<DataLabel, ExcelComError> {
        collection_item(
            &self.inner,
            Collection {
                name: "DataLabels",
                count: "excel.datalabels.count",
                item: "excel.datalabels.item",
                new_enum: "excel.datalabels.newenum",
            },
            index,
            DataLabel::from_dispatch,
        )
    }
}
/// One Excel data label.
pub struct DataLabel {
    inner: DispatchObject,
}
impl Debug for DataLabel {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("DataLabel").field(&self.inner).finish()
    }
}
impl DataLabel {
    fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
        Self {
            inner: dispatch("DataLabel", value),
        }
    }
}

/// A Series trendline collection.
pub struct Trendlines {
    inner: DispatchObject,
}
impl Debug for Trendlines {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("Trendlines").field(&self.inner).finish()
    }
}
impl Trendlines {
    fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
        Self {
            inner: dispatch("Trendlines", value),
        }
    }
    pub fn count(&self) -> Result<usize, ExcelComError> {
        collection_count(&self.inner, TRENDLINES)
    }
    pub fn item(&self, index: usize) -> Result<Trendline, ExcelComError> {
        collection_item(&self.inner, TRENDLINES, index, Trendline::from_dispatch)
    }
    pub fn iter(&self) -> Result<TrendlinesIter, ExcelComError> {
        Ok(TrendlinesIter {
            inner: EnumVariant::from_new_enum(
                &self.inner,
                MemberId::new(TRENDLINES.new_enum),
                TRENDLINES.name,
            )?,
            index: 0,
            done: false,
        })
    }
    pub fn add(&self, options: &TrendlineAddOptions<'_>) -> Result<Trendline, ExcelComError> {
        if let Some(value) = options.order {
            if value == 0 {
                return Err(ExcelComError::Unsupported {
                    detail: "trendline order must be positive",
                });
            }
        }
        if let Some(value) = options.period {
            if value == 0 {
                return Err(ExcelComError::Unsupported {
                    detail: "trendline period must be positive",
                });
            }
        }
        for value in [options.forward, options.backward, options.intercept]
            .into_iter()
            .flatten()
        {
            finite(value, "trendline value must be finite")?;
        }
        if let Some(value) = options.name {
            let _ = text_bstr(value)?;
        }
        let mut args = PositionalArguments::new();
        args.push_required(OwnedVariant::i32(options.trendline_type.raw()));
        args.push_optional(
            options
                .order
                .map(|value| i32::try_from(value).map(OwnedVariant::i32))
                .transpose()
                .map_err(|_| ExcelComError::Unsupported {
                    detail: "trendline order exceeds i32",
                })?,
        );
        args.push_optional(
            options
                .period
                .map(|value| i32::try_from(value).map(OwnedVariant::i32))
                .transpose()
                .map_err(|_| ExcelComError::Unsupported {
                    detail: "trendline period exceeds i32",
                })?,
        );
        args.push_optional(options.forward.map(OwnedVariant::f64));
        args.push_optional(options.backward.map(OwnedVariant::f64));
        args.push_optional(options.intercept.map(OwnedVariant::f64));
        args.push_optional(options.display_equation.map(OwnedVariant::bool));
        args.push_optional(options.display_r_squared.map(OwnedVariant::bool));
        match options.name {
            Some(value) => args.push_result(text_bstr(value))?,
            None => args.push_optional(None),
        };
        let mut value = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.trendlines.add"), false),
            args.into_inner(),
            false,
        )?;
        Ok(Trendline::from_dispatch(value.take_dispatch()?))
    }
}
pub struct TrendlinesIter {
    inner: EnumVariant,
    index: usize,
    done: bool,
}
impl Iterator for TrendlinesIter {
    type Item = Result<Trendline, ExcelComError>;
    fn next(&mut self) -> Option<Self::Item> {
        if self.done {
            return None;
        }
        match self.inner.next() {
            Ok(Some(mut value)) => {
                let index = self.index;
                self.index += 1;
                Some(
                    enumerated_dispatch(&mut value, "Trendlines", index)
                        .map(Trendline::from_dispatch),
                )
            }
            Ok(None) => {
                self.done = true;
                None
            }
            Err(error) => {
                self.done = true;
                Some(Err(error))
            }
        }
    }
}
impl FusedIterator for TrendlinesIter {}
pub struct Trendline {
    inner: DispatchObject,
}
impl Debug for Trendline {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("Trendline").field(&self.inner).finish()
    }
}
impl Trendline {
    fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
        Self {
            inner: dispatch("Trendline", value),
        }
    }
    pub fn trendline_type(&self) -> Result<TrendlineType, ExcelComError> {
        Ok(TrendlineType::from_raw(get_i32(
            &self.inner,
            "excel.trendline.type",
            "Trendline.Type was not an integer",
        )?))
    }
    pub fn display_equation(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.trendline.displayequation")
    }
    pub fn set_display_equation(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.trendline.displayequation",
            OwnedVariant::bool(value),
        )
    }
    pub fn display_r_squared(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.trendline.displayrsquared")
    }
    pub fn set_display_r_squared(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.trendline.displayrsquared",
            OwnedVariant::bool(value),
        )
    }
    pub fn delete(self) -> Result<(), ExcelComError> {
        call(&self.inner, "excel.trendline.delete", vec![])
    }
}

/// Primary/secondary axis selector backed by its owning Chart.
pub struct Axes {
    chart: DispatchObject,
}
impl Debug for Axes {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("Axes").field(&self.chart).finish()
    }
}
impl Axes {
    pub fn item(
        &self,
        axis_type: AxisType,
        axis_group: Option<AxisGroup>,
    ) -> Result<Option<Axis>, ExcelComError> {
        let mut args = PositionalArguments::new();
        args.push_required(OwnedVariant::i32(axis_type.raw()));
        args.push_optional(axis_group.map(|value| OwnedVariant::i32(value.raw())));
        let mut value = invoke(
            &self.chart.dispatch,
            member(MemberId::new("excel.chart.axes"), false),
            args.into_inner(),
            false,
        )?;
        Ok(value.take_optional_dispatch()?.map(Axis::from_dispatch))
    }
}
/// A primary or secondary Excel chart Axis.
pub struct Axis {
    inner: DispatchObject,
}
impl Debug for Axis {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("Axis").field(&self.inner).finish()
    }
}
impl Axis {
    fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
        Self {
            inner: dispatch("Axis", value),
        }
    }
    pub fn has_title(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.axis.hastitle")
    }
    pub fn set_has_title(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.axis.hastitle",
            OwnedVariant::bool(value),
        )
    }
    pub fn axis_title(&self) -> Result<Option<AxisTitle>, ExcelComError> {
        if !self.has_title()? {
            return Ok(None);
        }
        optional_dispatch(
            &self.inner,
            "excel.axis.axistitle",
            AxisTitle::from_dispatch,
        )
    }
    pub fn minimum_scale(&self) -> Result<f64, ExcelComError> {
        get_f64(
            &self.inner,
            "excel.axis.minimumscale",
            "Axis.MinimumScale was not numeric",
        )
    }
    pub fn set_minimum_scale(&self, value: f64) -> Result<(), ExcelComError> {
        finite(value, "axis scale must be finite")?;
        put(
            &self.inner,
            "excel.axis.minimumscale",
            OwnedVariant::f64(value),
        )
    }
    pub fn minimum_scale_is_auto(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.axis.minimumscaleisauto")
    }
    pub fn set_minimum_scale_auto(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.axis.minimumscaleisauto",
            OwnedVariant::bool(value),
        )
    }
    pub fn maximum_scale(&self) -> Result<f64, ExcelComError> {
        get_f64(
            &self.inner,
            "excel.axis.maximumscale",
            "Axis.MaximumScale was not numeric",
        )
    }
    pub fn set_maximum_scale(&self, value: f64) -> Result<(), ExcelComError> {
        finite(value, "axis scale must be finite")?;
        put(
            &self.inner,
            "excel.axis.maximumscale",
            OwnedVariant::f64(value),
        )
    }
    pub fn maximum_scale_is_auto(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.axis.maximumscaleisauto")
    }
    pub fn set_maximum_scale_auto(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.axis.maximumscaleisauto",
            OwnedVariant::bool(value),
        )
    }
    pub fn major_unit(&self) -> Result<f64, ExcelComError> {
        get_f64(
            &self.inner,
            "excel.axis.majorunit",
            "Axis.MajorUnit was not numeric",
        )
    }
    pub fn set_major_unit(&self, value: f64) -> Result<(), ExcelComError> {
        finite(value, "axis major unit must be finite")?;
        put(
            &self.inner,
            "excel.axis.majorunit",
            OwnedVariant::f64(value),
        )
    }
    pub fn major_unit_is_auto(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.axis.majorunitisauto")
    }
    pub fn set_major_unit_auto(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.axis.majorunitisauto",
            OwnedVariant::bool(value),
        )
    }
    pub fn scale_type(&self) -> Result<AxisScaleType, ExcelComError> {
        Ok(AxisScaleType::from_raw(get_i32(
            &self.inner,
            "excel.axis.scaletype",
            "Axis.ScaleType was not an integer",
        )?))
    }
    pub fn set_scale_type(&self, value: AxisScaleType) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.axis.scaletype",
            OwnedVariant::i32(value.raw()),
        )
    }
    pub fn logarithmic_base(&self) -> Result<f64, ExcelComError> {
        get_f64(
            &self.inner,
            "excel.axis.logbase",
            "Axis.LogBase was not numeric",
        )
    }
    pub fn set_logarithmic_base(&self, value: f64) -> Result<(), ExcelComError> {
        finite(value, "axis logarithmic base must be finite")?;
        put(&self.inner, "excel.axis.logbase", OwnedVariant::f64(value))
    }
    pub fn tick_label_position(&self) -> Result<TickLabelPosition, ExcelComError> {
        Ok(TickLabelPosition::from_raw(get_i32(
            &self.inner,
            "excel.axis.ticklabelposition",
            "Axis.TickLabelPosition was not an integer",
        )?))
    }
    pub fn set_tick_label_position(&self, value: TickLabelPosition) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.axis.ticklabelposition",
            OwnedVariant::i32(value.raw()),
        )
    }
    pub fn major_tick_mark(&self) -> Result<TickMark, ExcelComError> {
        Ok(TickMark::from_raw(get_i32(
            &self.inner,
            "excel.axis.majortickmark",
            "Axis.MajorTickMark was not an integer",
        )?))
    }
    pub fn set_major_tick_mark(&self, value: TickMark) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.axis.majortickmark",
            OwnedVariant::i32(value.raw()),
        )
    }
    pub fn minor_tick_mark(&self) -> Result<TickMark, ExcelComError> {
        Ok(TickMark::from_raw(get_i32(
            &self.inner,
            "excel.axis.minortickmark",
            "Axis.MinorTickMark was not an integer",
        )?))
    }
    pub fn set_minor_tick_mark(&self, value: TickMark) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.axis.minortickmark",
            OwnedVariant::i32(value.raw()),
        )
    }
    pub fn number_format(&self) -> Result<String, ExcelComError> {
        self.tick_labels()?.number_format()
    }
    pub fn set_number_format(&self, value: &str) -> Result<(), ExcelComError> {
        self.tick_labels()?.set_number_format(value)
    }
    /// Returns the labels associated with this chart axis.
    pub fn tick_labels(&self) -> Result<TickLabels, ExcelComError> {
        get_dispatch(
            &self.inner,
            "excel.axis.ticklabels",
            TickLabels::from_dispatch,
        )
    }
}
/// Typed labels of a chart axis.
pub struct TickLabels {
    inner: DispatchObject,
}
impl Debug for TickLabels {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("TickLabels").field(&self.inner).finish()
    }
}
impl TickLabels {
    fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
        Self {
            inner: dispatch("TickLabels", value),
        }
    }
    /// Returns Excel's format string for the axis labels.
    pub fn number_format(&self) -> Result<String, ExcelComError> {
        get_text(&self.inner, "excel.ticklabels.numberformat")
    }
    /// Sets Excel's format string for the axis labels.
    pub fn set_number_format(&self, value: &str) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.ticklabels.numberformat",
            text_bstr(value)?,
        )
    }
}

/// A workbook's collection of chart sheets.
pub struct Charts {
    inner: DispatchObject,
}
impl Debug for Charts {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("Charts").field(&self.inner).finish()
    }
}
impl Clone for Charts {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl Charts {
    pub(crate) fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
        Self {
            inner: dispatch("Charts", value),
        }
    }
    pub fn count(&self) -> Result<usize, ExcelComError> {
        collection_count(&self.inner, CHARTS)
    }
    pub fn item_by_index(&self, index: usize) -> Result<ChartSheet, ExcelComError> {
        collection_item(&self.inner, CHARTS, index, ChartSheet::from_dispatch)
    }
    pub fn item_by_name(&self, name: &str) -> Result<ChartSheet, ExcelComError> {
        collection_named(&self.inner, CHARTS, name, ChartSheet::from_dispatch)
    }
    pub fn iter(&self) -> Result<ChartsIter, ExcelComError> {
        Ok(ChartsIter {
            inner: EnumVariant::from_new_enum(
                &self.inner,
                MemberId::new(CHARTS.new_enum),
                CHARTS.name,
            )?,
            index: 0,
            done: false,
        })
    }
    pub fn add(
        &self,
        destination: &ChartSheetDestination<'_>,
    ) -> Result<ChartSheet, ExcelComError> {
        let mut args = PositionalArguments::new();
        match destination {
            ChartSheetDestination::Before(sheet) => {
                args.push_object(sheet.dispatch_object());
                args.push_optional(None);
                args.push_optional(None);
            }
            ChartSheetDestination::After(sheet) => {
                args.push_optional(None);
                args.push_object(sheet.dispatch_object());
                args.push_optional(None);
            }
            ChartSheetDestination::End => {
                args.push_optional(None);
                args.push_optional(None);
                args.push_optional(None);
            }
        };
        let mut value = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.charts.add"), false),
            args.into_inner(),
            false,
        )?;
        Ok(ChartSheet::from_dispatch(value.take_dispatch()?))
    }
}
pub struct ChartsIter {
    inner: EnumVariant,
    index: usize,
    done: bool,
}
impl Iterator for ChartsIter {
    type Item = Result<ChartSheet, ExcelComError>;
    fn next(&mut self) -> Option<Self::Item> {
        if self.done {
            return None;
        }
        match self.inner.next() {
            Ok(Some(mut value)) => {
                let index = self.index;
                self.index += 1;
                Some(
                    enumerated_dispatch(&mut value, "Charts", index).map(ChartSheet::from_dispatch),
                )
            }
            Ok(None) => {
                self.done = true;
                None
            }
            Err(error) => {
                self.done = true;
                Some(Err(error))
            }
        }
    }
}
impl FusedIterator for ChartsIter {}

/// A separate chart sheet in Excel's heterogeneous Sheets collection.
pub struct ChartSheet {
    inner: DispatchObject,
}
impl Debug for ChartSheet {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("ChartSheet").field(&self.inner).finish()
    }
}
impl Clone for ChartSheet {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl ChartSheet {
    pub(crate) fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
        Self {
            inner: dispatch("Chart", value),
        }
    }
    pub(crate) fn dispatch_object(&self) -> &DispatchObject {
        &self.inner
    }
    pub fn name(&self) -> Result<String, ExcelComError> {
        get_text(&self.inner, "excel.chart.name")
    }
    pub fn set_name(&self, value: &str) -> Result<(), ExcelComError> {
        put(&self.inner, "excel.chart.name", text_bstr(value)?)
    }
    pub fn index(&self) -> Result<usize, ExcelComError> {
        usize::try_from(get_i32(
            &self.inner,
            "excel.chart.index",
            "Chart.Index was not positive",
        )?)
        .ok()
        .filter(|value| *value > 0)
        .ok_or(ExcelComError::Unsupported {
            detail: "Chart.Index was not positive",
        })
    }
    pub fn activate(&self) -> Result<(), ExcelComError> {
        call(&self.inner, "excel.chart.activate", vec![])
    }
    pub fn chart(&self) -> Result<Chart, ExcelComError> {
        Ok(Chart {
            inner: self.inner.clone(),
        })
    }
    pub fn move_to(&self, destination: &SheetDestination<'_>) -> Result<(), ExcelComError> {
        sheet_copy_move(&self.inner, "excel.chart.move", destination)
    }
    pub fn copy(&self, destination: &SheetDestination<'_>) -> Result<ChartSheet, ExcelComError> {
        let args = sheet_destination_arguments(destination);
        let mut value = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.chart.copy"), false),
            args,
            false,
        )?;
        Ok(ChartSheet::from_dispatch(value.take_dispatch()?))
    }
    pub fn delete(self) -> Result<(), ExcelComError> {
        call(&self.inner, "excel.chart.delete", vec![])
    }
}
fn sheet_destination_arguments(destination: &SheetDestination<'_>) -> Vec<OwnedVariant> {
    let mut args = PositionalArguments::new();
    match destination {
        SheetDestination::Before(value) => {
            args.push_object(value.dispatch_object());
            args.push_optional(None);
        }
        SheetDestination::After(value) => {
            args.push_optional(None);
            args.push_object(value.dispatch_object());
        }
        SheetDestination::NewWorkbook => {
            args.push_optional(None);
            args.push_optional(None);
        }
    }
    args.into_inner()
}
fn sheet_copy_move(
    target: &DispatchObject,
    id: &'static str,
    destination: &SheetDestination<'_>,
) -> Result<(), ExcelComError> {
    call(target, id, sheet_destination_arguments(destination))
}
impl Workbook {
    /// Returns the workbook's chart-sheet collection.
    pub fn charts(&self) -> Result<Charts, ExcelComError> {
        get_dispatch(
            self.dispatch_object(),
            "excel.workbook.charts",
            Charts::from_dispatch,
        )
    }
}

/// Excel Shapes on a worksheet or chart.
pub struct Shapes {
    inner: DispatchObject,
}
impl Debug for Shapes {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("Shapes").field(&self.inner).finish()
    }
}
impl Clone for Shapes {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl Shapes {
    fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
        Self {
            inner: dispatch("Shapes", value),
        }
    }
    pub fn count(&self) -> Result<usize, ExcelComError> {
        collection_count(&self.inner, SHAPES)
    }
    pub fn item_by_index(&self, index: usize) -> Result<Shape, ExcelComError> {
        collection_item(&self.inner, SHAPES, index, Shape::from_dispatch)
    }
    pub fn item_by_name(&self, name: &str) -> Result<Shape, ExcelComError> {
        collection_named(&self.inner, SHAPES, name, Shape::from_dispatch)
    }
    pub fn iter(&self) -> Result<ShapesIter, ExcelComError> {
        Ok(ShapesIter {
            inner: EnumVariant::from_new_enum(
                &self.inner,
                MemberId::new(SHAPES.new_enum),
                SHAPES.name,
            )?,
            index: 0,
            done: false,
        })
    }
    pub fn add_shape(
        &self,
        shape_type: AutoShapeType,
        bounds: ShapeBounds,
    ) -> Result<Shape, ExcelComError> {
        shape_bounds(bounds)?;
        let mut args = PositionalArguments::new();
        args.push_required(OwnedVariant::i32(shape_type.raw()));
        for value in [bounds.left, bounds.top, bounds.width, bounds.height] {
            args.push_required(shape_f32(value)?);
        }
        let mut value = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.shapes.addshape"), false),
            args.into_inner(),
            false,
        )?;
        Ok(Shape::from_dispatch(value.take_dispatch()?))
    }
    pub fn add_line(&self, start: ShapePoint, end: ShapePoint) -> Result<Shape, ExcelComError> {
        for value in [start.x, start.y, end.x, end.y] {
            finite(value, "shape geometry must be finite")?;
        }
        let mut args = PositionalArguments::new();
        for value in [start.x, start.y, end.x, end.y] {
            args.push_required(shape_f32(value)?);
        }
        let mut value = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.shapes.addline"), false),
            args.into_inner(),
            false,
        )?;
        Ok(Shape::from_dispatch(value.take_dispatch()?))
    }
    pub fn add_text_box(&self, options: &TextBoxAddOptions<'_>) -> Result<Shape, ExcelComError> {
        shape_bounds(options.bounds)?;
        let _ = text_bstr(options.text)?;
        let mut args = PositionalArguments::new();
        args.push_required(OwnedVariant::i32(options.orientation.raw()));
        for value in [
            options.bounds.left,
            options.bounds.top,
            options.bounds.width,
            options.bounds.height,
        ] {
            args.push_required(shape_f32(value)?);
        }
        let mut value = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.shapes.addtextbox"), false),
            args.into_inner(),
            false,
        )?;
        let shape = Shape::from_dispatch(value.take_dispatch()?);
        if let Some(frame) = shape.text_frame()? {
            frame.text_range()?.set_text(options.text)?;
        }
        Ok(shape)
    }
    pub fn add_picture(&self, options: &PictureAddOptions<'_>) -> Result<Shape, ExcelComError> {
        if !options.path.is_file() {
            return Err(ExcelComError::InvalidPath {
                detail: "picture path must name an existing local file",
            });
        }
        if !options.link_to_file && !options.save_with_document {
            return Err(ExcelComError::Unsupported {
                detail: "picture must be linked or saved with the workbook",
            });
        }
        shape_bounds(options.bounds)?;
        let mut args = PositionalArguments::new();
        args.push_result(path_bstr(options.path))?;
        args.push_required(OwnedVariant::i32(if options.link_to_file { -1 } else { 0 }));
        args.push_required(OwnedVariant::i32(if options.save_with_document {
            -1
        } else {
            0
        }));
        for value in [
            options.bounds.left,
            options.bounds.top,
            options.bounds.width,
            options.bounds.height,
        ] {
            args.push_required(shape_f32(value)?);
        }
        let mut value = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.shapes.addpicture"), false),
            args.into_inner(),
            false,
        )?;
        Ok(Shape::from_dispatch(value.take_dispatch()?))
    }
    pub fn add_shape_over_range(
        &self,
        shape_type: AutoShapeType,
        range: &Range,
    ) -> Result<Shape, ExcelComError> {
        self.add_shape(shape_type, range.shape_bounds()?)
    }
    pub fn add_picture_over_range(
        &self,
        path: &Path,
        range: &Range,
    ) -> Result<Shape, ExcelComError> {
        self.add_picture(&PictureAddOptions {
            path,
            link_to_file: false,
            save_with_document: true,
            bounds: range.shape_bounds()?,
        })
    }
}
pub struct ShapesIter {
    inner: EnumVariant,
    index: usize,
    done: bool,
}
impl Iterator for ShapesIter {
    type Item = Result<Shape, ExcelComError>;
    fn next(&mut self) -> Option<Self::Item> {
        if self.done {
            return None;
        }
        match self.inner.next() {
            Ok(Some(mut value)) => {
                let index = self.index;
                self.index += 1;
                Some(enumerated_dispatch(&mut value, "Shapes", index).map(Shape::from_dispatch))
            }
            Ok(None) => {
                self.done = true;
                None
            }
            Err(error) => {
                self.done = true;
                Some(Err(error))
            }
        }
    }
}
impl FusedIterator for ShapesIter {}

/// An Excel or Office shape.
pub struct Shape {
    inner: DispatchObject,
}
impl Debug for Shape {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("Shape").field(&self.inner).finish()
    }
}
impl Clone for Shape {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl Shape {
    fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
        Self {
            inner: dispatch("Shape", value),
        }
    }
    pub fn name(&self) -> Result<String, ExcelComError> {
        get_text(&self.inner, "excel.shape.name")
    }
    pub fn set_name(&self, value: &str) -> Result<(), ExcelComError> {
        put(&self.inner, "excel.shape.name", text_bstr(value)?)
    }
    pub fn shape_type(&self) -> Result<ShapeType, ExcelComError> {
        Ok(ShapeType::from_raw(get_i32(
            &self.inner,
            "excel.shape.type",
            "Shape.Type was not an integer",
        )?))
    }
    pub fn left(&self) -> Result<f64, ExcelComError> {
        get_f64(
            &self.inner,
            "excel.shape.left",
            "Shape.Left was not numeric",
        )
    }
    pub fn set_left(&self, value: f64) -> Result<(), ExcelComError> {
        finite(value, "shape position must be finite")?;
        put(&self.inner, "excel.shape.left", OwnedVariant::f64(value))
    }
    pub fn top(&self) -> Result<f64, ExcelComError> {
        get_f64(&self.inner, "excel.shape.top", "Shape.Top was not numeric")
    }
    pub fn set_top(&self, value: f64) -> Result<(), ExcelComError> {
        finite(value, "shape position must be finite")?;
        put(&self.inner, "excel.shape.top", OwnedVariant::f64(value))
    }
    pub fn width(&self) -> Result<f64, ExcelComError> {
        get_f64(
            &self.inner,
            "excel.shape.width",
            "Shape.Width was not numeric",
        )
    }
    pub fn set_width(&self, value: f64) -> Result<(), ExcelComError> {
        positive(value, "shape width must be positive")?;
        put(&self.inner, "excel.shape.width", OwnedVariant::f64(value))
    }
    pub fn height(&self) -> Result<f64, ExcelComError> {
        get_f64(
            &self.inner,
            "excel.shape.height",
            "Shape.Height was not numeric",
        )
    }
    pub fn set_height(&self, value: f64) -> Result<(), ExcelComError> {
        positive(value, "shape height must be positive")?;
        put(&self.inner, "excel.shape.height", OwnedVariant::f64(value))
    }
    pub fn rotation(&self) -> Result<f64, ExcelComError> {
        get_f64(
            &self.inner,
            "excel.shape.rotation",
            "Shape.Rotation was not numeric",
        )
    }
    pub fn set_rotation(&self, value: f64) -> Result<(), ExcelComError> {
        finite(value, "shape rotation must be finite")?;
        put(
            &self.inner,
            "excel.shape.rotation",
            OwnedVariant::f64(value),
        )
    }
    pub fn visible(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.shape.visible")
    }
    pub fn set_visible(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.shape.visible",
            OwnedVariant::bool(value),
        )
    }
    pub fn placement(&self) -> Result<ShapePlacement, ExcelComError> {
        Ok(ShapePlacement::from_raw(get_i32(
            &self.inner,
            "excel.shape.placement",
            "Shape.Placement was not an integer",
        )?))
    }
    pub fn set_placement(&self, value: ShapePlacement) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.shape.placement",
            OwnedVariant::i32(value.raw()),
        )
    }
    pub fn lock_aspect_ratio(&self) -> Result<bool, ExcelComError> {
        get_i32(
            &self.inner,
            "excel.shape.lockaspectratio",
            "Shape.LockAspectRatio was not an integer",
        )
        .map(|value| value != 0)
    }
    pub fn set_lock_aspect_ratio(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.shape.lockaspectratio",
            OwnedVariant::i32(if value { -1 } else { 0 }),
        )
    }
    pub fn fill(&self) -> Result<FillFormat, ExcelComError> {
        get_dispatch(&self.inner, "excel.shape.fill", FillFormat::from_dispatch)
    }
    pub fn line(&self) -> Result<LineFormat, ExcelComError> {
        get_dispatch(&self.inner, "excel.shape.line", LineFormat::from_dispatch)
    }
    pub fn z_order(&self, value: ZOrderCommand) -> Result<(), ExcelComError> {
        call(
            &self.inner,
            "excel.shape.zorder",
            vec![OwnedVariant::i32(value.raw())],
        )
    }
    pub fn copy(&self) -> Result<(), ExcelComError> {
        call(&self.inner, "excel.shape.copy", vec![])
    }
    pub fn cut(&self) -> Result<(), ExcelComError> {
        call(&self.inner, "excel.shape.cut", vec![])
    }
    pub fn text_frame(&self) -> Result<Option<TextFrame>, ExcelComError> {
        optional_dispatch(
            &self.inner,
            "excel.shape.textframe2",
            TextFrame::from_dispatch,
        )
    }
    pub fn delete(self) -> Result<(), ExcelComError> {
        call(&self.inner, "excel.shape.delete", vec![])
    }
}
/// A typed Office text frame returned by `Shape.TextFrame2`.
pub struct TextFrame {
    inner: DispatchObject,
}
impl Debug for TextFrame {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("TextFrame").field(&self.inner).finish()
    }
}
impl TextFrame {
    fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
        Self {
            inner: dispatch("TextFrame2", value),
        }
    }
    pub fn text_range(&self) -> Result<TextRange, ExcelComError> {
        get_dispatch(
            &self.inner,
            "office.textframe2.textrange",
            TextRange::from_dispatch,
        )
    }
}
/// Typed Office text range for a Shape text frame.
pub struct TextRange {
    inner: DispatchObject,
}
impl Debug for TextRange {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("TextRange").field(&self.inner).finish()
    }
}
impl TextRange {
    fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
        Self {
            inner: dispatch("TextRange2", value),
        }
    }
    pub fn text(&self) -> Result<String, ExcelComError> {
        get_text(&self.inner, "office.textrange2.text")
    }
    pub fn set_text(&self, value: &str) -> Result<(), ExcelComError> {
        put(&self.inner, "office.textrange2.text", text_bstr(value)?)
    }
}

impl Range {
    /// Returns point-based bounds derived from Excel's Range geometry.
    pub fn shape_bounds(&self) -> Result<ShapeBounds, ExcelComError> {
        let left = get_f64(
            self.dispatch_object(),
            "excel.range.left",
            "Range.Left was not numeric",
        )?;
        let top = get_f64(
            self.dispatch_object(),
            "excel.range.top",
            "Range.Top was not numeric",
        )?;
        let width = get_f64(
            self.dispatch_object(),
            "excel.range.width",
            "Range.Width was not numeric",
        )?;
        let height = get_f64(
            self.dispatch_object(),
            "excel.range.height",
            "Range.Height was not numeric",
        )?;
        let result = ShapeBounds {
            left,
            top,
            width,
            height,
        };
        shape_bounds(result)?;
        Ok(result)
    }
    /// Copies this Range as an Excel-native picture. Clear `CutCopyMode` after controlled use.
    pub fn copy_picture(&self, options: &CopyPictureOptions) -> Result<(), ExcelComError> {
        copy_picture(self.dispatch_object(), "excel.range.copypicture", options)
    }
    /// Returns cell-bound Sparkline groups associated with this Range.
    pub fn sparkline_groups(&self) -> Result<SparklineGroups, ExcelComError> {
        get_dispatch(
            self.dispatch_object(),
            "excel.range.sparklinegroups",
            SparklineGroups::from_dispatch,
        )
    }
}
impl Application {
    /// Returns Excel's current copy/cut state without reading the system clipboard.
    pub fn cut_copy_mode(&self) -> Result<CutCopyMode, ExcelComError> {
        Ok(CutCopyMode::from_raw(get_i32(
            self.dispatch_object(),
            "excel.application.cutcopymode",
            "Application.CutCopyMode was not an integer",
        )?))
    }
    /// Clears Excel's copy/cut state without extracting arbitrary clipboard contents.
    pub fn clear_cut_copy_mode(&self) -> Result<(), ExcelComError> {
        put(
            self.dispatch_object(),
            "excel.application.cutcopymode",
            OwnedVariant::bool(false),
        )
    }
}

/// A cell-bound Excel SparklineGroups collection.
pub struct SparklineGroups {
    inner: DispatchObject,
}
impl Debug for SparklineGroups {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("SparklineGroups").field(&self.inner).finish()
    }
}
impl SparklineGroups {
    fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
        Self {
            inner: dispatch("SparklineGroups", value),
        }
    }
    pub fn count(&self) -> Result<usize, ExcelComError> {
        collection_count(&self.inner, SPARKLINE_GROUPS)
    }
    pub fn item(&self, index: usize) -> Result<SparklineGroup, ExcelComError> {
        collection_item(
            &self.inner,
            SPARKLINE_GROUPS,
            index,
            SparklineGroup::from_dispatch,
        )
    }
    pub fn iter(&self) -> Result<SparklineGroupsIter, ExcelComError> {
        Ok(SparklineGroupsIter {
            inner: EnumVariant::from_new_enum(
                &self.inner,
                MemberId::new(SPARKLINE_GROUPS.new_enum),
                SPARKLINE_GROUPS.name,
            )?,
            index: 0,
            done: false,
        })
    }
    /// Adds a cell-bound group using Excel source and location addresses.
    pub fn add(
        &self,
        sparkline_type: SparklineType,
        source_data: &Range,
        location: &Range,
    ) -> Result<SparklineGroup, ExcelComError> {
        let mut args = PositionalArguments::new();
        args.push_required(OwnedVariant::i32(sparkline_type.raw()));
        args.push_result(text_bstr(&source_data.address_a1()?))?;
        let mut value = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.sparklinegroups.add"), false),
            args.into_inner(),
            false,
        )?;
        let group = SparklineGroup::from_dispatch(value.take_dispatch()?);
        group.set_location(location)?;
        Ok(group)
    }
}
pub struct SparklineGroupsIter {
    inner: EnumVariant,
    index: usize,
    done: bool,
}
impl Iterator for SparklineGroupsIter {
    type Item = Result<SparklineGroup, ExcelComError>;
    fn next(&mut self) -> Option<Self::Item> {
        if self.done {
            return None;
        }
        match self.inner.next() {
            Ok(Some(mut value)) => {
                let index = self.index;
                self.index += 1;
                Some(
                    enumerated_dispatch(&mut value, "SparklineGroups", index)
                        .map(SparklineGroup::from_dispatch),
                )
            }
            Ok(None) => {
                self.done = true;
                None
            }
            Err(error) => {
                self.done = true;
                Some(Err(error))
            }
        }
    }
}
impl FusedIterator for SparklineGroupsIter {}
/// A group of cell-bound Excel sparklines.
pub struct SparklineGroup {
    inner: DispatchObject,
}
impl Debug for SparklineGroup {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("SparklineGroup").field(&self.inner).finish()
    }
}
impl SparklineGroup {
    fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
        Self {
            inner: dispatch("SparklineGroup", value),
        }
    }
    pub fn sparkline_type(&self) -> Result<SparklineType, ExcelComError> {
        Ok(SparklineType::from_raw(get_i32(
            &self.inner,
            "excel.sparklinegroup.type",
            "SparklineGroup.Type was not an integer",
        )?))
    }
    pub fn set_sparkline_type(&self, value: SparklineType) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.sparklinegroup.type",
            OwnedVariant::i32(value.raw()),
        )
    }
    pub fn source_data(&self) -> Result<String, ExcelComError> {
        get_text(&self.inner, "excel.sparklinegroup.sourcedata")
    }
    pub fn set_source_data(&self, source: &Range) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.sparklinegroup.sourcedata",
            text_bstr(&source.address_a1()?)?,
        )
    }
    fn set_location(&self, location: &Range) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.sparklinegroup.location",
            OwnedVariant::dispatch_borrowed(&location.dispatch_object().dispatch),
        )
    }
    pub fn delete(self) -> Result<(), ExcelComError> {
        call(&self.inner, "excel.sparklinegroup.delete", vec![])
    }
}

// The facade retains stable public exports. Focused subsystems follow it;
// `tests` is intentionally separate so test-only implementation detail does
// not inflate the production module.
#[cfg(test)]
#[path = "tests.rs"]
mod tests;
