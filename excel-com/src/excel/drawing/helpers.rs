//! Shared private drawing conversion and collection helpers.
#![allow(missing_docs)]
use super::types::*;
pub(super) fn dispatch(kind: &'static str, value: ComPtr<Dispatch>) -> DispatchObject {
    DispatchObject {
        dispatch: value,
        kind,
    }
}
pub(super) fn get_dispatch<T>(
    target: &DispatchObject,
    id: &'static str,
    from: impl FnOnce(ComPtr<Dispatch>) -> T,
) -> Result<T, ExcelComError> {
    let mut result = property_get(&target.dispatch, member(MemberId::new(id), false), vec![])?;
    Ok(from(result.take_dispatch()?))
}
pub(super) fn method_dispatch<T>(
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
pub(super) fn optional_dispatch<T>(
    target: &DispatchObject,
    id: &'static str,
    from: impl FnOnce(ComPtr<Dispatch>) -> T,
) -> Result<Option<T>, ExcelComError> {
    let mut result = property_get(&target.dispatch, member(MemberId::new(id), false), vec![])?;
    result.take_optional_dispatch()?.map(from).pipe(Ok)
}
pub(super) fn get_bool(target: &DispatchObject, id: &'static str) -> Result<bool, ExcelComError> {
    let value = property_get(&target.dispatch, member(MemberId::new(id), false), vec![])?;
    value.as_bool().ok_or(ExcelComError::Conversion(
        ConversionError::UnsupportedVariantType {
            vartype: value.vt(),
        },
    ))
}
pub(super) fn get_i32(
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
pub(super) fn get_f64(
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
pub(super) fn get_text(target: &DispatchObject, id: &'static str) -> Result<String, ExcelComError> {
    property_get(&target.dispatch, member(MemberId::new(id), false), vec![])?.as_string()
}
pub(super) fn put(
    target: &DispatchObject,
    id: &'static str,
    value: OwnedVariant,
) -> Result<(), ExcelComError> {
    property_put(&target.dispatch, member(MemberId::new(id), true), value).map(|_| ())
}
pub(super) fn call(
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
pub(super) fn finite(value: f64, detail: &'static str) -> Result<(), ExcelComError> {
    value
        .is_finite()
        .then_some(())
        .ok_or(ExcelComError::Unsupported { detail })
}
pub(super) fn positive(value: f64, detail: &'static str) -> Result<(), ExcelComError> {
    finite(value, detail)?;
    (value > 0.0)
        .then_some(())
        .ok_or(ExcelComError::Unsupported { detail })
}
pub(super) fn one_based(index: usize, detail: &'static str) -> Result<OwnedVariant, ExcelComError> {
    if index == 0 {
        return Err(ExcelComError::Unsupported { detail });
    }
    i32::try_from(index)
        .map(OwnedVariant::i32)
        .map_err(|_| ExcelComError::Unsupported { detail })
}
pub(super) fn chart_bounds(bounds: ChartBounds) -> Result<(), ExcelComError> {
    for value in [bounds.left, bounds.top] {
        finite(value, "chart bounds must be finite")?;
    }
    positive(bounds.width, "chart width must be positive")?;
    positive(bounds.height, "chart height must be positive")
}
pub(super) fn shape_bounds(bounds: ShapeBounds) -> Result<(), ExcelComError> {
    for value in [bounds.left, bounds.top, bounds.width, bounds.height] {
        finite(value, "shape geometry must be finite")?;
    }
    positive(bounds.width, "shape width must be positive")?;
    positive(bounds.height, "shape height must be positive")
}
pub(super) fn shape_f32(value: f64) -> Result<OwnedVariant, ExcelComError> {
    finite(value, "shape geometry must be finite")?;
    let value = value as f32;
    value
        .is_finite()
        .then_some(OwnedVariant::f32(value))
        .ok_or(ExcelComError::Unsupported {
            detail: "shape geometry exceeds Office single precision",
        })
}
pub(super) fn series_data(value: SeriesData<'_>) -> Result<OwnedVariant, ExcelComError> {
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

pub(super) trait Pipe: Sized {
    fn pipe<T>(self, value: impl FnOnce(Self) -> T) -> T {
        value(self)
    }
}
impl<T> Pipe for T {}

pub(super) const CHART_OBJECTS: Collection = Collection {
    name: "ChartObjects",
    count: "excel.chartobjects.count",
    item: "excel.chartobjects.item",
    new_enum: "excel.chartobjects.newenum",
};
pub(super) const CHARTS: Collection = Collection {
    name: "Charts",
    count: "excel.charts.count",
    item: "excel.charts.item",
    new_enum: "excel.charts.newenum",
};
pub(super) const SERIES: Collection = Collection {
    name: "SeriesCollection",
    count: "excel.seriescollection.count",
    item: "excel.seriescollection.item",
    new_enum: "excel.seriescollection.newenum",
};
pub(super) const SHAPES: Collection = Collection {
    name: "Shapes",
    count: "excel.shapes.count",
    item: "excel.shapes.item",
    new_enum: "excel.shapes.newenum",
};
pub(super) const TRENDLINES: Collection = Collection {
    name: "Trendlines",
    count: "excel.trendlines.count",
    item: "excel.trendlines.item",
    new_enum: "excel.trendlines.newenum",
};
pub(super) const SPARKLINE_GROUPS: Collection = Collection {
    name: "SparklineGroups",
    count: "excel.sparklinegroups.count",
    item: "excel.sparklinegroups.item",
    new_enum: "excel.sparklinegroups.newenum",
};
pub(super) struct Collection {
    pub(super) name: &'static str,
    pub(super) count: &'static str,
    pub(super) item: &'static str,
    pub(super) new_enum: &'static str,
}
pub(super) fn collection_count(
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
pub(super) fn collection_item<T>(
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
pub(super) fn collection_named<T>(
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
