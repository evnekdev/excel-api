//! Focused drawing implementation component.
#![allow(missing_docs)]
use super::export::copy_picture;
#[allow(unused_imports)]
use super::helpers::*;
#[allow(unused_imports)]
use super::types::*;
#[allow(unused_imports)]
use super::*;
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
