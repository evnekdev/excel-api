//! Focused drawing implementation component.
#![allow(missing_docs)]
use super::export::copy_picture;
#[allow(unused_imports)]
use super::helpers::*;
#[allow(unused_imports)]
use super::types::*;
#[allow(unused_imports)]
use super::*;
/// A chart's native Excel object.
pub struct Chart {
    pub(super) inner: DispatchObject,
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
