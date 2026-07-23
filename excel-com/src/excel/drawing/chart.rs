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
    /// Returns chart-type-specific groups such as column, pie, or bubble groups.
    pub fn chart_groups(&self) -> Result<ChartGroups, ExcelComError> {
        get_dispatch(
            &self.inner,
            "excel.chart.chartgroups",
            ChartGroups::from_dispatch,
        )
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
    /// Applies one of Excel's built-in chart layouts.
    pub fn apply_layout(&self, layout: i32) -> Result<(), ExcelComError> {
        call(
            &self.inner,
            "excel.chart.applylayout",
            vec![OwnedVariant::i32(layout)],
        )
    }
    /// Returns Excel's built-in chart style index.
    pub fn chart_style(&self) -> Result<i32, ExcelComError> {
        get_i32(
            &self.inner,
            "excel.chart.chartstyle",
            "Chart.ChartStyle was not an integer",
        )
    }
    /// Applies an Excel built-in chart style index.
    pub fn set_chart_style(&self, value: i32) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.chart.chartstyle",
            OwnedVariant::i32(value),
        )
    }
    /// Applies an Excel chart template at a UTF-16 Windows path.
    pub fn apply_chart_template(&self, path: &Path) -> Result<(), ExcelComError> {
        call(
            &self.inner,
            "excel.chart.applycharttemplate",
            vec![path_bstr(path)?],
        )
    }
    /// Saves this chart as an Excel chart template at a UTF-16 Windows path.
    pub fn save_chart_template(&self, path: &Path) -> Result<(), ExcelComError> {
        call(
            &self.inner,
            "excel.chart.savecharttemplate",
            vec![path_bstr(path)?],
        )
    }
    /// Returns the 3-D elevation for chart types that expose it.
    pub fn elevation(&self) -> Result<i32, ExcelComError> {
        get_i32(
            &self.inner,
            "excel.chart.elevation",
            "Chart.Elevation was not an integer",
        )
    }
    /// Sets the 3-D elevation for an applicable chart family.
    pub fn set_elevation(&self, value: i32) -> Result<(), ExcelComError> {
        if !(-90..=90).contains(&value) {
            return Err(ExcelComError::Unsupported {
                detail: "Chart.Elevation must be between -90 and 90",
            });
        }
        put(
            &self.inner,
            "excel.chart.elevation",
            OwnedVariant::i32(value),
        )
    }
    /// Returns the 3-D rotation for chart types that expose it.
    pub fn rotation(&self) -> Result<i32, ExcelComError> {
        get_i32(
            &self.inner,
            "excel.chart.rotation",
            "Chart.Rotation was not an integer",
        )
    }
    /// Sets the 3-D rotation for an applicable chart family.
    pub fn set_rotation(&self, value: i32) -> Result<(), ExcelComError> {
        if !(0..=360).contains(&value) {
            return Err(ExcelComError::Unsupported {
                detail: "Chart.Rotation must be between 0 and 360",
            });
        }
        put(
            &self.inner,
            "excel.chart.rotation",
            OwnedVariant::i32(value),
        )
    }
    /// Returns the 3-D perspective for an applicable chart family.
    pub fn perspective(&self) -> Result<i32, ExcelComError> {
        get_i32(
            &self.inner,
            "excel.chart.perspective",
            "Chart.Perspective was not an integer",
        )
    }
    /// Sets the 3-D perspective for an applicable chart family.
    pub fn set_perspective(&self, value: i32) -> Result<(), ExcelComError> {
        if !(0..=100).contains(&value) {
            return Err(ExcelComError::Unsupported {
                detail: "Chart.Perspective must be between 0 and 100",
            });
        }
        put(
            &self.inner,
            "excel.chart.perspective",
            OwnedVariant::i32(value),
        )
    }
    /// Returns whether Excel uses right-angle 3-D axes.
    pub fn right_angle_axes(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.chart.rightangleaxes")
    }
    /// Changes right-angle 3-D axes on an applicable chart family.
    pub fn set_right_angle_axes(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.chart.rightangleaxes",
            OwnedVariant::bool(value),
        )
    }
    /// Returns the 3-D height percentage for an applicable chart family.
    pub fn height_percent(&self) -> Result<i32, ExcelComError> {
        get_i32(
            &self.inner,
            "excel.chart.heightpercent",
            "Chart.HeightPercent was not an integer",
        )
    }
    /// Sets the 3-D height percentage for an applicable chart family.
    pub fn set_height_percent(&self, value: i32) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.chart.heightpercent",
            OwnedVariant::i32(value),
        )
    }
    /// Returns the 3-D depth percentage for an applicable chart family.
    pub fn depth_percent(&self) -> Result<i32, ExcelComError> {
        get_i32(
            &self.inner,
            "excel.chart.depthpercent",
            "Chart.DepthPercent was not an integer",
        )
    }
    /// Sets the 3-D depth percentage for an applicable chart family.
    pub fn set_depth_percent(&self, value: i32) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.chart.depthpercent",
            OwnedVariant::i32(value),
        )
    }
    /// Returns whether the chart auto-scales 3-D settings.
    pub fn auto_scaling(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.chart.autoscaling")
    }
    /// Changes 3-D auto-scaling on an applicable chart family.
    pub fn set_auto_scaling(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.chart.autoscaling",
            OwnedVariant::bool(value),
        )
    }
    /// Returns the 3-D gap depth for an applicable chart family.
    pub fn gap_depth(&self) -> Result<i32, ExcelComError> {
        get_i32(
            &self.inner,
            "excel.chart.gapdepth",
            "Chart.GapDepth was not an integer",
        )
    }
    /// Sets the 3-D gap depth for an applicable chart family.
    pub fn set_gap_depth(&self, value: i32) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.chart.gapdepth",
            OwnedVariant::i32(value),
        )
    }
    /// Returns the chart's 3-D walls when the selected chart type provides them.
    pub fn walls(&self) -> Result<Walls, ExcelComError> {
        get_dispatch(&self.inner, "excel.chart.walls", Walls::from_dispatch)
    }
    /// Returns the chart's 3-D floor when the selected chart type provides it.
    pub fn floor(&self) -> Result<Floor, ExcelComError> {
        get_dispatch(&self.inner, "excel.chart.floor", Floor::from_dispatch)
    }
}
