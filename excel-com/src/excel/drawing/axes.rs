//! Focused drawing implementation component.
#![allow(missing_docs)]
#[allow(unused_imports)]
use super::helpers::*;
#[allow(unused_imports)]
use super::types::*;
#[allow(unused_imports)]
use super::*;
/// Primary/secondary axis selector backed by its owning Chart.
pub struct Axes {
    pub(super) chart: DispatchObject,
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
    pub(super) fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
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
    /// Returns the manually configured minor unit for an applicable value axis.
    pub fn minor_unit(&self) -> Result<f64, ExcelComError> {
        get_f64(
            &self.inner,
            "excel.axis.minorunit",
            "Axis.MinorUnit was not numeric",
        )
    }
    /// Sets a finite manual minor unit for an applicable value axis.
    pub fn set_minor_unit(&self, value: f64) -> Result<(), ExcelComError> {
        finite(value, "axis minor unit must be finite")?;
        put(
            &self.inner,
            "excel.axis.minorunit",
            OwnedVariant::f64(value),
        )
    }
    /// Returns whether Excel automatically selects the minor unit.
    pub fn minor_unit_is_auto(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.axis.minorunitisauto")
    }
    /// Restores automatic minor-unit selection for an applicable axis.
    pub fn set_minor_unit_auto(&self) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.axis.minorunitisauto",
            OwnedVariant::bool(true),
        )
    }
    /// Returns this axis's crossing mode.
    pub fn crosses(&self) -> Result<AxisCrosses, ExcelComError> {
        Ok(AxisCrosses::from_raw(get_i32(
            &self.inner,
            "excel.axis.crosses",
            "Axis.Crosses was not an integer",
        )?))
    }
    /// Sets this axis's crossing mode.
    pub fn set_crosses(&self, value: AxisCrosses) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.axis.crosses",
            OwnedVariant::i32(value.raw()),
        )
    }
    /// Returns the custom crossing value for an applicable axis.
    pub fn crosses_at(&self) -> Result<f64, ExcelComError> {
        get_f64(
            &self.inner,
            "excel.axis.crossesat",
            "Axis.CrossesAt was not numeric",
        )
    }
    /// Sets a finite custom crossing value for an applicable axis.
    pub fn set_crosses_at(&self, value: f64) -> Result<(), ExcelComError> {
        finite(value, "axis crossing must be finite")?;
        put(
            &self.inner,
            "excel.axis.crossesat",
            OwnedVariant::f64(value),
        )
    }
    /// Returns whether this axis reverses plot order.
    pub fn reverse_plot_order(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.axis.reverseplotorder")
    }
    /// Changes plot-order reversal for this axis.
    pub fn set_reverse_plot_order(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.axis.reverseplotorder",
            OwnedVariant::bool(value),
        )
    }
    /// Returns category handling for an applicable category axis.
    pub fn category_type(&self) -> Result<CategoryType, ExcelComError> {
        Ok(CategoryType::from_raw(get_i32(
            &self.inner,
            "excel.axis.categorytype",
            "Axis.CategoryType was not an integer",
        )?))
    }
    /// Sets category handling for an applicable category axis.
    pub fn set_category_type(&self, value: CategoryType) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.axis.categorytype",
            OwnedVariant::i32(value.raw()),
        )
    }
    /// Returns the base time unit for an applicable time-scale axis.
    pub fn base_unit(&self) -> Result<TimeUnit, ExcelComError> {
        Ok(TimeUnit::from_raw(get_i32(
            &self.inner,
            "excel.axis.baseunit",
            "Axis.BaseUnit was not an integer",
        )?))
    }
    /// Sets the base time unit for an applicable time-scale axis.
    pub fn set_base_unit(&self, value: TimeUnit) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.axis.baseunit",
            OwnedVariant::i32(value.raw()),
        )
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
    /// Returns Office drawing formatting for this axis.
    pub fn format(&self) -> Result<ChartFormat, ExcelComError> {
        get_dispatch(&self.inner, "excel.axis.format", ChartFormat::from_dispatch)
    }
    /// Returns major gridlines when the axis currently exposes them.
    pub fn major_gridlines(&self) -> Result<Option<Gridlines>, ExcelComError> {
        optional_dispatch(
            &self.inner,
            "excel.axis.majorgridlines",
            Gridlines::from_dispatch,
        )
    }
    /// Returns minor gridlines when the axis currently exposes them.
    pub fn minor_gridlines(&self) -> Result<Option<Gridlines>, ExcelComError> {
        optional_dispatch(
            &self.inner,
            "excel.axis.minorgridlines",
            Gridlines::from_dispatch,
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
    pub(super) fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
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
    /// Returns Excel's raw tick-label orientation value.
    pub fn orientation(&self) -> Result<i32, ExcelComError> {
        get_i32(
            &self.inner,
            "excel.ticklabels.orientation",
            "TickLabels.Orientation was not an integer",
        )
    }
    /// Sets Excel's raw tick-label orientation value.
    pub fn set_orientation(&self, value: i32) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.ticklabels.orientation",
            OwnedVariant::i32(value),
        )
    }
    /// Returns the tick-label offset percentage.
    pub fn offset(&self) -> Result<i32, ExcelComError> {
        get_i32(
            &self.inner,
            "excel.ticklabels.offset",
            "TickLabels.Offset was not an integer",
        )
    }
    /// Sets the tick-label offset percentage from 0 through 1000.
    pub fn set_offset(&self, value: i32) -> Result<(), ExcelComError> {
        if !(0..=1000).contains(&value) {
            return Err(ExcelComError::Unsupported {
                detail: "TickLabels.Offset must be between 0 and 1000",
            });
        }
        put(
            &self.inner,
            "excel.ticklabels.offset",
            OwnedVariant::i32(value),
        )
    }
    /// Returns the legacy Excel font for these tick labels.
    pub fn font(&self) -> Result<Font, ExcelComError> {
        get_dispatch(&self.inner, "excel.ticklabels.font", Font::from_dispatch)
    }
}

/// Major or minor chart gridlines for an applicable axis.
pub struct Gridlines {
    inner: DispatchObject,
}
impl Debug for Gridlines {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_tuple("Gridlines")
            .field(&self.inner)
            .finish()
    }
}
impl Gridlines {
    pub(super) fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
        Self {
            inner: dispatch("Gridlines", value),
        }
    }
    /// Returns Office drawing formatting for these gridlines.
    pub fn format(&self) -> Result<ChartFormat, ExcelComError> {
        get_dispatch(
            &self.inner,
            "excel.gridlines.format",
            ChartFormat::from_dispatch,
        )
    }
    /// Returns the legacy Excel border for these gridlines.
    pub fn border(&self) -> Result<Border, ExcelComError> {
        get_dispatch(&self.inner, "excel.gridlines.border", Border::from_dispatch)
    }
}
