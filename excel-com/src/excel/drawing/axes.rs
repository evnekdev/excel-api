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
}
