//! Focused drawing implementation component.
#![allow(missing_docs)]
#[allow(unused_imports)]
use super::helpers::*;
#[allow(unused_imports)]
use super::types::*;
#[allow(unused_imports)]
use super::*;
pub struct Legend {
    inner: DispatchObject,
}
impl Debug for Legend {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("Legend").field(&self.inner).finish()
    }
}
impl Legend {
    pub(super) fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
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
    pub(super) fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
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
    /// Returns chart-area geometry in points.
    pub fn bounds(&self) -> Result<ChartElementBounds, ExcelComError> {
        chart_element_bounds(&self.inner, "excel.chartarea")
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
    pub(super) fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
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
    /// Returns plot-area geometry in points.
    pub fn bounds(&self) -> Result<ChartElementBounds, ExcelComError> {
        chart_element_bounds(&self.inner, "excel.plotarea")
    }
    /// Returns plot-area inner geometry in points, where Excel exposes it.
    pub fn inside_bounds(&self) -> Result<ChartElementBounds, ExcelComError> {
        Ok(ChartElementBounds {
            left: get_f64(
                &self.inner,
                "excel.plotarea.insideleft",
                "PlotArea.InsideLeft was not numeric",
            )?,
            top: get_f64(
                &self.inner,
                "excel.plotarea.insidetop",
                "PlotArea.InsideTop was not numeric",
            )?,
            width: get_f64(
                &self.inner,
                "excel.plotarea.insidewidth",
                "PlotArea.InsideWidth was not numeric",
            )?,
            height: get_f64(
                &self.inner,
                "excel.plotarea.insideheight",
                "PlotArea.InsideHeight was not numeric",
            )?,
        })
    }
}
/// Office drawing-format entry point for a chart element.
pub struct ChartFormat {
    inner: DispatchObject,
}
impl Debug for ChartFormat {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("ChartFormat").field(&self.inner).finish()
    }
}
impl ChartFormat {
    pub(super) fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
        Self {
            inner: dispatch("ChartFormat", value),
        }
    }
    /// Returns the Office fill-format surface.
    pub fn fill(&self) -> Result<FillFormat, ExcelComError> {
        get_dispatch(
            &self.inner,
            "excel.chartformat.fill",
            FillFormat::from_dispatch,
        )
    }
    /// Returns the Office line-format surface.
    pub fn line(&self) -> Result<LineFormat, ExcelComError> {
        get_dispatch(
            &self.inner,
            "excel.chartformat.line",
            LineFormat::from_dispatch,
        )
    }
}
/// Typed Office fill object for solid, gradient, pattern, colour, and transparency settings.
pub struct FillFormat {
    inner: DispatchObject,
}
impl Debug for FillFormat {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("FillFormat").field(&self.inner).finish()
    }
}
impl FillFormat {
    pub(super) fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
        Self {
            inner: dispatch("FillFormat", value),
        }
    }
    /// Applies a solid Office drawing fill.
    pub fn solid(&self) -> Result<(), ExcelComError> {
        call(&self.inner, "excel.fillformat.solid", vec![])
    }
    /// Removes the Office drawing fill.
    pub fn no_fill(&self) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.fillformat.visible",
            OwnedVariant::bool(false),
        )
    }
    /// Returns whether Excel renders this Office drawing fill.
    pub fn visible(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.fillformat.visible")
    }
    /// Returns the Office foreground colour surface.
    pub fn fore_color(&self) -> Result<ColorFormat, ExcelComError> {
        get_dispatch(
            &self.inner,
            "excel.fillformat.forecolor",
            ColorFormat::from_dispatch,
        )
    }
    /// Returns the Office background colour surface.
    pub fn back_color(&self) -> Result<ColorFormat, ExcelComError> {
        get_dispatch(
            &self.inner,
            "excel.fillformat.backcolor",
            ColorFormat::from_dispatch,
        )
    }
    /// Returns Office fill transparency from 0 through 1.
    pub fn transparency(&self) -> Result<f64, ExcelComError> {
        get_f64(
            &self.inner,
            "excel.fillformat.transparency",
            "FillFormat.Transparency was not numeric",
        )
    }
    /// Sets Office fill transparency from 0 through 1.
    pub fn set_transparency(&self, value: f64) -> Result<(), ExcelComError> {
        if !value.is_finite() || !(0.0..=1.0).contains(&value) {
            return Err(ExcelComError::Unsupported {
                detail: "FillFormat.Transparency must be between 0 and 1",
            });
        }
        put(
            &self.inner,
            "excel.fillformat.transparency",
            OwnedVariant::f64(value),
        )
    }
    /// Applies an Office one-colour gradient using Excel/Office raw style and variant values.
    pub fn one_color_gradient(
        &self,
        style: i32,
        variant: i32,
        degree: f64,
    ) -> Result<(), ExcelComError> {
        finite(degree, "FillFormat gradient degree must be finite")?;
        call(
            &self.inner,
            "excel.fillformat.onecolorgradient",
            vec![
                OwnedVariant::i32(style),
                OwnedVariant::i32(variant),
                OwnedVariant::f64(degree),
            ],
        )
    }
    /// Applies an Office two-colour gradient using raw Office style and variant values.
    pub fn two_color_gradient(&self, style: i32, variant: i32) -> Result<(), ExcelComError> {
        call(
            &self.inner,
            "excel.fillformat.twocolorgradient",
            vec![OwnedVariant::i32(style), OwnedVariant::i32(variant)],
        )
    }
    /// Applies an Office preset gradient using raw Office style, variant, and preset values.
    pub fn preset_gradient(
        &self,
        style: i32,
        variant: i32,
        preset: i32,
    ) -> Result<(), ExcelComError> {
        call(
            &self.inner,
            "excel.fillformat.presetgradient",
            vec![
                OwnedVariant::i32(style),
                OwnedVariant::i32(variant),
                OwnedVariant::i32(preset),
            ],
        )
    }
    /// Applies an Office pattern fill using a raw Office pattern value.
    pub fn patterned(&self, pattern: i32) -> Result<(), ExcelComError> {
        call(
            &self.inner,
            "excel.fillformat.patterned",
            vec![OwnedVariant::i32(pattern)],
        )
    }
}
/// Typed Office line object for visibility, colour, transparency, weight, and dash settings.
pub struct LineFormat {
    inner: DispatchObject,
}
impl Debug for LineFormat {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("LineFormat").field(&self.inner).finish()
    }
}
impl LineFormat {
    pub(super) fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
        Self {
            inner: dispatch("LineFormat", value),
        }
    }
    /// Returns whether the Office drawing line is visible.
    pub fn visible(&self) -> Result<bool, ExcelComError> {
        get_bool(&self.inner, "excel.lineformat.visible")
    }
    /// Changes Office drawing-line visibility.
    pub fn set_visible(&self, value: bool) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.lineformat.visible",
            OwnedVariant::bool(value),
        )
    }
    /// Returns the Office line foreground colour surface.
    pub fn fore_color(&self) -> Result<ColorFormat, ExcelComError> {
        get_dispatch(
            &self.inner,
            "excel.lineformat.forecolor",
            ColorFormat::from_dispatch,
        )
    }
    /// Returns Office line transparency from 0 through 1.
    pub fn transparency(&self) -> Result<f64, ExcelComError> {
        get_f64(
            &self.inner,
            "excel.lineformat.transparency",
            "LineFormat.Transparency was not numeric",
        )
    }
    /// Sets Office line transparency from 0 through 1.
    pub fn set_transparency(&self, value: f64) -> Result<(), ExcelComError> {
        if !value.is_finite() || !(0.0..=1.0).contains(&value) {
            return Err(ExcelComError::Unsupported {
                detail: "LineFormat.Transparency must be between 0 and 1",
            });
        }
        put(
            &self.inner,
            "excel.lineformat.transparency",
            OwnedVariant::f64(value),
        )
    }
    /// Returns Office line weight in points.
    pub fn weight(&self) -> Result<f64, ExcelComError> {
        get_f64(
            &self.inner,
            "excel.lineformat.weight",
            "LineFormat.Weight was not numeric",
        )
    }
    /// Sets Office line weight in points.
    pub fn set_weight(&self, value: f64) -> Result<(), ExcelComError> {
        positive(value, "LineFormat.Weight must be positive")?;
        put(
            &self.inner,
            "excel.lineformat.weight",
            OwnedVariant::f64(value),
        )
    }
    /// Returns the raw Office dash-style value.
    pub fn dash_style(&self) -> Result<i32, ExcelComError> {
        get_i32(
            &self.inner,
            "excel.lineformat.dashstyle",
            "LineFormat.DashStyle was not an integer",
        )
    }
    /// Sets the raw Office dash-style value.
    pub fn set_dash_style(&self, value: i32) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.lineformat.dashstyle",
            OwnedVariant::i32(value),
        )
    }
}

/// Office colour surface used by chart fills and lines.
pub struct ColorFormat {
    inner: DispatchObject,
}
impl Debug for ColorFormat {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_tuple("ColorFormat")
            .field(&self.inner)
            .finish()
    }
}
impl ColorFormat {
    pub(super) fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
        Self {
            inner: dispatch("ColorFormat", value),
        }
    }
    /// Returns direct RGB when Excel currently exposes a direct colour value.
    pub fn rgb(&self) -> Result<ExcelColor, ExcelComError> {
        Ok(ExcelColor::from_raw(get_i32(
            &self.inner,
            "excel.colorformat.rgb",
            "ColorFormat.RGB was not an integer",
        )?))
    }
    /// Sets direct Excel RGB in Excel's red-low-byte colour order.
    pub fn set_rgb(&self, value: ExcelColor) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.colorformat.rgb",
            OwnedVariant::i32(value.raw()),
        )
    }
    /// Returns the raw Office theme-colour selector.
    pub fn theme_color(&self) -> Result<ThemeColor, ExcelComError> {
        Ok(ThemeColor::from_raw(get_i32(
            &self.inner,
            "excel.colorformat.objectthemecolor",
            "ColorFormat.ObjectThemeColor was not an integer",
        )?))
    }
    /// Sets the Office theme-colour selector.
    pub fn set_theme_color(&self, value: ThemeColor) -> Result<(), ExcelComError> {
        put(
            &self.inner,
            "excel.colorformat.objectthemecolor",
            OwnedVariant::i32(value.raw()),
        )
    }
    /// Returns Office theme tint-and-shade.
    pub fn tint_and_shade(&self) -> Result<f64, ExcelComError> {
        get_f64(
            &self.inner,
            "excel.colorformat.tintandshade",
            "ColorFormat.TintAndShade was not numeric",
        )
    }
    /// Sets Office theme tint-and-shade in the inclusive -1 through 1 range.
    pub fn set_tint_and_shade(&self, value: f64) -> Result<(), ExcelComError> {
        if !value.is_finite() || !(-1.0..=1.0).contains(&value) {
            return Err(ExcelComError::Unsupported {
                detail: "ColorFormat.TintAndShade must be between -1 and 1",
            });
        }
        put(
            &self.inner,
            "excel.colorformat.tintandshade",
            OwnedVariant::f64(value),
        )
    }
}

/// Formatting surface for the 3-D walls returned by an applicable Chart.
pub struct Walls {
    inner: DispatchObject,
}
impl Debug for Walls {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter.debug_tuple("Walls").field(&self.inner).finish()
    }
}
impl Walls {
    pub(super) fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
        Self {
            inner: dispatch("Walls", value),
        }
    }
    /// Returns Office drawing formatting for these walls.
    pub fn format(&self) -> Result<ChartFormat, ExcelComError> {
        get_dispatch(
            &self.inner,
            "excel.walls.format",
            ChartFormat::from_dispatch,
        )
    }
}

/// Formatting surface for the 3-D floor returned by an applicable Chart.
pub struct Floor {
    inner: DispatchObject,
}
impl Debug for Floor {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter.debug_tuple("Floor").field(&self.inner).finish()
    }
}
impl Floor {
    pub(super) fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
        Self {
            inner: dispatch("Floor", value),
        }
    }
    /// Returns Office drawing formatting for the floor.
    pub fn format(&self) -> Result<ChartFormat, ExcelComError> {
        get_dispatch(
            &self.inner,
            "excel.floor.format",
            ChartFormat::from_dispatch,
        )
    }
}

fn chart_element_bounds(
    target: &DispatchObject,
    prefix: &'static str,
) -> Result<ChartElementBounds, ExcelComError> {
    let left = if prefix == "excel.chartarea" {
        "excel.chartarea.left"
    } else {
        "excel.plotarea.left"
    };
    let top = if prefix == "excel.chartarea" {
        "excel.chartarea.top"
    } else {
        "excel.plotarea.top"
    };
    let width = if prefix == "excel.chartarea" {
        "excel.chartarea.width"
    } else {
        "excel.plotarea.width"
    };
    let height = if prefix == "excel.chartarea" {
        "excel.chartarea.height"
    } else {
        "excel.plotarea.height"
    };
    Ok(ChartElementBounds {
        left: get_f64(target, left, "chart element Left was not numeric")?,
        top: get_f64(target, top, "chart element Top was not numeric")?,
        width: get_f64(target, width, "chart element Width was not numeric")?,
        height: get_f64(target, height, "chart element Height was not numeric")?,
    })
}
