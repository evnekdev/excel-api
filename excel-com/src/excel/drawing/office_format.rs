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
    pub(super) fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
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
    pub(super) fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
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
    pub(super) fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
        Self {
            inner: dispatch("LineFormat", value),
        }
    }
}
