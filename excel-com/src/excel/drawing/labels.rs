//! Focused drawing implementation component.
#![allow(missing_docs)]
#[allow(unused_imports)]
use super::helpers::*;
#[allow(unused_imports)]
use super::types::*;
#[allow(unused_imports)]
use super::*;
pub struct ChartTitle {
    inner: DispatchObject,
}
impl Debug for ChartTitle {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("ChartTitle").field(&self.inner).finish()
    }
}
impl ChartTitle {
    pub(super) fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
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
    pub(super) fn from_dispatch(value: ComPtr<Dispatch>) -> Self {
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
