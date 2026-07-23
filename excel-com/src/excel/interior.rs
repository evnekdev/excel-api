//! Excel Interior wrapper.

use std::fmt::{Debug, Formatter};

use crate::ExcelComError;
use crate::automation::OwnedVariant;
use crate::excel::formatting::{
    ExcelColor, ExcelColorIndex, FillPattern, MixedValue, finite, map_mixed, mixed_f64, mixed_i32,
    property_mixed_get, property_put_value,
};
use crate::excel::{DispatchObject, ThemeColor};
use crate::internal::{ComPtr, Dispatch};

/// An apartment-bound Excel fill object returned by [`crate::Range::interior`].
///
/// Its getters preserve mixed and empty Excel results through [`MixedValue`].
/// This wrapper is neither `Send` nor `Sync`.
pub struct Interior {
    inner: DispatchObject,
}

impl Debug for Interior {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_tuple("Interior")
            .field(&self.inner)
            .finish()
    }
}

impl Clone for Interior {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}

impl Interior {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "Interior",
            },
        }
    }

    /// Returns the fill color or a mixed result.
    pub fn color(&self) -> Result<MixedValue<ExcelColor>, ExcelComError> {
        property_mixed_get(&self.inner, "excel.interior.color", |value| {
            mixed_i32(value).map(|result| map_mixed(result, ExcelColor::from_raw))
        })
    }

    /// Sets the signed raw Excel fill color.
    pub fn set_color(&self, color: ExcelColor) -> Result<(), ExcelComError> {
        property_put_value(
            &self.inner,
            "excel.interior.color",
            OwnedVariant::i32(color.raw()),
        )
    }

    /// Returns the fill color index or a mixed result.
    pub fn color_index(&self) -> Result<MixedValue<ExcelColorIndex>, ExcelComError> {
        property_mixed_get(&self.inner, "excel.interior.colorindex", |value| {
            mixed_i32(value).map(|result| map_mixed(result, ExcelColorIndex::from_raw))
        })
    }

    /// Sets the forward-compatible Excel fill color index.
    pub fn set_color_index(&self, index: ExcelColorIndex) -> Result<(), ExcelComError> {
        property_put_value(
            &self.inner,
            "excel.interior.colorindex",
            OwnedVariant::i32(index.raw()),
        )
    }

    /// Returns the fill pattern or a mixed result.
    pub fn pattern(&self) -> Result<MixedValue<FillPattern>, ExcelComError> {
        property_mixed_get(&self.inner, "excel.interior.pattern", |value| {
            mixed_i32(value).map(|result| map_mixed(result, FillPattern::from_raw))
        })
    }

    /// Sets the forward-compatible Excel fill pattern.
    pub fn set_pattern(&self, pattern: FillPattern) -> Result<(), ExcelComError> {
        property_put_value(
            &self.inner,
            "excel.interior.pattern",
            OwnedVariant::i32(pattern.raw()),
        )
    }

    /// Returns the fill pattern color or a mixed result.
    pub fn pattern_color(&self) -> Result<MixedValue<ExcelColor>, ExcelComError> {
        property_mixed_get(&self.inner, "excel.interior.patterncolor", |value| {
            mixed_i32(value).map(|result| map_mixed(result, ExcelColor::from_raw))
        })
    }

    /// Sets the signed raw Excel pattern color.
    pub fn set_pattern_color(&self, color: ExcelColor) -> Result<(), ExcelComError> {
        property_put_value(
            &self.inner,
            "excel.interior.patterncolor",
            OwnedVariant::i32(color.raw()),
        )
    }

    /// Returns the fill pattern color index or a mixed result.
    pub fn pattern_color_index(&self) -> Result<MixedValue<ExcelColorIndex>, ExcelComError> {
        property_mixed_get(&self.inner, "excel.interior.patterncolorindex", |value| {
            mixed_i32(value).map(|result| map_mixed(result, ExcelColorIndex::from_raw))
        })
    }

    /// Sets the forward-compatible Excel pattern color index.
    pub fn set_pattern_color_index(&self, index: ExcelColorIndex) -> Result<(), ExcelComError> {
        property_put_value(
            &self.inner,
            "excel.interior.patterncolorindex",
            OwnedVariant::i32(index.raw()),
        )
    }

    /// Returns the fill theme color or a mixed result.
    pub fn theme_color(&self) -> Result<MixedValue<ThemeColor>, ExcelComError> {
        property_mixed_get(&self.inner, "excel.interior.themecolor", |value| {
            mixed_i32(value).map(|result| map_mixed(result, ThemeColor::from_raw))
        })
    }
    /// Sets the fill theme color.
    pub fn set_theme_color(&self, value: ThemeColor) -> Result<(), ExcelComError> {
        property_put_value(
            &self.inner,
            "excel.interior.themecolor",
            OwnedVariant::i32(value.raw()),
        )
    }
    /// Returns fill theme tint or a mixed result.
    pub fn tint_and_shade(&self) -> Result<MixedValue<f64>, ExcelComError> {
        property_mixed_get(&self.inner, "excel.interior.tintandshade", mixed_f64)
    }
    /// Sets fill theme tint.
    pub fn set_tint_and_shade(&self, value: f64) -> Result<(), ExcelComError> {
        tint_put(&self.inner, "excel.interior.tintandshade", value)
    }
    /// Returns pattern theme tint or a mixed result.
    pub fn pattern_tint_and_shade(&self) -> Result<MixedValue<f64>, ExcelComError> {
        property_mixed_get(&self.inner, "excel.interior.patterntintandshade", mixed_f64)
    }
    /// Sets pattern theme tint.
    pub fn set_pattern_tint_and_shade(&self, value: f64) -> Result<(), ExcelComError> {
        tint_put(&self.inner, "excel.interior.patterntintandshade", value)
    }
}

fn tint_put(target: &DispatchObject, id: &'static str, value: f64) -> Result<(), ExcelComError> {
    finite(value)?;
    if !(-1.0..=1.0).contains(&value) {
        return Err(ExcelComError::Unsupported {
            detail: "TintAndShade must be between -1.0 and 1.0",
        });
    }
    property_put_value(target, id, OwnedVariant::f64(value))
}
