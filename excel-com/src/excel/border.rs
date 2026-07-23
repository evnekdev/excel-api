//! Excel Border wrapper.

use std::fmt::{Debug, Formatter};

use crate::ExcelComError;
use crate::automation::OwnedVariant;
use crate::excel::formatting::{
    BorderLineStyle, BorderWeight, ExcelColor, ExcelColorIndex, MixedValue, finite, map_mixed,
    mixed_f64, mixed_i32, property_mixed_get, property_put_value,
};
use crate::excel::{DispatchObject, ThemeColor};
use crate::internal::{ComPtr, Dispatch};

/// An apartment-bound Excel Border object returned by [`crate::Borders::item`].
///
/// Setting [`BorderLineStyle::NONE`] removes the selected border according to
/// Excel's own line-style policy. This wrapper is neither `Send` nor `Sync`.
pub struct Border {
    inner: DispatchObject,
}

impl Debug for Border {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter.debug_tuple("Border").field(&self.inner).finish()
    }
}

impl Clone for Border {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}

impl Border {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "Border",
            },
        }
    }

    /// Returns the border color or a mixed result.
    pub fn color(&self) -> Result<MixedValue<ExcelColor>, ExcelComError> {
        property_mixed_get(&self.inner, "excel.border.color", |value| {
            mixed_i32(value).map(|result| map_mixed(result, ExcelColor::from_raw))
        })
    }

    /// Sets the signed raw Excel border color.
    pub fn set_color(&self, color: ExcelColor) -> Result<(), ExcelComError> {
        property_put_value(
            &self.inner,
            "excel.border.color",
            OwnedVariant::i32(color.raw()),
        )
    }

    /// Returns the border color index or a mixed result.
    pub fn color_index(&self) -> Result<MixedValue<ExcelColorIndex>, ExcelComError> {
        property_mixed_get(&self.inner, "excel.border.colorindex", |value| {
            mixed_i32(value).map(|result| map_mixed(result, ExcelColorIndex::from_raw))
        })
    }

    /// Sets the forward-compatible Excel border color index.
    pub fn set_color_index(&self, index: ExcelColorIndex) -> Result<(), ExcelComError> {
        property_put_value(
            &self.inner,
            "excel.border.colorindex",
            OwnedVariant::i32(index.raw()),
        )
    }

    /// Returns the border line style or a mixed result.
    pub fn line_style(&self) -> Result<MixedValue<BorderLineStyle>, ExcelComError> {
        property_mixed_get(&self.inner, "excel.border.linestyle", |value| {
            mixed_i32(value).map(|result| map_mixed(result, BorderLineStyle::from_raw))
        })
    }

    /// Sets the Excel border line style.
    ///
    /// Use [`BorderLineStyle::NONE`] to ask Excel to remove the border.
    pub fn set_line_style(&self, style: BorderLineStyle) -> Result<(), ExcelComError> {
        property_put_value(
            &self.inner,
            "excel.border.linestyle",
            OwnedVariant::i32(style.raw()),
        )
    }

    /// Returns the border weight or a mixed result.
    pub fn weight(&self) -> Result<MixedValue<BorderWeight>, ExcelComError> {
        property_mixed_get(&self.inner, "excel.border.weight", |value| {
            mixed_i32(value).map(|result| map_mixed(result, BorderWeight::from_raw))
        })
    }

    /// Sets the Excel border weight.
    pub fn set_weight(&self, weight: BorderWeight) -> Result<(), ExcelComError> {
        property_put_value(
            &self.inner,
            "excel.border.weight",
            OwnedVariant::i32(weight.raw()),
        )
    }

    /// Returns the border theme color or a mixed result.
    pub fn theme_color(&self) -> Result<MixedValue<ThemeColor>, ExcelComError> {
        property_mixed_get(&self.inner, "excel.border.themecolor", |value| {
            mixed_i32(value).map(|result| map_mixed(result, ThemeColor::from_raw))
        })
    }
    /// Sets the border theme color.
    pub fn set_theme_color(&self, value: ThemeColor) -> Result<(), ExcelComError> {
        property_put_value(
            &self.inner,
            "excel.border.themecolor",
            OwnedVariant::i32(value.raw()),
        )
    }
    /// Returns the border theme tint or a mixed result.
    pub fn tint_and_shade(&self) -> Result<MixedValue<f64>, ExcelComError> {
        property_mixed_get(&self.inner, "excel.border.tintandshade", mixed_f64)
    }
    /// Sets the border theme tint.
    pub fn set_tint_and_shade(&self, value: f64) -> Result<(), ExcelComError> {
        finite(value)?;
        if !(-1.0..=1.0).contains(&value) {
            return Err(ExcelComError::Unsupported {
                detail: "TintAndShade must be between -1.0 and 1.0",
            });
        }
        property_put_value(
            &self.inner,
            "excel.border.tintandshade",
            OwnedVariant::f64(value),
        )
    }
}
