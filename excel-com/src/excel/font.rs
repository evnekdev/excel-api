//! Excel Font wrapper.

use std::fmt::{Debug, Formatter};

use crate::ExcelComError;
use crate::automation::OwnedVariant;
use crate::excel::formatting::{
    ExcelColor, ExcelColorIndex, MixedValue, UnderlineStyle, finite, map_mixed, mixed_bool,
    mixed_f64, mixed_i32, mixed_string, property_mixed_get, property_put_value,
};
use crate::excel::{DispatchObject, ThemeColor, ThemeFont, text::text_bstr};
use crate::internal::{ComPtr, Dispatch};

/// An apartment-bound Excel Font object returned by [`crate::Range::font`].
///
/// Font getters return [`MixedValue`] when the source Range spans cells with
/// different formatting. This wrapper is neither `Send` nor `Sync`.
pub struct Font {
    inner: DispatchObject,
}

impl Debug for Font {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter.debug_tuple("Font").field(&self.inner).finish()
    }
}

impl Clone for Font {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}

impl Font {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "Font",
            },
        }
    }

    /// Returns the font name or a mixed result.
    pub fn name(&self) -> Result<MixedValue<String>, ExcelComError> {
        property_mixed_get(&self.inner, "excel.font.name", mixed_string)
    }

    /// Sets the font name after rejecting embedded NUL text before COM.
    pub fn set_name(&self, name: &str) -> Result<(), ExcelComError> {
        property_put_value(&self.inner, "excel.font.name", text_bstr(name)?)
    }

    /// Returns the font size in points or a mixed result.
    pub fn size(&self) -> Result<MixedValue<f64>, ExcelComError> {
        property_mixed_get(&self.inner, "excel.font.size", mixed_f64)
    }

    /// Sets the font size in points.
    ///
    /// Non-finite sizes are rejected before COM; Excel remains responsible for
    /// supported font families and size bounds.
    pub fn set_size(&self, points: f64) -> Result<(), ExcelComError> {
        finite(points)?;
        property_put_value(&self.inner, "excel.font.size", OwnedVariant::f64(points))
    }

    /// Returns whether the font is bold or a mixed result.
    pub fn bold(&self) -> Result<MixedValue<bool>, ExcelComError> {
        property_mixed_get(&self.inner, "excel.font.bold", mixed_bool)
    }

    /// Sets whether the font is bold.
    pub fn set_bold(&self, value: bool) -> Result<(), ExcelComError> {
        property_put_value(&self.inner, "excel.font.bold", OwnedVariant::bool(value))
    }

    /// Returns whether the font is italic or a mixed result.
    pub fn italic(&self) -> Result<MixedValue<bool>, ExcelComError> {
        property_mixed_get(&self.inner, "excel.font.italic", mixed_bool)
    }

    /// Sets whether the font is italic.
    pub fn set_italic(&self, value: bool) -> Result<(), ExcelComError> {
        property_put_value(&self.inner, "excel.font.italic", OwnedVariant::bool(value))
    }

    /// Returns the underline style or a mixed result.
    pub fn underline(&self) -> Result<MixedValue<UnderlineStyle>, ExcelComError> {
        property_mixed_get(&self.inner, "excel.font.underline", |value| {
            mixed_i32(value).map(|result| map_mixed(result, UnderlineStyle::from_raw))
        })
    }

    /// Sets the forward-compatible Excel underline style.
    pub fn set_underline(&self, value: UnderlineStyle) -> Result<(), ExcelComError> {
        property_put_value(
            &self.inner,
            "excel.font.underline",
            OwnedVariant::i32(value.raw()),
        )
    }

    /// Returns whether the font is struck through or a mixed result.
    pub fn strikethrough(&self) -> Result<MixedValue<bool>, ExcelComError> {
        property_mixed_get(&self.inner, "excel.font.strikethrough", mixed_bool)
    }

    /// Sets whether the font is struck through.
    pub fn set_strikethrough(&self, value: bool) -> Result<(), ExcelComError> {
        property_put_value(
            &self.inner,
            "excel.font.strikethrough",
            OwnedVariant::bool(value),
        )
    }

    /// Returns the font color or a mixed result.
    pub fn color(&self) -> Result<MixedValue<ExcelColor>, ExcelComError> {
        property_mixed_get(&self.inner, "excel.font.color", |value| {
            mixed_i32(value).map(|result| map_mixed(result, ExcelColor::from_raw))
        })
    }

    /// Sets the signed raw Excel font color.
    pub fn set_color(&self, color: ExcelColor) -> Result<(), ExcelComError> {
        property_put_value(
            &self.inner,
            "excel.font.color",
            OwnedVariant::i32(color.raw()),
        )
    }

    /// Returns the font color index or a mixed result.
    pub fn color_index(&self) -> Result<MixedValue<ExcelColorIndex>, ExcelComError> {
        property_mixed_get(&self.inner, "excel.font.colorindex", |value| {
            mixed_i32(value).map(|result| map_mixed(result, ExcelColorIndex::from_raw))
        })
    }

    /// Sets the forward-compatible Excel font color index.
    pub fn set_color_index(&self, index: ExcelColorIndex) -> Result<(), ExcelComError> {
        property_put_value(
            &self.inner,
            "excel.font.colorindex",
            OwnedVariant::i32(index.raw()),
        )
    }

    /// Returns the subscript setting or a mixed result.
    pub fn subscript(&self) -> Result<MixedValue<bool>, ExcelComError> {
        property_mixed_get(&self.inner, "excel.font.subscript", mixed_bool)
    }
    /// Sets the subscript setting.
    pub fn set_subscript(&self, value: bool) -> Result<(), ExcelComError> {
        property_put_value(
            &self.inner,
            "excel.font.subscript",
            OwnedVariant::bool(value),
        )
    }
    /// Returns the superscript setting or a mixed result.
    pub fn superscript(&self) -> Result<MixedValue<bool>, ExcelComError> {
        property_mixed_get(&self.inner, "excel.font.superscript", mixed_bool)
    }
    /// Sets the superscript setting.
    pub fn set_superscript(&self, value: bool) -> Result<(), ExcelComError> {
        property_put_value(
            &self.inner,
            "excel.font.superscript",
            OwnedVariant::bool(value),
        )
    }
    /// Returns the shadow setting or a mixed result.
    pub fn shadow(&self) -> Result<MixedValue<bool>, ExcelComError> {
        property_mixed_get(&self.inner, "excel.font.shadow", mixed_bool)
    }
    /// Sets the shadow setting.
    pub fn set_shadow(&self, value: bool) -> Result<(), ExcelComError> {
        property_put_value(&self.inner, "excel.font.shadow", OwnedVariant::bool(value))
    }
    /// Returns the outline setting or a mixed result.
    pub fn outline_font(&self) -> Result<MixedValue<bool>, ExcelComError> {
        property_mixed_get(&self.inner, "excel.font.outlinefont", mixed_bool)
    }
    /// Sets the outline setting.
    pub fn set_outline_font(&self, value: bool) -> Result<(), ExcelComError> {
        property_put_value(
            &self.inner,
            "excel.font.outlinefont",
            OwnedVariant::bool(value),
        )
    }
    /// Returns Excel's font-style text or a mixed result.
    pub fn font_style(&self) -> Result<MixedValue<String>, ExcelComError> {
        property_mixed_get(&self.inner, "excel.font.fontstyle", mixed_string)
    }
    /// Sets Excel's font-style text.
    pub fn set_font_style(&self, value: &str) -> Result<(), ExcelComError> {
        property_put_value(&self.inner, "excel.font.fontstyle", text_bstr(value)?)
    }
    /// Returns the theme-font choice or a mixed result.
    pub fn theme_font(&self) -> Result<MixedValue<ThemeFont>, ExcelComError> {
        property_mixed_get(&self.inner, "excel.font.themefont", |value| {
            mixed_i32(value).map(|result| map_mixed(result, ThemeFont::from_raw))
        })
    }
    /// Sets the theme-font choice.
    pub fn set_theme_font(&self, value: ThemeFont) -> Result<(), ExcelComError> {
        property_put_value(
            &self.inner,
            "excel.font.themefont",
            OwnedVariant::i32(value.raw()),
        )
    }
    /// Returns the theme color or a mixed result.
    pub fn theme_color(&self) -> Result<MixedValue<ThemeColor>, ExcelComError> {
        property_mixed_get(&self.inner, "excel.font.themecolor", |value| {
            mixed_i32(value).map(|result| map_mixed(result, ThemeColor::from_raw))
        })
    }
    /// Sets the theme color.
    pub fn set_theme_color(&self, value: ThemeColor) -> Result<(), ExcelComError> {
        property_put_value(
            &self.inner,
            "excel.font.themecolor",
            OwnedVariant::i32(value.raw()),
        )
    }
    /// Returns theme tint or a mixed result.
    pub fn tint_and_shade(&self) -> Result<MixedValue<f64>, ExcelComError> {
        property_mixed_get(&self.inner, "excel.font.tintandshade", mixed_f64)
    }
    /// Sets the theme tint in Excel's inclusive `-1.0..=1.0` range.
    pub fn set_tint_and_shade(&self, value: f64) -> Result<(), ExcelComError> {
        finite(value)?;
        if !(-1.0..=1.0).contains(&value) {
            return Err(ExcelComError::Unsupported {
                detail: "TintAndShade must be between -1.0 and 1.0",
            });
        }
        property_put_value(
            &self.inner,
            "excel.font.tintandshade",
            OwnedVariant::f64(value),
        )
    }
}

#[cfg(test)]
mod tests {
    use crate::excel::text::text_bstr;

    #[test]
    fn font_text_rejects_embedded_nul_before_com() {
        assert!(text_bstr("A\0rial").is_err());
    }
}
