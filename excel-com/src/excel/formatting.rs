//! Shared formatting representations and Range formatting operations.

use crate::automation::{OwnedVariant, invoke, property_get, property_put};
use crate::excel::{Borders, DispatchObject, Font, Interior, Range, text::text_bstr};
use crate::object_model::{MemberId, member};
use crate::{ConversionError, ExcelComError};
use windows_sys::Win32::System::Variant::{VT_EMPTY, VT_NULL};

/// The result of a formatting getter that can span differently formatted cells.
///
/// Excel commonly returns `VT_NULL` when the selected cells disagree. The
/// wrapper exposes that outcome as [`MixedValue::Mixed`] instead of coercing
/// it to an ordinary scalar. `VT_EMPTY` is retained as [`MixedValue::Empty`].
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum MixedValue<T> {
    /// Every selected cell has one concrete value.
    Uniform(T),
    /// Excel reported that the selected cells do not have one common value.
    Mixed,
    /// Excel returned an empty/unset formatting value.
    Empty,
}

/// An Excel Automation color integer.
///
/// Excel stores RGB channels in OLE `COLORREF` order: red occupies the low
/// byte, green the next byte, and blue the third byte. The signed raw form
/// preserves sentinel values that Excel can return without treating them as
/// web-style `0xRRGGBB` colors.
#[derive(Clone, Copy, Debug, Eq, PartialEq, Hash)]
#[repr(transparent)]
pub struct ExcelColor(i32);

impl ExcelColor {
    /// Creates a color from Excel's signed raw Automation representation.
    pub const fn from_raw(value: i32) -> Self {
        Self(value)
    }

    /// Returns Excel's signed raw Automation representation.
    pub const fn raw(self) -> i32 {
        self.0
    }

    /// Creates a color from red, green, and blue channels in Excel's RGB order.
    pub const fn from_rgb(red: u8, green: u8, blue: u8) -> Self {
        Self(red as i32 | ((green as i32) << 8) | ((blue as i32) << 16))
    }

    /// Returns the red channel from the low byte of the raw representation.
    pub const fn red(self) -> u8 {
        self.0 as u8
    }

    /// Returns the green channel from the second byte of the raw representation.
    pub const fn green(self) -> u8 {
        (self.0 >> 8) as u8
    }

    /// Returns the blue channel from the third byte of the raw representation.
    pub const fn blue(self) -> u8 {
        (self.0 >> 16) as u8
    }
}

/// A forward-compatible Excel indexed-color value.
///
/// Palette indices depend on the workbook palette and Excel version. Unknown
/// values are retained through [`Self::from_raw`].
#[derive(Clone, Copy, Debug, Eq, PartialEq, Hash)]
#[repr(transparent)]
pub struct ExcelColorIndex(i32);

impl ExcelColorIndex {
    /// Excel chooses the automatic color.
    pub const AUTOMATIC: Self = Self(-4105);
    /// Excel uses no indexed color.
    pub const NONE: Self = Self(-4142);

    /// Creates an indexed color from its raw Excel value.
    pub const fn from_raw(value: i32) -> Self {
        Self(value)
    }

    /// Returns the raw Excel indexed-color value.
    pub const fn raw(self) -> i32 {
        self.0
    }
}

macro_rules! raw_formatting_type {
    ($(#[$meta:meta])* $name:ident { $($(#[$constant_meta:meta])* $constant:ident = $value:expr;)* }) => {
        $(#[$meta])*
        #[derive(Clone, Copy, Debug, Eq, PartialEq, Hash)]
        #[repr(transparent)]
        pub struct $name(i32);

        impl $name {
            $($(#[$constant_meta])* pub const $constant: Self = Self($value);)*

            /// Creates this formatting value from its raw Excel value.
            pub const fn from_raw(value: i32) -> Self {
                Self(value)
            }

            /// Returns the raw Excel value, preserving unknown constants.
            pub const fn raw(self) -> i32 {
                self.0
            }
        }
    };
}

raw_formatting_type! {
    /// A forward-compatible `XlHAlign` value.
    HorizontalAlignment {
        /// `xlHAlignGeneral`.
        GENERAL = 1;
        /// `xlHAlignLeft`.
        LEFT = -4131;
        /// `xlHAlignCenter`.
        CENTER = -4108;
        /// `xlHAlignRight`.
        RIGHT = -4152;
        /// `xlHAlignFill`.
        FILL = 5;
        /// `xlHAlignJustify`.
        JUSTIFY = -4130;
        /// `xlHAlignCenterAcrossSelection`.
        CENTER_ACROSS_SELECTION = 7;
        /// `xlHAlignDistributed`.
        DISTRIBUTED = -4117;
    }
}

raw_formatting_type! {
    /// A forward-compatible `XlVAlign` value.
    VerticalAlignment {
        /// `xlVAlignTop`.
        TOP = -4160;
        /// `xlVAlignCenter`.
        CENTER = -4108;
        /// `xlVAlignBottom`.
        BOTTOM = -4107;
        /// `xlVAlignJustify`.
        JUSTIFY = -4130;
        /// `xlVAlignDistributed`.
        DISTRIBUTED = -4117;
    }
}

raw_formatting_type! {
    /// A forward-compatible `XlUnderlineStyle` value.
    UnderlineStyle {
        /// `xlUnderlineStyleNone`.
        NONE = -4142;
        /// `xlUnderlineStyleSingle`.
        SINGLE = 2;
        /// `xlUnderlineStyleDouble`.
        DOUBLE = -4119;
        /// `xlUnderlineStyleSingleAccounting`.
        SINGLE_ACCOUNTING = 4;
        /// `xlUnderlineStyleDoubleAccounting`.
        DOUBLE_ACCOUNTING = 5;
    }
}

raw_formatting_type! {
    /// A forward-compatible `XlPattern` fill value.
    FillPattern {
        /// `xlPatternNone`.
        NONE = -4142;
        /// `xlPatternSolid`.
        SOLID = 1;
        /// `xlPatternAutomatic`.
        AUTOMATIC = -4105;
        /// `xlPatternGray25`.
        GRAY25 = -4124;
        /// `xlPatternGray50`.
        GRAY50 = -4125;
        /// `xlPatternGray75`.
        GRAY75 = -4126;
        /// `xlPatternHorizontal`.
        HORIZONTAL = -4128;
        /// `xlPatternVertical`.
        VERTICAL = -4166;
    }
}

raw_formatting_type! {
    /// A forward-compatible `XlLineStyle` border value.
    BorderLineStyle {
        /// `xlLineStyleNone`, which removes an Excel border.
        NONE = -4142;
        /// `xlContinuous`.
        CONTINUOUS = 1;
        /// `xlDash`.
        DASH = -4115;
        /// `xlDashDot`.
        DASH_DOT = 4;
        /// `xlDashDotDot`.
        DASH_DOT_DOT = 5;
        /// `xlDot`.
        DOT = -4118;
        /// `xlDouble`.
        DOUBLE = -4119;
        /// `xlSlantDashDot`.
        SLANT_DASH_DOT = 13;
    }
}

raw_formatting_type! {
    /// A forward-compatible `XlBorderWeight` value.
    BorderWeight {
        /// `xlHairline`.
        HAIRLINE = 1;
        /// `xlThin`.
        THIN = 2;
        /// `xlMedium`.
        MEDIUM = -4138;
        /// `xlThick`.
        THICK = 4;
    }
}

raw_formatting_type! {
    /// An enum key accepted by Excel's `Borders.Item` property.
    BorderIndex {
        /// `xlDiagonalDown`.
        DIAGONAL_DOWN = 5;
        /// `xlDiagonalUp`.
        DIAGONAL_UP = 6;
        /// `xlEdgeLeft`.
        EDGE_LEFT = 7;
        /// `xlEdgeTop`.
        EDGE_TOP = 8;
        /// `xlEdgeBottom`.
        EDGE_BOTTOM = 9;
        /// `xlEdgeRight`.
        EDGE_RIGHT = 10;
        /// `xlInsideVertical`.
        INSIDE_VERTICAL = 11;
        /// `xlInsideHorizontal`.
        INSIDE_HORIZONTAL = 12;
    }
}

pub(crate) fn mixed<T>(
    value: &OwnedVariant,
    decode: impl FnOnce(&OwnedVariant) -> Result<T, ExcelComError>,
) -> Result<MixedValue<T>, ExcelComError> {
    match value.vt() {
        VT_NULL => Ok(MixedValue::Mixed),
        VT_EMPTY => Ok(MixedValue::Empty),
        _ => decode(value).map(MixedValue::Uniform),
    }
}

pub(crate) fn mixed_bool(value: &OwnedVariant) -> Result<MixedValue<bool>, ExcelComError> {
    mixed(value, |value| {
        value.as_bool().ok_or(ExcelComError::Conversion(
            ConversionError::UnsupportedVariantType {
                vartype: value.vt(),
            },
        ))
    })
}

pub(crate) fn mixed_f64(value: &OwnedVariant) -> Result<MixedValue<f64>, ExcelComError> {
    mixed(value, |value| {
        value
            .as_f64()
            .or_else(|| value.as_i32().map(f64::from))
            .filter(|value| value.is_finite())
            .ok_or(ExcelComError::Conversion(
                ConversionError::UnsupportedVariantType {
                    vartype: value.vt(),
                },
            ))
    })
}

pub(crate) fn mixed_i32(value: &OwnedVariant) -> Result<MixedValue<i32>, ExcelComError> {
    mixed(value, |value| {
        value
            .as_i32()
            .or_else(|| {
                value.as_f64().and_then(|number| {
                    (number.is_finite()
                        && number.fract() == 0.0
                        && number >= f64::from(i32::MIN)
                        && number <= f64::from(i32::MAX))
                    .then_some(number as i32)
                })
            })
            .ok_or(ExcelComError::Conversion(
                ConversionError::UnsupportedVariantType {
                    vartype: value.vt(),
                },
            ))
    })
}

pub(crate) fn mixed_string(value: &OwnedVariant) -> Result<MixedValue<String>, ExcelComError> {
    mixed(value, OwnedVariant::as_string)
}

pub(crate) fn finite_nonnegative(value: f64) -> Result<(), ExcelComError> {
    if value.is_finite() && value >= 0.0 {
        Ok(())
    } else {
        Err(ExcelComError::Conversion(ConversionError::NonFiniteNumber))
    }
}

pub(crate) fn finite(value: f64) -> Result<(), ExcelComError> {
    value
        .is_finite()
        .then_some(())
        .ok_or(ExcelComError::Conversion(ConversionError::NonFiniteNumber))
}

impl Range {
    /// Returns the apartment-bound Font wrapper for this Range.
    pub fn font(&self) -> Result<Font, ExcelComError> {
        object_property(
            self.dispatch_object(),
            "excel.range.font",
            Font::from_dispatch,
        )
    }

    /// Returns the apartment-bound Interior wrapper for this Range.
    pub fn interior(&self) -> Result<Interior, ExcelComError> {
        object_property(
            self.dispatch_object(),
            "excel.range.interior",
            Interior::from_dispatch,
        )
    }

    /// Returns the apartment-bound Borders collection wrapper for this Range.
    pub fn borders(&self) -> Result<Borders, ExcelComError> {
        object_property(
            self.dispatch_object(),
            "excel.range.borders",
            Borders::from_dispatch,
        )
    }

    /// Returns the invariant Excel `NumberFormat` string or a mixed result.
    ///
    /// Excel's format-code syntax remains Excel-defined and can contain
    /// locale-sensitive constructs. This method deliberately does not use
    /// `NumberFormatLocal`.
    pub fn number_format(&self) -> Result<MixedValue<String>, ExcelComError> {
        range_mixed_get(self, "excel.range.numberformat", mixed_string)
    }

    /// Sets the invariant Excel `NumberFormat` string.
    pub fn set_number_format(&self, format_code: &str) -> Result<(), ExcelComError> {
        range_put(self, "excel.range.numberformat", text_bstr(format_code)?)
    }

    /// Returns the horizontal alignment or a mixed result.
    pub fn horizontal_alignment(&self) -> Result<MixedValue<HorizontalAlignment>, ExcelComError> {
        range_mixed_get(self, "excel.range.horizontalalignment", |value| {
            mixed_i32(value).map(|result| map_mixed(result, HorizontalAlignment::from_raw))
        })
    }

    /// Sets a forward-compatible horizontal alignment value.
    pub fn set_horizontal_alignment(
        &self,
        alignment: HorizontalAlignment,
    ) -> Result<(), ExcelComError> {
        range_put(
            self,
            "excel.range.horizontalalignment",
            OwnedVariant::i32(alignment.raw()),
        )
    }

    /// Returns the vertical alignment or a mixed result.
    pub fn vertical_alignment(&self) -> Result<MixedValue<VerticalAlignment>, ExcelComError> {
        range_mixed_get(self, "excel.range.verticalalignment", |value| {
            mixed_i32(value).map(|result| map_mixed(result, VerticalAlignment::from_raw))
        })
    }

    /// Sets a forward-compatible vertical alignment value.
    pub fn set_vertical_alignment(
        &self,
        alignment: VerticalAlignment,
    ) -> Result<(), ExcelComError> {
        range_put(
            self,
            "excel.range.verticalalignment",
            OwnedVariant::i32(alignment.raw()),
        )
    }

    /// Returns whether text wraps, or that the selected cells are mixed.
    pub fn wrap_text(&self) -> Result<MixedValue<bool>, ExcelComError> {
        range_mixed_get(self, "excel.range.wraptext", mixed_bool)
    }

    /// Enables or disables Excel text wrapping for the Range.
    pub fn set_wrap_text(&self, enabled: bool) -> Result<(), ExcelComError> {
        range_put(self, "excel.range.wraptext", OwnedVariant::bool(enabled))
    }

    /// Returns the row height in points, or a mixed result.
    pub fn row_height(&self) -> Result<MixedValue<f64>, ExcelComError> {
        range_mixed_get(self, "excel.range.rowheight", mixed_f64)
    }

    /// Sets the Range row height in points.
    ///
    /// `NaN`, infinities, and negative values are rejected before COM. Excel
    /// defines the meaning of zero and values above its supported maximum.
    pub fn set_row_height(&self, points: f64) -> Result<(), ExcelComError> {
        finite_nonnegative(points)?;
        range_put(self, "excel.range.rowheight", OwnedVariant::f64(points))
    }

    /// Returns the column width or a mixed result.
    ///
    /// A column-width unit is based on the Normal style font, not pixels.
    pub fn column_width(&self) -> Result<MixedValue<f64>, ExcelComError> {
        range_mixed_get(self, "excel.range.columnwidth", mixed_f64)
    }

    /// Sets the Excel column width.
    ///
    /// `NaN`, infinities, and negative values are rejected before COM. Zero
    /// is preserved for Excel to interpret, including its hidden-column policy.
    pub fn set_column_width(&self, width: f64) -> Result<(), ExcelComError> {
        finite_nonnegative(width)?;
        range_put(self, "excel.range.columnwidth", OwnedVariant::f64(width))
    }

    /// Invokes Excel `Range.AutoFit` without changing this Range's shape.
    ///
    /// Excel accepts rows or columns (normally [`Range::entire_row`] or
    /// [`Range::entire_column`]) and reports an invocation error for unsupported
    /// rectangular selections.
    pub fn auto_fit(&self) -> Result<(), ExcelComError> {
        let _ = invoke(
            &self.dispatch_object().dispatch,
            member(MemberId::new("excel.range.autofit"), false),
            vec![],
            false,
        )?;
        Ok(())
    }
}

pub(crate) fn object_property<T>(
    target: &DispatchObject,
    id: &'static str,
    construct: impl FnOnce(crate::internal::ComPtr<crate::internal::Dispatch>) -> T,
) -> Result<T, ExcelComError> {
    let mut value = property_get(&target.dispatch, member(MemberId::new(id), false), vec![])?;
    Ok(construct(value.take_dispatch()?))
}

pub(crate) fn property_mixed_get<T>(
    target: &DispatchObject,
    id: &'static str,
    decode: impl FnOnce(&OwnedVariant) -> Result<MixedValue<T>, ExcelComError>,
) -> Result<MixedValue<T>, ExcelComError> {
    let value = property_get(&target.dispatch, member(MemberId::new(id), false), vec![])?;
    decode(&value)
}

pub(crate) fn property_put_value(
    target: &DispatchObject,
    id: &'static str,
    value: OwnedVariant,
) -> Result<(), ExcelComError> {
    let _ = property_put(&target.dispatch, member(MemberId::new(id), true), value)?;
    Ok(())
}

fn range_mixed_get<T>(
    range: &Range,
    id: &'static str,
    decode: impl FnOnce(&OwnedVariant) -> Result<MixedValue<T>, ExcelComError>,
) -> Result<MixedValue<T>, ExcelComError> {
    property_mixed_get(range.dispatch_object(), id, decode)
}

fn range_put(range: &Range, id: &'static str, value: OwnedVariant) -> Result<(), ExcelComError> {
    property_put_value(range.dispatch_object(), id, value)
}

pub(crate) fn map_mixed<T, U>(value: MixedValue<T>, map: impl FnOnce(T) -> U) -> MixedValue<U> {
    match value {
        MixedValue::Uniform(value) => MixedValue::Uniform(map(value)),
        MixedValue::Mixed => MixedValue::Mixed,
        MixedValue::Empty => MixedValue::Empty,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::automation::{MemberDescriptor, MemberKind};
    use crate::{ComApartment, OwnedApplication};

    #[test]
    fn color_uses_excel_rgb_byte_order() {
        let color = ExcelColor::from_rgb(12, 34, 56);
        assert_eq!(color.raw(), 0x0038_220c);
        assert_eq!((color.red(), color.green(), color.blue()), (12, 34, 56));
        assert_eq!(ExcelColor::from_rgb(255, 0, 0).raw(), 255);
        assert_eq!(ExcelColor::from_rgb(0, 255, 0).raw(), 65_280);
        assert_eq!(ExcelColor::from_rgb(0, 0, 255).raw(), 16_711_680);
        assert_eq!(ExcelColor::from_raw(-1).raw(), -1);
    }

    #[test]
    fn mixed_decoding_retains_null_and_empty() {
        assert_eq!(
            mixed_bool(&OwnedVariant::bool(true)),
            Ok(MixedValue::Uniform(true))
        );
        assert_eq!(
            mixed_f64(&OwnedVariant::f64(12.0)),
            Ok(MixedValue::Uniform(12.0))
        );
        assert_eq!(mixed_i32(&OwnedVariant::i32(7)), Ok(MixedValue::Uniform(7)));
        assert_eq!(
            mixed_i32(&OwnedVariant::f64(7.0)),
            Ok(MixedValue::Uniform(7))
        );
        assert_eq!(
            mixed_string(&OwnedVariant::bstr("font").expect("BSTR")),
            Ok(MixedValue::Uniform("font".to_owned()))
        );
        assert_eq!(mixed_bool(&OwnedVariant::null()), Ok(MixedValue::Mixed));
        assert_eq!(mixed_bool(&OwnedVariant::empty()), Ok(MixedValue::Empty));
        assert!(mixed_bool(&OwnedVariant::i32(1)).is_err());
    }

    #[test]
    fn formatting_constants_match_the_registered_typelib() {
        assert_eq!(HorizontalAlignment::CENTER.raw(), -4108);
        assert_eq!(VerticalAlignment::TOP.raw(), -4160);
        assert_eq!(UnderlineStyle::DOUBLE.raw(), -4119);
        assert_eq!(FillPattern::SOLID.raw(), 1);
        assert_eq!(BorderLineStyle::NONE.raw(), -4142);
        assert_eq!(BorderWeight::MEDIUM.raw(), -4138);
        assert_eq!(BorderIndex::EDGE_BOTTOM.raw(), 9);
        assert_eq!(ExcelColorIndex::AUTOMATIC.raw(), -4105);
    }

    #[test]
    fn dimensions_reject_nonfinite_and_negative_values() {
        assert!(finite_nonnegative(0.0).is_ok());
        assert!(finite_nonnegative(1.0).is_ok());
        assert!(finite_nonnegative(-1.0).is_err());
        assert!(finite_nonnegative(f64::NAN).is_err());
        assert!(finite_nonnegative(f64::INFINITY).is_err());
        assert!(finite(f64::NEG_INFINITY).is_err());
    }

    #[test]
    #[ignore = "launches a fresh visible Excel process to record private formatting VARIANT tags"]
    fn physical_formatting_variants_are_observed_without_public_exposure()
    -> Result<(), Box<dyn std::error::Error>> {
        let apartment = ComApartment::sta()?;
        let application = OwnedApplication::new(&apartment)?;
        application.set_visible(true)?;
        let mut cleanup_workbook = None;
        let outcome = std::panic::catch_unwind(std::panic::AssertUnwindSafe(
            || -> Result<(), Box<dyn std::error::Error>> {
                let workbook = application.workbooks()?.add()?;
                cleanup_workbook = Some(workbook.clone());
                let worksheet = workbook.worksheets()?.item_by_index(1)?;
                let range = worksheet.range("A1:B2")?;
                let font = range.font()?;
                font.set_name("Arial")?;
                font.set_size(12.0)?;
                font.set_bold(true)?;
                font.set_color(ExcelColor::from_rgb(12, 34, 56))?;
                range.interior()?.set_color(ExcelColor::from_rgb(1, 2, 3))?;
                range.set_number_format("0.00")?;
                range.set_horizontal_alignment(HorizontalAlignment::CENTER)?;
                range.set_wrap_text(true)?;
                range.set_row_height(20.0)?;
                range.set_column_width(16.0)?;
                let bottom = range.borders()?.item(BorderIndex::EDGE_BOTTOM)?;
                bottom.set_line_style(BorderLineStyle::CONTINUOUS)?;
                bottom.set_color(ExcelColor::from_rgb(4, 5, 6))?;
                let automatic_font = {
                    font.set_color_index(ExcelColorIndex::AUTOMATIC)?;
                    font.color()?
                };
                let automatic_interior = {
                    range
                        .interior()?
                        .set_color_index(ExcelColorIndex::AUTOMATIC)?;
                    range.interior()?.color()?
                };
                font.set_color(ExcelColor::from_rgb(12, 34, 56))?;
                range.interior()?.set_color(ExcelColor::from_rgb(1, 2, 3))?;

                let uniform = [
                    (
                        "Font.Bold",
                        object_member_vt(&range, "excel.range.font", "excel.font.bold")?,
                    ),
                    (
                        "Font.Name",
                        object_member_vt(&range, "excel.range.font", "excel.font.name")?,
                    ),
                    (
                        "Font.Size",
                        object_member_vt(&range, "excel.range.font", "excel.font.size")?,
                    ),
                    (
                        "Font.Color",
                        object_member_vt(&range, "excel.range.font", "excel.font.color")?,
                    ),
                    (
                        "Interior.Color",
                        object_member_vt(&range, "excel.range.interior", "excel.interior.color")?,
                    ),
                    (
                        "NumberFormat",
                        range_member_vt(&range, "excel.range.numberformat")?,
                    ),
                    (
                        "HorizontalAlignment",
                        range_member_vt(&range, "excel.range.horizontalalignment")?,
                    ),
                    ("WrapText", range_member_vt(&range, "excel.range.wraptext")?),
                    (
                        "RowHeight",
                        range_member_vt(&range, "excel.range.rowheight")?,
                    ),
                    (
                        "ColumnWidth",
                        range_member_vt(&range, "excel.range.columnwidth")?,
                    ),
                    (
                        "Border.LineStyle",
                        border_member_vt(
                            &range,
                            BorderIndex::EDGE_BOTTOM,
                            "excel.border.linestyle",
                        )?,
                    ),
                ];

                worksheet.range("A1")?.font()?.set_bold(true)?;
                worksheet.range("B1")?.font()?.set_bold(false)?;
                worksheet.range("A1")?.font()?.set_name("Arial")?;
                worksheet.range("B1")?.font()?.set_name("Calibri")?;
                worksheet.range("A1")?.font()?.set_size(11.0)?;
                worksheet.range("B1")?.font()?.set_size(13.0)?;
                worksheet
                    .range("A1")?
                    .font()?
                    .set_color(ExcelColor::from_rgb(1, 2, 3))?;
                worksheet
                    .range("B1")?
                    .font()?
                    .set_color(ExcelColor::from_rgb(3, 2, 1))?;
                worksheet.range("A1")?.set_number_format("0.00")?;
                worksheet.range("B1")?.set_number_format("0%")?;
                worksheet
                    .range("A1")?
                    .set_horizontal_alignment(HorizontalAlignment::LEFT)?;
                worksheet
                    .range("B1")?
                    .set_horizontal_alignment(HorizontalAlignment::RIGHT)?;
                worksheet.range("A1")?.set_wrap_text(true)?;
                worksheet.range("B1")?.set_wrap_text(false)?;
                worksheet.range("A1:B1")?.set_row_height(20.0)?;
                worksheet.range("A2:B2")?.set_row_height(30.0)?;
                worksheet.range("A1:A2")?.set_column_width(14.0)?;
                worksheet.range("B1:B2")?.set_column_width(18.0)?;
                let second_row = worksheet.range("A2:B2")?.entire_row()?;
                second_row.set_row_height(0.0)?;
                let zero_row_hidden =
                    direct_bool_member(&second_row, "excel.range.hidden", "Hidden")?;
                second_row.set_row_height(20.0)?;
                let second_column = worksheet.range("B1:B2")?.entire_column()?;
                second_column.set_column_width(0.0)?;
                let zero_column_hidden =
                    direct_bool_member(&second_column, "excel.range.hidden", "Hidden")?;
                second_column.set_column_width(18.0)?;
                second_row.set_row_height(30.0)?;
                let mixed = [
                    (
                        "Font.Bold",
                        object_member_vt(&range, "excel.range.font", "excel.font.bold")?,
                    ),
                    (
                        "Font.Name",
                        object_member_vt(&range, "excel.range.font", "excel.font.name")?,
                    ),
                    (
                        "Font.Size",
                        object_member_vt(&range, "excel.range.font", "excel.font.size")?,
                    ),
                    (
                        "Font.Color",
                        object_member_vt(&range, "excel.range.font", "excel.font.color")?,
                    ),
                    (
                        "NumberFormat",
                        range_member_vt(&range, "excel.range.numberformat")?,
                    ),
                    (
                        "HorizontalAlignment",
                        range_member_vt(&range, "excel.range.horizontalalignment")?,
                    ),
                    ("WrapText", range_member_vt(&range, "excel.range.wraptext")?),
                    (
                        "RowHeight",
                        range_member_vt(&range, "excel.range.rowheight")?,
                    ),
                    (
                        "ColumnWidth",
                        range_member_vt(&range, "excel.range.columnwidth")?,
                    ),
                ];
                eprintln!(
                    "uniform formatting VARTYPEs={uniform:?}; mixed VARTYPEs={mixed:?}; automatic colors font={automatic_font:?} interior={automatic_interior:?}; zero row hidden={zero_row_hidden}; zero column hidden={zero_column_hidden}"
                );

                cleanup_workbook
                    .take()
                    .expect("fresh workbook must be available for cleanup")
                    .close_without_saving()?;
                Ok(())
            },
        ));
        if let Some(workbook) = cleanup_workbook.take() {
            let _ = workbook.close_without_saving();
        }
        let quit = application.quit();
        match outcome {
            Ok(result) => result?,
            Err(payload) => std::panic::resume_unwind(payload),
        }
        quit?;
        Ok(())
    }

    fn range_member_vt(range: &Range, id: &'static str) -> Result<u16, ExcelComError> {
        Ok(property_get(
            &range.dispatch_object().dispatch,
            member(MemberId::new(id), false),
            vec![],
        )?
        .vt())
    }

    fn object_member_vt(
        range: &Range,
        object_id: &'static str,
        member_id: &'static str,
    ) -> Result<u16, ExcelComError> {
        let mut object = property_get(
            &range.dispatch_object().dispatch,
            member(MemberId::new(object_id), false),
            vec![],
        )?;
        let object = DispatchObject {
            dispatch: object.take_dispatch()?,
            kind: "formatting diagnostic",
        };
        Ok(property_get(
            &object.dispatch,
            member(MemberId::new(member_id), false),
            vec![],
        )?
        .vt())
    }

    fn border_member_vt(
        range: &Range,
        index: BorderIndex,
        member_id: &'static str,
    ) -> Result<u16, ExcelComError> {
        let mut borders = property_get(
            &range.dispatch_object().dispatch,
            member(MemberId::new("excel.range.borders"), false),
            vec![],
        )?;
        let borders = DispatchObject {
            dispatch: borders.take_dispatch()?,
            kind: "formatting diagnostic",
        };
        let mut border = property_get(
            &borders.dispatch,
            member(MemberId::new("excel.borders.item"), false),
            vec![OwnedVariant::i32(index.raw())],
        )?;
        let border = DispatchObject {
            dispatch: border.take_dispatch()?,
            kind: "formatting diagnostic",
        };
        Ok(property_get(
            &border.dispatch,
            member(MemberId::new(member_id), false),
            vec![],
        )?
        .vt())
    }

    fn direct_bool_member(
        range: &Range,
        id: &'static str,
        name: &'static str,
    ) -> Result<bool, ExcelComError> {
        property_get(
            &range.dispatch_object().dispatch,
            MemberDescriptor {
                id: MemberId::new(id),
                name,
                kind: MemberKind::PropertyGet,
            },
            vec![],
        )?
        .as_bool()
        .ok_or(ExcelComError::Unsupported {
            detail: "expected Boolean formatting diagnostic result",
        })
    }
}
