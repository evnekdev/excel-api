use std::fmt::{Display, Formatter};

use super::Range;

/// Excel's global reference-notation setting.
///
/// A1 is the concise default for the crate's range-selection APIs. R1C1 is
/// explicit through [`super::Worksheet::range_r1c1`] and address options. The
/// wrapper preserves unrecognized raw values so a newer Excel version cannot
/// make an otherwise valid setting unrepresentable.
#[derive(Clone, Copy, Debug, Eq, PartialEq, Hash)]
pub struct ReferenceStyle(i32);

impl ReferenceStyle {
    /// The A1 reference style (`xlA1`, raw value `1`).
    pub const A1: Self = Self(1);
    /// The R1C1 reference style (`xlR1C1`, raw value `-4150`).
    pub const R1C1: Self = Self(-4150);

    /// Preserves an Excel reference-style value without validation.
    pub const fn from_raw(value: i32) -> Self {
        Self(value)
    }

    /// Returns the raw Excel `XlReferenceStyle` value.
    pub const fn raw(self) -> i32 {
        self.0
    }
}

impl Display for ReferenceStyle {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        match *self {
            Self::A1 => formatter.write_str("A1"),
            Self::R1C1 => formatter.write_str("R1C1"),
            Self(value) => write!(formatter, "ReferenceStyle({value})"),
        }
    }
}

/// The absolute/relative conversion mode used by Excel reference APIs.
///
/// This is the forward-compatible representation of Excel's
/// `XlReferenceType` values. Unknown raw values are preserved.
#[derive(Clone, Copy, Debug, Eq, PartialEq, Hash)]
pub struct ReferenceAbsoluteMode(i32);

impl ReferenceAbsoluteMode {
    /// Absolute row and absolute column (`xlAbsolute`, raw value `1`).
    pub const ABSOLUTE: Self = Self(1);
    /// Absolute row and relative column (`xlAbsRowRelColumn`, raw value `2`).
    pub const ABSOLUTE_ROW_RELATIVE_COLUMN: Self = Self(2);
    /// Relative row and absolute column (`xlRelRowAbsColumn`, raw value `3`).
    pub const RELATIVE_ROW_ABSOLUTE_COLUMN: Self = Self(3);
    /// Relative row and relative column (`xlRelative`, raw value `4`).
    pub const RELATIVE: Self = Self(4);

    /// Preserves an Excel reference-type value without validation.
    pub const fn from_raw(value: i32) -> Self {
        Self(value)
    }

    /// Returns the raw Excel `XlReferenceType` value.
    pub const fn raw(self) -> i32 {
        self.0
    }
}

impl Display for ReferenceAbsoluteMode {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        match *self {
            Self::ABSOLUTE => formatter.write_str("absolute"),
            Self::ABSOLUTE_ROW_RELATIVE_COLUMN => {
                formatter.write_str("absolute-row-relative-column")
            }
            Self::RELATIVE_ROW_ABSOLUTE_COLUMN => {
                formatter.write_str("relative-row-absolute-column")
            }
            Self::RELATIVE => formatter.write_str("relative"),
            Self(value) => write!(formatter, "ReferenceAbsoluteMode({value})"),
        }
    }
}

/// Explicit options for [`Range::address_with_options`].
///
/// Omitted boolean fields are sent as Excel `Missing` arguments. Relative
/// R1C1 output needs `relative_to` so Excel, rather than Rust arithmetic,
/// establishes the reference base. The default requests absolute A1 output
/// with no external workbook qualification.
#[derive(Debug)]
pub struct RangeAddressOptions<'a> {
    /// Whether the row should be absolute, or `None` to send `Missing`.
    pub row_absolute: Option<bool>,
    /// Whether the column should be absolute, or `None` to send `Missing`.
    pub column_absolute: Option<bool>,
    /// The requested output notation.
    pub reference_style: ReferenceStyle,
    /// Whether Excel should qualify the workbook and worksheet context.
    pub external: Option<bool>,
    /// The base Range for relative output, or `None` to send `Missing`.
    pub relative_to: Option<&'a Range>,
}

impl Default for RangeAddressOptions<'_> {
    fn default() -> Self {
        Self {
            row_absolute: Some(true),
            column_absolute: Some(true),
            reference_style: ReferenceStyle::A1,
            external: Some(false),
            relative_to: None,
        }
    }
}

/// Optional arguments for [`super::Application::convert_formula`].
///
/// Excel performs the conversion; this crate does not parse or rewrite the
/// formula. `relative_to` supplies Excel's base Range when relative output is
/// requested.
#[derive(Debug, Default)]
pub struct FormulaConversionOptions<'a> {
    /// The requested absolute/relative mode, or `None` to send `Missing`.
    pub to_absolute: Option<ReferenceAbsoluteMode>,
    /// The reference base for a relative conversion, or `None` for `Missing`.
    pub relative_to: Option<&'a Range>,
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn reference_style_preserves_known_and_unknown_raw_values() {
        assert_eq!(ReferenceStyle::A1.raw(), 1);
        assert_eq!(ReferenceStyle::R1C1.raw(), -4150);
        assert_eq!(ReferenceStyle::from_raw(701).raw(), 701);
        assert_eq!(ReferenceStyle::A1.to_string(), "A1");
        assert!(format!("{:?}", ReferenceStyle::R1C1).contains("-4150"));
    }

    #[test]
    fn absolute_mode_preserves_unknown_raw_values() {
        assert_eq!(ReferenceAbsoluteMode::ABSOLUTE.raw(), 1);
        assert_eq!(ReferenceAbsoluteMode::RELATIVE.raw(), 4);
        assert_eq!(ReferenceAbsoluteMode::from_raw(701).raw(), 701);
        assert_eq!(ReferenceAbsoluteMode::RELATIVE.to_string(), "relative");
    }

    #[test]
    fn address_defaults_are_explicit() {
        let options = RangeAddressOptions::default();
        assert_eq!(options.row_absolute, Some(true));
        assert_eq!(options.column_absolute, Some(true));
        assert_eq!(options.reference_style, ReferenceStyle::A1);
        assert_eq!(options.external, Some(false));
        assert!(options.relative_to.is_none());
    }
}
