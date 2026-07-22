//! Pre-COM conversion failures.

#[derive(Clone, Debug, Eq, PartialEq)]
pub(crate) enum ConversionError {
    UnsupportedVariantType { vartype: u16 },
    UnsupportedSafeArrayRank { rank: u32 },
    UnsupportedSafeArrayElementType { vartype: u16 },
    #[allow(dead_code)]
    ShapeMismatch { source_rows: usize, source_columns: usize, target_rows: usize, target_columns: usize },
    InvalidElementCount,
    NumericPrecisionLoss,
    NonFiniteNumber,
    EmbeddedNul,
    InvalidDateForPolicy,
    #[allow(dead_code)]
    CurrencyOverflow,
    StringTooLong,
    InvalidExcelErrorScode,
    InvalidUtf16String,
    SafeArrayConstructionFailed,
    SafeArrayElementFailed { row: usize, column: usize },
}
