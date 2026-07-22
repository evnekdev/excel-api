//! Errors detected before an Automation call is made.

/// A semantic Automation conversion failure.
///
/// These errors are deterministic and do not indicate that Excel was called.
/// In particular, [`ShapeMismatch`](Self::ShapeMismatch) is returned before a
/// range write when the supplied rectangular array cannot fit the target.
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum ConversionError {
    /// A `VARIANT` tag is outside this crate's supported value vocabulary.
    UnsupportedVariantType {
        /// The unsupported raw `VARTYPE` value.
        vartype: u16,
    },
    /// A SAFEARRAY rank other than two was returned for a Range value.
    UnsupportedSafeArrayRank {
        /// The reported SAFEARRAY rank.
        rank: u32,
    },
    /// A SAFEARRAY did not contain `VARIANT` elements.
    UnsupportedSafeArrayElementType {
        /// The reported element `VARTYPE`.
        vartype: u16,
    },
    /// A source array and target range have different exact dimensions.
    ShapeMismatch {
        /// Source array row count.
        source_rows: usize,
        /// Source array column count.
        source_columns: usize,
        /// Target range row count.
        target_rows: usize,
        /// Target range column count.
        target_columns: usize,
    },
    /// An array dimension or element count cannot be represented safely.
    InvalidElementCount,
    /// An integral result cannot be represented exactly as `f64`.
    NumericPrecisionLoss,
    /// A non-finite floating-point value was supplied or returned.
    NonFiniteNumber,
    /// A string contains an embedded NUL unsupported by this BSTR boundary.
    EmbeddedNul,
    /// A date is invalid for the selected Range member's write policy.
    InvalidDateForPolicy,
    /// A malformed or non-finite OLE Automation date was encountered.
    InvalidDate,
    /// The OS could not construct the required SAFEARRAY.
    SafeArrayConstructionFailed,
    /// One SAFEARRAY element could not be converted.
    SafeArrayElementFailed {
        /// Zero-based row of the failed element.
        row: usize,
        /// Zero-based column of the failed element.
        column: usize,
    },
    /// A returned BSTR did not contain valid UTF-16.
    InvalidUtf16String,
    /// An argument conversion failed before a positional Automation call.
    ArgumentConversion {
        /// Zero-based position in the wrapper's logical Excel argument order.
        position: usize,
        /// The underlying conversion failure.
        source: Box<ConversionError>,
    },
}
