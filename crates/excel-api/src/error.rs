use core::fmt;

use crate::DecodeError;

/// Safe representation of worksheet errors.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum ExcelError {
    /// `#NULL!`: an invalid intersection.
    Null,
    /// `#DIV/0!`: division by zero.
    Div0,
    /// `#VALUE!`: an invalid value or argument type.
    Value,
    /// `#REF!`: an invalid worksheet reference.
    Ref,
    /// `#NAME?`: an unrecognized name.
    Name,
    /// `#NUM!`: an invalid numeric result.
    Num,
    /// `#N/A`: data is unavailable.
    Na,
    /// `#GETTING_DATA`: an asynchronous data source is still pending.
    GettingData,
}

impl From<excel_api_sys::XlError> for ExcelError {
    fn from(value: excel_api_sys::XlError) -> Self {
        match value {
            excel_api_sys::XlError::Null => Self::Null,
            excel_api_sys::XlError::Div0 => Self::Div0,
            excel_api_sys::XlError::Value => Self::Value,
            excel_api_sys::XlError::Ref => Self::Ref,
            excel_api_sys::XlError::Name => Self::Name,
            excel_api_sys::XlError::Num => Self::Num,
            excel_api_sys::XlError::Na => Self::Na,
            excel_api_sys::XlError::GettingData => Self::GettingData,
        }
    }
}

/// Strict UTF-16 decoding failed because the payload contains an unpaired
/// surrogate.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct Utf16ConversionError;

impl fmt::Display for Utf16ConversionError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("UTF-16 payload contains an unpaired surrogate")
    }
}

impl std::error::Error for Utf16ConversionError {}

/// Failure to construct an owned semantic value directly.
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum OwnedValueError {
    /// Multiplying the requested dimensions overflowed `usize`.
    ArrayShapeOverflow {
        /// Requested row count.
        rows: usize,
        /// Requested column count.
        columns: usize,
    },
    /// The supplied flat element count does not match the dimensions.
    InvalidArrayShape {
        /// Requested row count.
        rows: usize,
        /// Requested column count.
        columns: usize,
        /// Provided flat element count.
        elements: usize,
    },
    /// An owned array attempted to contain another array.
    NestedArrayUnsupported,
}

impl fmt::Display for OwnedValueError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::ArrayShapeOverflow { rows, columns } => {
                write!(formatter, "array shape {rows} x {columns} overflows usize")
            }
            Self::InvalidArrayShape {
                rows,
                columns,
                elements,
            } => write!(
                formatter,
                "array shape {rows} x {columns} does not match {elements} elements"
            ),
            Self::NestedArrayUnsupported => {
                formatter.write_str("nested Excel arrays are unsupported")
            }
        }
    }
}

impl std::error::Error for OwnedValueError {}

/// Failure while converting between callback-borrowed and owned semantic
/// values.
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum ConversionError {
    /// The source kind does not match the requested conversion.
    UnexpectedType {
        /// Human-readable requested kind.
        expected: &'static str,
        /// Human-readable source kind.
        actual: &'static str,
    },
    /// A callback reference requires an explicit reference-preserving path.
    UnsupportedReference,
    /// A UTF-16 payload contained an unpaired surrogate.
    InvalidUtf16,
    /// A floating-point input was NaN or infinite.
    NonFiniteNumber,
    /// A floating-point input had a fractional component.
    NonIntegralNumber,
    /// An integral value did not fit the requested integer type.
    IntegerOutOfRange,
    /// The source array dimensions and elements were inconsistent.
    InvalidArrayShape,
    /// Text exceeded the configured conversion limit.
    StringLimitExceeded {
        /// Observed UTF-16 code-unit count.
        actual: usize,
        /// Configured maximum.
        maximum: usize,
    },
    /// An array exceeded the configured element limit.
    ArrayElementLimitExceeded {
        /// Observed element count.
        actual: usize,
        /// Configured maximum.
        maximum: usize,
    },
    /// Conversion would exceed its aggregate owned-allocation budget.
    AggregateByteLimitExceeded {
        /// Minimum required bytes.
        required: usize,
        /// Configured maximum bytes.
        maximum: usize,
    },
    /// Nested Excel arrays have no supported semantic representation.
    NestedArrayUnsupported,
    /// Recursive conversion exceeded its configured depth.
    ConversionDepthExceeded {
        /// Observed conversion depth.
        depth: usize,
        /// Configured maximum depth.
        maximum: usize,
    },
    /// A callback view became malformed while being decoded.
    BorrowedValueDecode(DecodeError),
}

impl fmt::Display for ConversionError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::UnexpectedType { expected, actual } => {
                write!(formatter, "expected {expected}, received {actual}")
            }
            Self::UnsupportedReference => {
                formatter.write_str("Excel references require explicit coercion")
            }
            Self::InvalidUtf16 => formatter.write_str("Excel text is not valid UTF-16"),
            Self::NonFiniteNumber => formatter.write_str("number is not finite"),
            Self::NonIntegralNumber => formatter.write_str("number is not an integer"),
            Self::IntegerOutOfRange => formatter.write_str("integer is out of range"),
            Self::InvalidArrayShape => formatter.write_str("array shape is invalid"),
            Self::StringLimitExceeded { actual, maximum } => write!(
                formatter,
                "string has {actual} UTF-16 code units; conversion limit is {maximum}"
            ),
            Self::ArrayElementLimitExceeded { actual, maximum } => write!(
                formatter,
                "array has {actual} elements; conversion limit is {maximum}"
            ),
            Self::AggregateByteLimitExceeded { required, maximum } => write!(
                formatter,
                "conversion requires at least {required} bytes; limit is {maximum}"
            ),
            Self::NestedArrayUnsupported => {
                formatter.write_str("nested Excel arrays are unsupported")
            }
            Self::ConversionDepthExceeded { depth, maximum } => write!(
                formatter,
                "conversion depth {depth} exceeds limit {maximum}"
            ),
            Self::BorrowedValueDecode(error) => {
                write!(formatter, "borrowed Excel value became invalid: {error}")
            }
        }
    }
}

impl std::error::Error for ConversionError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::BorrowedValueDecode(error) => Some(error),
            _ => None,
        }
    }
}

impl From<Utf16ConversionError> for ConversionError {
    fn from(_: Utf16ConversionError) -> Self {
        Self::InvalidUtf16
    }
}

impl From<DecodeError> for ConversionError {
    fn from(error: DecodeError) -> Self {
        Self::BorrowedValueDecode(error)
    }
}

impl From<OwnedValueError> for ConversionError {
    fn from(error: OwnedValueError) -> Self {
        match error {
            OwnedValueError::NestedArrayUnsupported => Self::NestedArrayUnsupported,
            OwnedValueError::ArrayShapeOverflow { .. }
            | OwnedValueError::InvalidArrayShape { .. } => Self::InvalidArrayShape,
        }
    }
}

/// Failure to describe a future ABI-compatible Excel return.
///
/// Return errors contain only owned scalar metadata. They never contain raw
/// pointers, callback lifetimes, or partially allocated return storage.
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum ReturnError {
    /// Text exceeds Excel's representable counted-string limit.
    StringTooLongForExcel {
        /// Requested UTF-16 code units.
        actual: usize,
        /// Excel's maximum UTF-16 code units.
        maximum: usize,
    },
    /// Text exceeds the library's configured return limit.
    StringLimitExceeded {
        /// Requested UTF-16 code units.
        actual: usize,
        /// Configured maximum.
        maximum: usize,
    },
    /// Flat elements do not match the requested return dimensions.
    InvalidArrayShape {
        /// Requested rows.
        rows: usize,
        /// Requested columns.
        columns: usize,
        /// Provided element count.
        elements: usize,
    },
    /// Empty multidimensional ABI returns are not representable.
    EmptyArrayUnsupported,
    /// A dimension exceeds the Excel 12 ABI width or bounds.
    ArrayDimensionExceedsAbi {
        /// Requested rows.
        rows: usize,
        /// Requested columns.
        columns: usize,
    },
    /// Multiplying return dimensions overflowed `usize`.
    ArrayElementCountOverflow {
        /// Requested rows.
        rows: usize,
        /// Requested columns.
        columns: usize,
    },
    /// A return array exceeds its configured element budget.
    ArrayElementLimitExceeded {
        /// Requested element count.
        actual: usize,
        /// Configured maximum.
        maximum: usize,
    },
    /// Nested arrays cannot be materialized as Excel 12 return storage.
    NestedArrayUnsupported,
    /// Semantic references are not valid return values in this framework.
    ReferenceUnsupported,
    /// A semantic value variant has no supported return representation.
    UnsupportedSemanticVariant {
        /// Stable semantic variant name.
        variant: &'static str,
    },
    /// Total return-storage byte accounting overflowed.
    TotalByteOverflow,
    /// The return plan exceeds its aggregate allocation budget.
    TotalByteLimitExceeded {
        /// Minimum required bytes.
        required: usize,
        /// Configured maximum bytes.
        maximum: usize,
    },
    /// Return allocation-count accounting overflowed.
    AllocationCountOverflow,
    /// The plan requires too many separate allocations.
    AllocationCountLimitExceeded {
        /// Required allocation count.
        required: usize,
        /// Configured maximum.
        maximum: usize,
    },
    /// Recursive return planning exceeded its configured depth.
    PlanningDepthExceeded {
        /// Observed planning depth.
        depth: usize,
        /// Configured maximum depth.
        maximum: usize,
    },
}

impl fmt::Display for ReturnError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::StringTooLongForExcel { actual, maximum } => write!(
                formatter,
                "return string has {actual} UTF-16 code units; Excel supports at most {maximum}"
            ),
            Self::StringLimitExceeded { actual, maximum } => write!(
                formatter,
                "return string has {actual} UTF-16 code units; project limit is {maximum}"
            ),
            Self::InvalidArrayShape {
                rows,
                columns,
                elements,
            } => write!(
                formatter,
                "return array shape {rows} x {columns} does not match {elements} elements"
            ),
            Self::EmptyArrayUnsupported => formatter.write_str(
                "zero-dimensional xltypeMulti returns are unsupported; return Empty instead",
            ),
            Self::ArrayDimensionExceedsAbi { rows, columns } => write!(
                formatter,
                "return array dimensions {rows} x {columns} exceed the Excel 12 ABI limits"
            ),
            Self::ArrayElementCountOverflow { rows, columns } => write!(
                formatter,
                "return array element count {rows} x {columns} overflows usize"
            ),
            Self::ArrayElementLimitExceeded { actual, maximum } => write!(
                formatter,
                "return array has {actual} elements; project limit is {maximum}"
            ),
            Self::NestedArrayUnsupported => {
                formatter.write_str("nested Excel return arrays are unsupported")
            }
            Self::ReferenceUnsupported => {
                formatter.write_str("Excel references are unsupported return values")
            }
            Self::UnsupportedSemanticVariant { variant } => {
                write!(
                    formatter,
                    "semantic variant {variant} is unsupported for returns"
                )
            }
            Self::TotalByteOverflow => {
                formatter.write_str("return storage byte accounting overflowed usize")
            }
            Self::TotalByteLimitExceeded { required, maximum } => write!(
                formatter,
                "return requires {required} ABI storage bytes; project limit is {maximum}"
            ),
            Self::AllocationCountOverflow => {
                formatter.write_str("return allocation-count accounting overflowed usize")
            }
            Self::AllocationCountLimitExceeded { required, maximum } => write!(
                formatter,
                "return requires {required} allocations; project limit is {maximum}"
            ),
            Self::PlanningDepthExceeded { depth, maximum } => write!(
                formatter,
                "return planning depth {depth} exceeds project limit {maximum}"
            ),
        }
    }
}

impl std::error::Error for ReturnError {}

/// Failure while turning a validated return plan into stable ABI storage.
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum ReturnMaterializationError {
    /// Materialized storage contradicted an already validated plan field.
    PlanStorageInvariantMismatch {
        /// Invariant field name.
        field: &'static str,
        /// Planned amount.
        planned: usize,
        /// Actual materialized amount.
        actual: usize,
    },
    /// UTF-8 to UTF-16 encoding did not match the planned length.
    Utf8EncodedLengthMismatch {
        /// Planned UTF-16 units.
        planned: usize,
        /// Actual UTF-16 units.
        actual: usize,
    },
    /// A counted-string buffer length disagreed with its plan.
    StringBufferLengthMismatch {
        /// Planned units.
        planned: usize,
        /// Actual units.
        actual: usize,
    },
    /// Array data contradicted the shape selected during planning.
    ArrayShapeMismatch {
        /// Planned rows.
        rows: usize,
        /// Planned columns.
        columns: usize,
        /// Actual element count.
        elements: usize,
    },
    /// A planned semantic variant cannot be emitted as ABI storage.
    UnsupportedPlannedValue {
        /// Stable planned variant name.
        variant: &'static str,
    },
    /// A required stable return allocation failed.
    AllocationFailure {
        /// Name of the storage allocation.
        storage: &'static str,
    },
    /// A test-only injected materialization failure.
    InjectedTestFailure {
        /// Test-injected stage name.
        stage: &'static str,
    },
}

impl fmt::Display for ReturnMaterializationError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::PlanStorageInvariantMismatch {
                field,
                planned,
                actual,
            } => write!(
                formatter,
                "return plan {field} is {planned}, but materialized storage requires {actual}"
            ),
            Self::Utf8EncodedLengthMismatch { planned, actual } => write!(
                formatter,
                "planned UTF-8 text requires {planned} UTF-16 units, but encoding produced {actual}"
            ),
            Self::StringBufferLengthMismatch { planned, actual } => write!(
                formatter,
                "planned counted string requires {planned} units, but materialization produced {actual}"
            ),
            Self::ArrayShapeMismatch {
                rows,
                columns,
                elements,
            } => write!(
                formatter,
                "planned return array shape {rows} x {columns} does not match {elements} elements"
            ),
            Self::UnsupportedPlannedValue { variant } => {
                write!(formatter, "planned value {variant} cannot be materialized")
            }
            Self::AllocationFailure { storage } => {
                write!(formatter, "could not allocate {storage} return storage")
            }
            Self::InjectedTestFailure { stage } => {
                write!(formatter, "test injected a failure at {stage}")
            }
        }
    }
}

impl std::error::Error for ReturnMaterializationError {}

#[derive(Debug)]
/// Failure contained by a generated Excel callback thunk.
pub enum ThunkError {
    /// Excel supplied a null argument root where a value was required.
    NullArgument,
    /// The callback-borrowed argument could not be decoded safely.
    Decode(DecodeError),
    /// A decoded argument could not become the requested Rust input.
    Conversion(ConversionError),
    /// User function code deliberately returned a worksheet error.
    Function(ExcelError),
    /// Pointer-free return planning failed.
    ReturnPlanning(ReturnError),
    /// Stable ABI return storage could not be materialized.
    Materialization(ReturnMaterializationError),
}

impl fmt::Display for ThunkError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::NullArgument => formatter.write_str("Excel supplied a null argument pointer"),
            Self::Decode(error) => write!(formatter, "Excel argument decoding failed: {error}"),
            Self::Conversion(error) => {
                write!(formatter, "Excel argument conversion failed: {error}")
            }
            Self::Function(error) => write!(formatter, "worksheet function returned {error:?}"),
            Self::ReturnPlanning(error) => write!(formatter, "return planning failed: {error}"),
            Self::Materialization(error) => {
                write!(formatter, "return materialization failed: {error}")
            }
        }
    }
}

impl std::error::Error for ThunkError {}

impl fmt::Display for DecodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(formatter, "{self:?}")
    }
}

impl std::error::Error for DecodeError {}
