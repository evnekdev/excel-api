use core::fmt;

use crate::DecodeError;

/// Safe representation of worksheet errors.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum ExcelError {
    Null,
    Div0,
    Value,
    Ref,
    Name,
    Num,
    Na,
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
    ArrayShapeOverflow {
        rows: usize,
        columns: usize,
    },
    InvalidArrayShape {
        rows: usize,
        columns: usize,
        elements: usize,
    },
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
#[derive(Debug, Eq, PartialEq)]
pub enum ConversionError {
    UnexpectedType {
        expected: &'static str,
        actual: &'static str,
    },
    UnsupportedReference,
    InvalidUtf16,
    NonFiniteNumber,
    NonIntegralNumber,
    IntegerOutOfRange,
    InvalidArrayShape,
    StringLimitExceeded {
        actual: usize,
        maximum: usize,
    },
    ArrayElementLimitExceeded {
        actual: usize,
        maximum: usize,
    },
    AggregateByteLimitExceeded {
        required: usize,
        maximum: usize,
    },
    NestedArrayUnsupported,
    ConversionDepthExceeded {
        depth: usize,
        maximum: usize,
    },
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

impl fmt::Display for DecodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(formatter, "{self:?}")
    }
}

impl std::error::Error for DecodeError {}
