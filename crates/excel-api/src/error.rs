use core::fmt;

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

/// Failure while converting between Excel and Rust values.
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum ConversionError {
    UnexpectedType {
        expected: &'static str,
        actual: &'static str,
    },
    NonIntegralNumber,
    IntegerOutOfRange,
    UnsupportedValue(&'static str),
}

impl fmt::Display for ConversionError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::UnexpectedType { expected, actual } => {
                write!(formatter, "expected {expected}, received {actual}")
            }
            Self::NonIntegralNumber => formatter.write_str("number is not an integer"),
            Self::IntegerOutOfRange => formatter.write_str("integer is out of range"),
            Self::UnsupportedValue(value) => write!(formatter, "unsupported Excel value: {value}"),
        }
    }
}

impl std::error::Error for ConversionError {}
