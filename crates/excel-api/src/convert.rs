use crate::{ConversionError, ExcelError, ExcelValue, ExcelValueRef, OptionalValue};

/// Convert one borrowed Excel input into a Rust value.
pub trait FromExcel<'a>: Sized {
    fn from_excel(value: ExcelValueRef<'a>) -> Result<Self, ConversionError>;
}

/// Convert a Rust value into an owned, safe intermediate Excel value.
///
/// A later memory layer will transform this value into an ABI-compatible
/// `ExcelReturn` allocation.
pub trait IntoExcel {
    fn into_excel(self) -> Result<ExcelValue, ConversionError>;
}

impl<'a> FromExcel<'a> for f64 {
    fn from_excel(value: ExcelValueRef<'a>) -> Result<Self, ConversionError> {
        match value {
            ExcelValueRef::Number(value) => Ok(value),
            other => Err(ConversionError::UnexpectedType {
                expected: "number",
                actual: other.kind_name(),
            }),
        }
    }
}

impl IntoExcel for f64 {
    fn into_excel(self) -> Result<ExcelValue, ConversionError> {
        Ok(ExcelValue::Number(self))
    }
}

impl<'a> FromExcel<'a> for bool {
    fn from_excel(value: ExcelValueRef<'a>) -> Result<Self, ConversionError> {
        match value {
            ExcelValueRef::Boolean(value) => Ok(value),
            other => Err(ConversionError::UnexpectedType {
                expected: "boolean",
                actual: other.kind_name(),
            }),
        }
    }
}

impl IntoExcel for bool {
    fn into_excel(self) -> Result<ExcelValue, ConversionError> {
        Ok(ExcelValue::Boolean(self))
    }
}

impl<'a> FromExcel<'a> for String {
    fn from_excel(value: ExcelValueRef<'a>) -> Result<Self, ConversionError> {
        match value {
            ExcelValueRef::Text(value) => Ok(value.to_owned()),
            other => Err(ConversionError::UnexpectedType {
                expected: "text",
                actual: other.kind_name(),
            }),
        }
    }
}

impl IntoExcel for String {
    fn into_excel(self) -> Result<ExcelValue, ConversionError> {
        Ok(ExcelValue::Text(self))
    }
}

impl IntoExcel for &str {
    fn into_excel(self) -> Result<ExcelValue, ConversionError> {
        Ok(ExcelValue::Text(self.to_owned()))
    }
}

impl<'a, T: FromExcel<'a>> FromExcel<'a> for Option<T> {
    fn from_excel(value: ExcelValueRef<'a>) -> Result<Self, ConversionError> {
        match value {
            ExcelValueRef::Missing | ExcelValueRef::Empty => Ok(None),
            value => T::from_excel(value).map(Some),
        }
    }
}

impl<'a, T: FromExcel<'a>> FromExcel<'a> for OptionalValue<T> {
    fn from_excel(value: ExcelValueRef<'a>) -> Result<Self, ConversionError> {
        match value {
            ExcelValueRef::Missing => Ok(Self::Missing),
            ExcelValueRef::Empty => Ok(Self::Empty),
            value => T::from_excel(value).map(Self::Value),
        }
    }
}

impl IntoExcel for ExcelError {
    fn into_excel(self) -> Result<ExcelValue, ConversionError> {
        Ok(ExcelValue::Error(self))
    }
}

macro_rules! impl_integer_conversion {
    ($($integer:ty),* $(,)?) => {
        $(
            impl<'a> FromExcel<'a> for $integer {
                fn from_excel(value: ExcelValueRef<'a>) -> Result<Self, ConversionError> {
                    let value = f64::from_excel(value)?;
                    if value.fract() != 0.0 {
                        return Err(ConversionError::NonIntegralNumber);
                    }
                    if value < <$integer>::MIN as f64 || value > <$integer>::MAX as f64 {
                        return Err(ConversionError::IntegerOutOfRange);
                    }
                    Ok(value as $integer)
                }
            }

            impl IntoExcel for $integer {
                fn into_excel(self) -> Result<ExcelValue, ConversionError> {
                    Ok(ExcelValue::Number(self as f64))
                }
            }
        )*
    };
}

impl_integer_conversion!(i16, i32, u16, u32);

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn option_accepts_missing_and_empty() {
        assert_eq!(Option::<f64>::from_excel(ExcelValueRef::Missing), Ok(None));
        assert_eq!(Option::<f64>::from_excel(ExcelValueRef::Empty), Ok(None));
    }

    #[test]
    fn integer_conversion_is_checked() {
        assert_eq!(
            i32::from_excel(ExcelValueRef::Number(1.5)),
            Err(ConversionError::NonIntegralNumber)
        );
    }
}
