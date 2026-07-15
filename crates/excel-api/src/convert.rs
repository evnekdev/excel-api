use core::mem::size_of;

use crate::{
    ConversionError, ExcelArray, ExcelArrayView, ExcelError, ExcelString, ExcelValue,
    ExcelValueRef, OptionalValue,
};

/// Resource limits applied before callback-owned data is copied.
///
/// Defaults deliberately allow far less than Excel's maximum worksheet grid:
/// 65,536 elements and 16 MiB of conservatively counted destination storage.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct ConversionLimits {
    /// Maximum UTF-16 code units copied for one string.
    pub max_string_code_units: usize,
    /// Maximum elements copied from one rectangular array.
    pub max_array_elements: usize,
    /// Maximum conservatively counted destination bytes.
    pub max_aggregate_bytes: usize,
    /// Maximum supported conversion nesting depth.
    pub max_depth: usize,
}

impl ConversionLimits {
    /// Conservative conversion bounds used by [`Default`].
    pub const DEFAULT: Self = Self {
        max_string_code_units: excel_api_sys::EXCEL12_MAX_STRING_CODE_UNITS,
        max_array_elements: 65_536,
        max_aggregate_bytes: 16 * 1024 * 1024,
        max_depth: 8,
    };
}

impl Default for ConversionLimits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// Convert one borrowed Excel input into a Rust value without invoking Excel
/// coercion rules.
pub trait FromExcel<'call>: Sized {
    /// Converts a callback-borrowed value without invoking Excel coercion.
    ///
    /// Implementations must not retain callback-scoped pointers in `Self`.
    fn from_excel(value: ExcelValueRef<'call>) -> Result<Self, ConversionError>;
}

/// Convert a Rust value into an owned semantic Excel value.
///
/// A later memory layer will transform this value into an ABI-compatible
/// return allocation.
pub trait IntoExcel {
    /// Produces an owned semantic Excel value suitable for later return planning.
    fn into_excel(self) -> Result<ExcelValue, ConversionError>;
}

impl ExcelValue {
    /// Deep-copies a callback value using [`ConversionLimits::default`].
    pub fn from_borrowed(value: ExcelValueRef<'_>) -> Result<Self, ConversionError> {
        Self::from_borrowed_with_limits(value, &ConversionLimits::default())
    }

    /// Deep-copies a callback value after a non-allocating preflight pass.
    pub fn from_borrowed_with_limits(
        value: ExcelValueRef<'_>,
        limits: &ConversionLimits,
    ) -> Result<Self, ConversionError> {
        let required = preflight_value(&value, limits, 0)?;
        enforce_aggregate_limit(required, limits)?;
        materialize_value(value, 0)
    }
}

impl ExcelString {
    /// Copies callback UTF-16 using [`ConversionLimits::default`].
    pub fn from_borrowed(value: ExcelValueRef<'_>) -> Result<Self, ConversionError> {
        Self::from_borrowed_with_limits(value, &ConversionLimits::default())
    }

    /// Copies callback UTF-16 directly, without an intermediate UTF-8 string.
    pub fn from_borrowed_with_limits(
        value: ExcelValueRef<'_>,
        limits: &ConversionLimits,
    ) -> Result<Self, ConversionError> {
        let ExcelValueRef::Text(value) = value else {
            return Err(unexpected("text", &value));
        };
        let required = preflight_string(value.len(), size_of::<u16>(), limits)?;
        enforce_aggregate_limit(required, limits)?;
        Ok(value.into())
    }
}

impl ExcelArray {
    /// Deep-copies a callback multi using [`ConversionLimits::default`].
    pub fn from_borrowed(value: ExcelValueRef<'_>) -> Result<Self, ConversionError> {
        Self::from_borrowed_with_limits(value, &ConversionLimits::default())
    }

    /// Deep-copies a callback multi after checking shape and all configured
    /// limits before allocating destination storage.
    pub fn from_borrowed_with_limits(
        value: ExcelValueRef<'_>,
        limits: &ConversionLimits,
    ) -> Result<Self, ConversionError> {
        let ExcelValueRef::Array(value) = value else {
            return Err(unexpected("array", &value));
        };
        let required = preflight_array(&value, limits, 0)?;
        enforce_aggregate_limit(required, limits)?;
        materialize_array(value, 0)
    }
}

impl<'call> FromExcel<'call> for ExcelValue {
    fn from_excel(value: ExcelValueRef<'call>) -> Result<Self, ConversionError> {
        Self::from_borrowed(value)
    }
}

impl<'call> FromExcel<'call> for ExcelString {
    fn from_excel(value: ExcelValueRef<'call>) -> Result<Self, ConversionError> {
        Self::from_borrowed(value)
    }
}

impl<'call> FromExcel<'call> for String {
    fn from_excel(value: ExcelValueRef<'call>) -> Result<Self, ConversionError> {
        let ExcelValueRef::Text(value) = value else {
            return Err(unexpected("text", &value));
        };
        let limits = ConversionLimits::default();
        let required = preflight_string(value.len(), 3, &limits)?;
        enforce_aggregate_limit(required, &limits)?;
        String::from_utf16(value.as_utf16()).map_err(|_| ConversionError::InvalidUtf16)
    }
}

impl<'call> FromExcel<'call> for ExcelArray {
    fn from_excel(value: ExcelValueRef<'call>) -> Result<Self, ConversionError> {
        Self::from_borrowed(value)
    }
}

impl<'call> FromExcel<'call> for f64 {
    fn from_excel(value: ExcelValueRef<'call>) -> Result<Self, ConversionError> {
        match value {
            ExcelValueRef::Number(value) => Ok(value),
            ExcelValueRef::Integer(value) => Ok(f64::from(value)),
            other => Err(unexpected("number", &other)),
        }
    }
}

impl<'call> FromExcel<'call> for bool {
    fn from_excel(value: ExcelValueRef<'call>) -> Result<Self, ConversionError> {
        match value {
            ExcelValueRef::Boolean(value) => Ok(value),
            other => Err(unexpected("boolean", &other)),
        }
    }
}

impl<'call> FromExcel<'call> for ExcelError {
    fn from_excel(value: ExcelValueRef<'call>) -> Result<Self, ConversionError> {
        match value {
            ExcelValueRef::Error(value) => Ok(value),
            other => Err(unexpected("error", &other)),
        }
    }
}

impl<'call, T: FromExcel<'call>> FromExcel<'call> for Option<T> {
    /// Missing and empty values both map to `None`. Use `OptionalValue<T>` when
    /// that distinction must be retained.
    fn from_excel(value: ExcelValueRef<'call>) -> Result<Self, ConversionError> {
        match value {
            ExcelValueRef::Missing(_) | ExcelValueRef::Nil(_) => Ok(None),
            value => T::from_excel(value).map(Some),
        }
    }
}

impl<'call, T: FromExcel<'call>> FromExcel<'call> for OptionalValue<T> {
    fn from_excel(value: ExcelValueRef<'call>) -> Result<Self, ConversionError> {
        match value {
            ExcelValueRef::Missing(_) => Ok(Self::Missing),
            ExcelValueRef::Nil(_) => Ok(Self::Empty),
            value => T::from_excel(value).map(Self::Value),
        }
    }
}

macro_rules! impl_integer_from_excel {
    ($($integer:ty),* $(,)?) => {
        $(
            impl<'call> FromExcel<'call> for $integer {
                fn from_excel(value: ExcelValueRef<'call>) -> Result<Self, ConversionError> {
                    match value {
                        ExcelValueRef::Integer(value) => {
                            Self::try_from(value).map_err(|_| ConversionError::IntegerOutOfRange)
                        }
                        ExcelValueRef::Number(value) => checked_float_to_integer(
                            value,
                            <$integer>::MIN as f64,
                            <$integer>::MAX as f64,
                        ),
                        other => Err(unexpected("integer", &other)),
                    }
                }
            }
        )*
    };
}

impl_integer_from_excel!(i16, i32, u16, u32);

fn checked_float_to_integer<T>(value: f64, minimum: f64, maximum: f64) -> Result<T, ConversionError>
where
    T: TryFrom<i64>,
{
    if !value.is_finite() {
        return Err(ConversionError::NonFiniteNumber);
    }
    if value.fract() != 0.0 {
        return Err(ConversionError::NonIntegralNumber);
    }

    // Every currently supported target is at most 32 bits, so all target
    // bounds and integral inputs are represented exactly by f64 and i64.
    if value < minimum || value > maximum {
        return Err(ConversionError::IntegerOutOfRange);
    }
    let integer = value as i64;
    T::try_from(integer).map_err(|_| ConversionError::IntegerOutOfRange)
}

impl IntoExcel for ExcelValue {
    fn into_excel(self) -> Result<ExcelValue, ConversionError> {
        Ok(self)
    }
}

impl IntoExcel for ExcelArray {
    fn into_excel(self) -> Result<ExcelValue, ConversionError> {
        Ok(ExcelValue::Array(self))
    }
}

impl IntoExcel for ExcelString {
    fn into_excel(self) -> Result<ExcelValue, ConversionError> {
        Ok(ExcelValue::Text(self))
    }
}

impl IntoExcel for f64 {
    fn into_excel(self) -> Result<ExcelValue, ConversionError> {
        Ok(ExcelValue::Number(self))
    }
}

impl IntoExcel for bool {
    fn into_excel(self) -> Result<ExcelValue, ConversionError> {
        Ok(ExcelValue::Boolean(self))
    }
}

impl IntoExcel for String {
    fn into_excel(self) -> Result<ExcelValue, ConversionError> {
        Ok(ExcelValue::Text(self.into()))
    }
}

impl IntoExcel for &str {
    fn into_excel(self) -> Result<ExcelValue, ConversionError> {
        Ok(ExcelValue::Text(self.into()))
    }
}

impl IntoExcel for ExcelError {
    fn into_excel(self) -> Result<ExcelValue, ConversionError> {
        Ok(ExcelValue::Error(self))
    }
}

macro_rules! impl_i32_backed_into_excel {
    ($($integer:ty),* $(,)?) => {
        $(
            impl IntoExcel for $integer {
                fn into_excel(self) -> Result<ExcelValue, ConversionError> {
                    Ok(ExcelValue::Integer(i32::from(self)))
                }
            }
        )*
    };
}

impl_i32_backed_into_excel!(i16, i32, u16);

impl IntoExcel for u32 {
    fn into_excel(self) -> Result<ExcelValue, ConversionError> {
        match i32::try_from(self) {
            Ok(value) => Ok(ExcelValue::Integer(value)),
            Err(_) => Ok(ExcelValue::Number(f64::from(self))),
        }
    }
}

fn preflight_value(
    value: &ExcelValueRef<'_>,
    limits: &ConversionLimits,
    depth: usize,
) -> Result<usize, ConversionError> {
    enforce_depth(depth, limits)?;
    match value {
        ExcelValueRef::Text(value) => preflight_string(value.len(), size_of::<u16>(), limits),
        ExcelValueRef::Reference(_) => Err(ConversionError::UnsupportedReference),
        ExcelValueRef::Array(value) => {
            if depth != 0 {
                return Err(ConversionError::NestedArrayUnsupported);
            }
            preflight_array(value, limits, depth)
        }
        ExcelValueRef::Number(_)
        | ExcelValueRef::Boolean(_)
        | ExcelValueRef::Error(_)
        | ExcelValueRef::Missing(_)
        | ExcelValueRef::Nil(_)
        | ExcelValueRef::Integer(_) => Ok(0),
    }
}

fn preflight_array(
    array: &ExcelArrayView<'_>,
    limits: &ConversionLimits,
    depth: usize,
) -> Result<usize, ConversionError> {
    let element_depth = depth
        .checked_add(1)
        .ok_or(ConversionError::ConversionDepthExceeded {
            depth: usize::MAX,
            maximum: limits.max_depth,
        })?;
    enforce_depth(element_depth, limits)?;

    let (rows, columns) = array.dimensions();
    let count = rows
        .checked_mul(columns)
        .ok_or(ConversionError::InvalidArrayShape)?;
    if count > limits.max_array_elements {
        return Err(ConversionError::ArrayElementLimitExceeded {
            actual: count,
            maximum: limits.max_array_elements,
        });
    }

    let mut required = checked_storage_bytes(count, size_of::<ExcelValue>(), limits)?;
    enforce_aggregate_limit(required, limits)?;
    for row in 0..rows {
        for column in 0..columns {
            let value = array
                .get(row, column)?
                .ok_or(ConversionError::InvalidArrayShape)?;
            let element_bytes = preflight_value(&value, limits, element_depth)?;
            required = required.checked_add(element_bytes).ok_or(
                ConversionError::AggregateByteLimitExceeded {
                    required: usize::MAX,
                    maximum: limits.max_aggregate_bytes,
                },
            )?;
            enforce_aggregate_limit(required, limits)?;
        }
    }
    Ok(required)
}

fn preflight_string(
    code_units: usize,
    bytes_per_unit: usize,
    limits: &ConversionLimits,
) -> Result<usize, ConversionError> {
    if code_units > limits.max_string_code_units {
        return Err(ConversionError::StringLimitExceeded {
            actual: code_units,
            maximum: limits.max_string_code_units,
        });
    }
    checked_storage_bytes(code_units, bytes_per_unit, limits)
}

fn checked_storage_bytes(
    count: usize,
    element_size: usize,
    limits: &ConversionLimits,
) -> Result<usize, ConversionError> {
    count
        .checked_mul(element_size)
        .ok_or(ConversionError::AggregateByteLimitExceeded {
            required: usize::MAX,
            maximum: limits.max_aggregate_bytes,
        })
}

fn enforce_aggregate_limit(
    required: usize,
    limits: &ConversionLimits,
) -> Result<(), ConversionError> {
    if required > limits.max_aggregate_bytes {
        Err(ConversionError::AggregateByteLimitExceeded {
            required,
            maximum: limits.max_aggregate_bytes,
        })
    } else {
        Ok(())
    }
}

fn enforce_depth(depth: usize, limits: &ConversionLimits) -> Result<(), ConversionError> {
    if depth > limits.max_depth {
        Err(ConversionError::ConversionDepthExceeded {
            depth,
            maximum: limits.max_depth,
        })
    } else {
        Ok(())
    }
}

fn materialize_value(
    value: ExcelValueRef<'_>,
    depth: usize,
) -> Result<ExcelValue, ConversionError> {
    match value {
        ExcelValueRef::Number(value) => Ok(ExcelValue::Number(value)),
        ExcelValueRef::Text(value) => Ok(ExcelValue::Text(value.into())),
        ExcelValueRef::Boolean(value) => Ok(ExcelValue::Boolean(value)),
        ExcelValueRef::Reference(_) => Err(ConversionError::UnsupportedReference),
        ExcelValueRef::Error(value) => Ok(ExcelValue::Error(value)),
        ExcelValueRef::Array(value) => {
            if depth != 0 {
                return Err(ConversionError::NestedArrayUnsupported);
            }
            materialize_array(value, depth).map(ExcelValue::Array)
        }
        ExcelValueRef::Missing(_) => Ok(ExcelValue::Missing),
        ExcelValueRef::Nil(_) => Ok(ExcelValue::Empty),
        ExcelValueRef::Integer(value) => Ok(ExcelValue::Integer(value)),
    }
}

fn materialize_array(
    array: ExcelArrayView<'_>,
    depth: usize,
) -> Result<ExcelArray, ConversionError> {
    let (rows, columns) = array.dimensions();
    let count = rows
        .checked_mul(columns)
        .ok_or(ConversionError::InvalidArrayShape)?;
    let mut values = Vec::with_capacity(count);
    for row in 0..rows {
        for column in 0..columns {
            let value = array
                .get(row, column)?
                .ok_or(ConversionError::InvalidArrayShape)?;
            values.push(materialize_value(value, depth + 1)?);
        }
    }
    ExcelArray::new(rows, columns, values).map_err(Into::into)
}

fn unexpected(expected: &'static str, actual: &ExcelValueRef<'_>) -> ConversionError {
    ConversionError::UnexpectedType {
        expected,
        actual: actual.kind_name(),
    }
}

#[cfg(test)]
mod tests {
    use excel_api_sys::{
        XLOPER12, XLOPER12Array, XLOPER12SRef, XLOPER12Value, XLREF12, xltypeBool, xltypeErr,
        xltypeInt, xltypeMissing, xltypeMulti, xltypeNil, xltypeNum, xltypeSRef, xltypeStr,
    };

    use crate::{RawExcelValue, Utf16ConversionError};

    use super::*;

    fn raw(value: XLOPER12Value, xltype: u32) -> XLOPER12 {
        XLOPER12 { val: value, xltype }
    }

    fn decode(value: &XLOPER12) -> ExcelValueRef<'_> {
        // SAFETY: each fixture keeps all reachable SDK storage initialized,
        // immutable, and alive while the returned callback view is used.
        unsafe { RawExcelValue::from_callback(value) }
            .decode()
            .unwrap()
    }

    #[test]
    fn deep_copy_preserves_scalars_missing_empty_and_integer() {
        let fixtures = [
            (
                raw(XLOPER12Value { num: 1.25 }, xltypeNum),
                ExcelValue::Number(1.25),
            ),
            (
                raw(XLOPER12Value { w: -7 }, xltypeInt),
                ExcelValue::Integer(-7),
            ),
            (
                raw(XLOPER12Value { xbool: 1 }, xltypeBool),
                ExcelValue::Boolean(true),
            ),
            (
                raw(
                    XLOPER12Value {
                        err: excel_api_sys::xlerrRef,
                    },
                    xltypeErr,
                ),
                ExcelValue::Error(ExcelError::Ref),
            ),
            (
                raw(XLOPER12Value { w: 0 }, xltypeMissing),
                ExcelValue::Missing,
            ),
            (raw(XLOPER12Value { w: 0 }, xltypeNil), ExcelValue::Empty),
        ];

        for (fixture, expected) in &fixtures {
            assert_eq!(
                ExcelValue::from_excel(decode(fixture)),
                Ok(expected.clone())
            );
        }
    }

    #[test]
    fn borrowed_text_copy_is_independent_and_preserves_code_units() {
        let mut storage = [4_u16, b'A' as u16, 0, 0xD800, 0xDC00];
        let raw = raw(
            XLOPER12Value {
                str: storage.as_mut_ptr(),
            },
            xltypeStr,
        );
        let owned = ExcelString::from_excel(decode(&raw)).unwrap();
        storage[1..].fill(b'Z' as u16);

        assert_eq!(owned.as_utf16(), &[b'A' as u16, 0, 0xD800, 0xDC00]);
    }

    #[test]
    fn strict_string_conversion_rejects_unpaired_surrogates() {
        let mut storage = [1_u16, 0xD800];
        let raw = raw(
            XLOPER12Value {
                str: storage.as_mut_ptr(),
            },
            xltypeStr,
        );
        assert_eq!(
            String::from_excel(decode(&raw)),
            Err(ConversionError::InvalidUtf16)
        );

        let owned = ExcelString::from_utf16_units([0xD800].to_vec());
        assert_eq!(String::try_from(&owned), Err(Utf16ConversionError));
    }

    #[test]
    fn deep_copied_mixed_array_is_independent_and_row_major() {
        let mut text = [3_u16, b'X' as u16, 0, 0xD800];
        let mut elements = [
            raw(XLOPER12Value { num: 1.5 }, xltypeNum),
            raw(XLOPER12Value { w: 2 }, xltypeInt),
            raw(
                XLOPER12Value {
                    str: text.as_mut_ptr(),
                },
                xltypeStr,
            ),
            raw(
                XLOPER12Value {
                    err: excel_api_sys::xlerrNA,
                },
                xltypeErr,
            ),
        ];
        let root = raw(
            XLOPER12Value {
                array: XLOPER12Array {
                    lparray: elements.as_mut_ptr(),
                    rows: 2,
                    columns: 2,
                },
            },
            xltypeMulti,
        );

        let owned = ExcelArray::from_excel(decode(&root)).unwrap();
        text[1..].fill(b'Z' as u16);
        elements[0] = raw(XLOPER12Value { num: 99.0 }, xltypeNum);
        assert_eq!(elements[0].xltype, xltypeNum);
        let ExcelValueRef::Array(overwritten_source) = decode(&root) else {
            panic!("expected source array");
        };
        assert!(matches!(
            overwritten_source.get(0, 0),
            Ok(Some(ExcelValueRef::Number(99.0)))
        ));

        assert_eq!(owned.get(0, 0), Some(&ExcelValue::Number(1.5)));
        assert_eq!(owned.get(0, 1), Some(&ExcelValue::Integer(2)));
        assert_eq!(
            owned.get(1, 0),
            Some(&ExcelValue::Text(ExcelString::from_utf16_units(
                [b'X' as u16, 0, 0xD800].to_vec()
            )))
        );
        assert_eq!(owned.get(1, 1), Some(&ExcelValue::Error(ExcelError::Na)));
    }

    #[test]
    fn references_are_rejected_without_coercion() {
        let reference = raw(
            XLOPER12Value {
                sref: XLOPER12SRef {
                    count: 1,
                    reference: XLREF12::default(),
                },
            },
            xltypeSRef,
        );
        assert_eq!(
            ExcelValue::from_excel(decode(&reference)),
            Err(ConversionError::UnsupportedReference)
        );
    }

    #[test]
    fn conversion_limits_are_enforced_before_copying() {
        let mut text = [3_u16, b'A' as u16, b'B' as u16, b'C' as u16];
        let raw_text = raw(
            XLOPER12Value {
                str: text.as_mut_ptr(),
            },
            xltypeStr,
        );
        let limits = ConversionLimits {
            max_string_code_units: 2,
            ..ConversionLimits::default()
        };
        assert_eq!(
            ExcelValue::from_borrowed_with_limits(decode(&raw_text), &limits),
            Err(ConversionError::StringLimitExceeded {
                actual: 3,
                maximum: 2,
            })
        );

        let mut elements = [
            raw(XLOPER12Value { num: 1.0 }, xltypeNum),
            raw(XLOPER12Value { num: 2.0 }, xltypeNum),
        ];
        let root = raw(
            XLOPER12Value {
                array: XLOPER12Array {
                    lparray: elements.as_mut_ptr(),
                    rows: 1,
                    columns: 2,
                },
            },
            xltypeMulti,
        );
        let limits = ConversionLimits {
            max_array_elements: 1,
            ..ConversionLimits::default()
        };
        assert_eq!(
            ExcelArray::from_borrowed_with_limits(decode(&root), &limits),
            Err(ConversionError::ArrayElementLimitExceeded {
                actual: 2,
                maximum: 1,
            })
        );

        let limits = ConversionLimits {
            max_aggregate_bytes: size_of::<ExcelValue>() * 2 - 1,
            ..ConversionLimits::default()
        };
        assert!(matches!(
            ExcelArray::from_borrowed_with_limits(decode(&root), &limits),
            Err(ConversionError::AggregateByteLimitExceeded { .. })
        ));

        let mut text_element = raw(
            XLOPER12Value {
                str: text.as_mut_ptr(),
            },
            xltypeStr,
        );
        let text_array = raw(
            XLOPER12Value {
                array: XLOPER12Array {
                    lparray: &mut text_element,
                    rows: 1,
                    columns: 1,
                },
            },
            xltypeMulti,
        );
        let required = size_of::<ExcelValue>() + 3 * size_of::<u16>();
        let limits = ConversionLimits {
            max_aggregate_bytes: required - 1,
            ..ConversionLimits::default()
        };
        assert_eq!(
            ExcelArray::from_borrowed_with_limits(decode(&text_array), &limits),
            Err(ConversionError::AggregateByteLimitExceeded {
                required,
                maximum: required - 1,
            })
        );

        let limits = ConversionLimits {
            max_depth: 0,
            ..ConversionLimits::default()
        };
        assert_eq!(
            ExcelArray::from_borrowed_with_limits(decode(&root), &limits),
            Err(ConversionError::ConversionDepthExceeded {
                depth: 1,
                maximum: 0,
            })
        );
    }

    #[test]
    fn numeric_conversions_follow_checked_policy() {
        assert_eq!(f64::from_excel(ExcelValueRef::Integer(-7)), Ok(-7.0));
        assert_eq!(i32::from_excel(ExcelValueRef::Integer(-7)), Ok(-7));
        assert_eq!(
            u32::from_excel(ExcelValueRef::Integer(-1)),
            Err(ConversionError::IntegerOutOfRange)
        );
        assert_eq!(i32::from_excel(ExcelValueRef::Number(42.0)), Ok(42));
        assert_eq!(
            i32::from_excel(ExcelValueRef::Number(1.5)),
            Err(ConversionError::NonIntegralNumber)
        );
        assert_eq!(
            i32::from_excel(ExcelValueRef::Number(f64::NAN)),
            Err(ConversionError::NonFiniteNumber)
        );
        assert_eq!(
            i32::from_excel(ExcelValueRef::Number(f64::INFINITY)),
            Err(ConversionError::NonFiniteNumber)
        );
        assert_eq!(
            i16::from_excel(ExcelValueRef::Number(f64::from(i16::MIN))),
            Ok(i16::MIN)
        );
        assert_eq!(
            i16::from_excel(ExcelValueRef::Number(f64::from(i16::MAX))),
            Ok(i16::MAX)
        );
        assert_eq!(
            i32::from_excel(ExcelValueRef::Number(f64::from(i32::MIN))),
            Ok(i32::MIN)
        );
        assert_eq!(
            i32::from_excel(ExcelValueRef::Number(f64::from(i32::MAX))),
            Ok(i32::MAX)
        );
        assert_eq!(u16::from_excel(ExcelValueRef::Number(0.0)), Ok(u16::MIN));
        assert_eq!(
            u16::from_excel(ExcelValueRef::Number(f64::from(u16::MAX))),
            Ok(u16::MAX)
        );
        assert_eq!(u32::from_excel(ExcelValueRef::Number(0.0)), Ok(u32::MIN));
        assert_eq!(
            u32::from_excel(ExcelValueRef::Number(f64::from(u32::MAX))),
            Ok(u32::MAX)
        );
        assert_eq!(
            i16::from_excel(ExcelValueRef::Number(f64::from(i16::MAX) + 1.0)),
            Err(ConversionError::IntegerOutOfRange)
        );
        assert_eq!(
            u32::from_excel(ExcelValueRef::Number(-1.0)),
            Err(ConversionError::IntegerOutOfRange)
        );
    }

    #[test]
    fn boolean_and_error_conversions_require_exact_types() {
        assert_eq!(bool::from_excel(ExcelValueRef::Boolean(true)), Ok(true));
        assert_eq!(
            ExcelError::from_excel(ExcelValueRef::Error(ExcelError::Div0)),
            Ok(ExcelError::Div0)
        );
        assert!(matches!(
            bool::from_excel(ExcelValueRef::Number(1.0)),
            Err(ConversionError::UnexpectedType {
                expected: "boolean",
                actual: "number",
            })
        ));
    }

    #[test]
    fn into_excel_preserves_i32_backed_integer_semantics() {
        assert_eq!(7_i16.into_excel(), Ok(ExcelValue::Integer(7)));
        assert_eq!(7_u16.into_excel(), Ok(ExcelValue::Integer(7)));
        assert_eq!(7_i32.into_excel(), Ok(ExcelValue::Integer(7)));
        assert_eq!(
            u32::MAX.into_excel(),
            Ok(ExcelValue::Number(f64::from(u32::MAX)))
        );
    }

    #[test]
    fn option_and_optional_value_have_documented_empty_policy() {
        let missing = raw(XLOPER12Value { w: 0 }, xltypeMissing);
        let empty = raw(XLOPER12Value { w: 0 }, xltypeNil);
        let integer = raw(XLOPER12Value { w: 3 }, xltypeInt);

        assert_eq!(Option::<i32>::from_excel(decode(&missing)), Ok(None));
        assert_eq!(Option::<i32>::from_excel(decode(&empty)), Ok(None));
        assert_eq!(
            OptionalValue::<i32>::from_excel(decode(&missing)),
            Ok(OptionalValue::Missing)
        );
        assert_eq!(
            OptionalValue::<i32>::from_excel(decode(&empty)),
            Ok(OptionalValue::Empty)
        );
        assert_eq!(
            OptionalValue::<i32>::from_excel(decode(&integer)),
            Ok(OptionalValue::Value(3))
        );
    }
}
