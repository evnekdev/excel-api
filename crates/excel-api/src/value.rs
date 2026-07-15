use core::slice;

use crate::{ExcelError, ExcelStr, OwnedValueError, Utf16ConversionError};

/// Owned UTF-16 payload independent of Excel callback memory.
///
/// The box contains payload code units only: no Excel length prefix and no
/// trailing NUL requirement. Arbitrary UTF-16 code units are preserved.
#[derive(Clone, Debug, Default, Eq, Hash, PartialEq)]
pub struct ExcelString {
    units: Box<[u16]>,
}

impl ExcelString {
    /// Owns the supplied UTF-16 code units without validating Unicode scalar
    /// values. Conversion limits apply when callback memory is copied, not to
    /// semantic strings constructed directly by Rust callers.
    pub fn from_utf16_units(units: impl Into<Box<[u16]>>) -> Self {
        Self {
            units: units.into(),
        }
    }

    /// Encodes valid UTF-8 as owned UTF-16 code units.
    pub fn from_utf8(value: &str) -> Self {
        Self::from_utf16_units(value.encode_utf16().collect::<Vec<_>>())
    }

    /// Returns the exact owned UTF-16 code units, including embedded NULs.
    pub fn as_utf16(&self) -> &[u16] {
        &self.units
    }

    /// Returns the number of UTF-16 code units, not Unicode scalar values.
    pub const fn len_utf16(&self) -> usize {
        self.units.len()
    }

    /// Reports whether the string has no UTF-16 payload units.
    pub const fn is_empty(&self) -> bool {
        self.units.is_empty()
    }

    /// Strictly decodes the UTF-16 payload.
    pub fn to_string(&self) -> Result<String, Utf16ConversionError> {
        String::from_utf16(&self.units).map_err(|_| Utf16ConversionError)
    }

    /// Explicitly decodes the payload, replacing unpaired surrogates with the
    /// Unicode replacement character.
    pub fn to_string_lossy(&self) -> String {
        String::from_utf16_lossy(&self.units)
    }
}

impl From<&str> for ExcelString {
    fn from(value: &str) -> Self {
        Self::from_utf8(value)
    }
}

impl From<String> for ExcelString {
    fn from(value: String) -> Self {
        Self::from_utf8(&value)
    }
}

impl From<ExcelStr<'_>> for ExcelString {
    fn from(value: ExcelStr<'_>) -> Self {
        Self::from_utf16_units(Box::<[u16]>::from(value.as_utf16()))
    }
}

impl TryFrom<&ExcelString> for String {
    type Error = Utf16ConversionError;

    fn try_from(value: &ExcelString) -> Result<Self, Self::Error> {
        value.to_string()
    }
}

/// Ordinary owned Rust representation of an Excel value.
///
/// Every variant is independent of callback storage. References deliberately
/// remain outside this semantic value model until an owned-reference contract
/// is specified.
#[derive(Clone, Debug, PartialEq)]
pub enum ExcelValue {
    /// An IEEE-754 worksheet number.
    Number(f64),
    /// An Excel integer represented as an `i32`.
    Integer(i32),
    /// An Excel Boolean.
    Boolean(bool),
    /// A controlled Excel worksheet error.
    Error(ExcelError),
    /// An omitted Excel argument.
    Missing,
    /// An empty Excel value.
    Empty,
    /// Owned UTF-16 text.
    Text(ExcelString),
    /// An owned rectangular array.
    Array(ExcelArray),
}

impl ExcelValue {
    /// Returns the stable diagnostic name of this semantic variant.
    pub const fn kind_name(&self) -> &'static str {
        match self {
            Self::Number(_) => "number",
            Self::Integer(_) => "integer",
            Self::Boolean(_) => "boolean",
            Self::Error(_) => "error",
            Self::Missing => "missing",
            Self::Empty => "empty",
            Self::Text(_) => "text",
            Self::Array(_) => "array",
        }
    }
}

/// Owned immutable rectangular array in row-major order.
#[derive(Clone, Debug, PartialEq)]
pub struct ExcelArray {
    rows: usize,
    columns: usize,
    values: Box<[ExcelValue]>,
}

impl ExcelArray {
    /// Creates a rectangular row-major array.
    ///
    /// Returns an error when `rows * columns` overflows, the element count does
    /// not match the shape, or an element is itself an array.
    pub fn new(
        rows: usize,
        columns: usize,
        values: impl Into<Box<[ExcelValue]>>,
    ) -> Result<Self, OwnedValueError> {
        let values = values.into();
        let expected = rows
            .checked_mul(columns)
            .ok_or(OwnedValueError::ArrayShapeOverflow { rows, columns })?;
        if values.len() != expected {
            return Err(OwnedValueError::InvalidArrayShape {
                rows,
                columns,
                elements: values.len(),
            });
        }
        if values
            .iter()
            .any(|value| matches!(value, ExcelValue::Array(_)))
        {
            return Err(OwnedValueError::NestedArrayUnsupported);
        }

        Ok(Self {
            rows,
            columns,
            values,
        })
    }

    /// Returns the number of rows.
    pub const fn rows(&self) -> usize {
        self.rows
    }

    /// Returns the number of columns.
    pub const fn columns(&self) -> usize {
        self.columns
    }

    /// Returns the total row-major element count.
    pub const fn len(&self) -> usize {
        self.values.len()
    }

    /// Reports whether the array contains no elements.
    pub const fn is_empty(&self) -> bool {
        self.values.is_empty()
    }

    /// Returns all values in row-major order.
    pub fn values(&self) -> &[ExcelValue] {
        &self.values
    }

    /// Returns the value at a zero-based row and column, if in bounds.
    pub fn get(&self, row: usize, column: usize) -> Option<&ExcelValue> {
        if row >= self.rows || column >= self.columns {
            return None;
        }
        self.values.get(row * self.columns + column)
    }

    /// Returns one zero-based row as a contiguous row-major slice.
    pub fn row(&self, row: usize) -> Option<&[ExcelValue]> {
        if row >= self.rows {
            return None;
        }
        let start = row * self.columns;
        self.values.get(start..start + self.columns)
    }

    /// Returns an iterator over one zero-based column, if in bounds.
    pub fn column(&self, column: usize) -> Option<ExcelArrayColumn<'_>> {
        (column < self.columns).then_some(ExcelArrayColumn {
            array: self,
            column,
            next_row: 0,
        })
    }

    /// Iterates all values in row-major order.
    pub fn iter(&self) -> slice::Iter<'_, ExcelValue> {
        self.values.iter()
    }

    pub(crate) fn into_parts(self) -> (usize, usize, Box<[ExcelValue]>) {
        (self.rows, self.columns, self.values)
    }
}

/// Checked iterator over one column of an owned rectangular array.
#[derive(Debug)]
pub struct ExcelArrayColumn<'array> {
    array: &'array ExcelArray,
    column: usize,
    next_row: usize,
}

impl<'array> Iterator for ExcelArrayColumn<'array> {
    type Item = &'array ExcelValue;

    fn next(&mut self) -> Option<Self::Item> {
        let row = self.next_row;
        let value = self.array.get(row, self.column)?;
        self.next_row += 1;
        Some(value)
    }

    fn size_hint(&self) -> (usize, Option<usize>) {
        let remaining = self.array.rows - self.next_row;
        (remaining, Some(remaining))
    }
}

impl ExactSizeIterator for ExcelArrayColumn<'_> {}

/// Exact optional-value semantics for callers that must distinguish omitted
/// arguments from blank cells.
#[derive(Clone, Debug, PartialEq)]
pub enum OptionalValue<T> {
    /// An omitted argument.
    Missing,
    /// An empty Excel value.
    Empty,
    /// A present converted value.
    Value(T),
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn excel_string_preserves_every_required_utf16_form() {
        let cases: &[&[u16]] = &[
            &[],
            &[b'A' as u16],
            &[0x00E9, 0x6C34],
            &[0xD83D, 0xDE03],
            &[0xD800],
            &[0xDC00],
            &[b'A' as u16, 0, b'B' as u16],
        ];
        for units in cases {
            let value = ExcelString::from_utf16_units(units.to_vec());
            assert_eq!(value.as_utf16(), *units);
            assert_eq!(value.len_utf16(), units.len());
            assert_eq!(value.is_empty(), units.is_empty());
        }
    }

    #[test]
    fn strict_and_lossy_utf16_conversion_are_distinct() {
        let valid = ExcelString::from_utf16_units([0xD83D, 0xDE03].to_vec());
        assert_eq!(valid.to_string().as_deref(), Ok("😃"));

        let invalid = ExcelString::from_utf16_units([0xD800].to_vec());
        assert_eq!(invalid.to_string(), Err(Utf16ConversionError));
        assert_eq!(invalid.to_string_lossy(), "�");

        let invalid = ExcelString::from_utf16_units([0xDC00].to_vec());
        assert_eq!(invalid.to_string(), Err(Utf16ConversionError));
        assert_eq!(invalid.to_string_lossy(), "�");
    }

    #[test]
    fn utf8_conversion_is_infallible_and_exact() {
        let value = ExcelString::from("Aé水😃");
        assert_eq!(value.to_string().as_deref(), Ok("Aé水😃"));

        let value = ExcelString::from(String::from("owned 😃"));
        assert_eq!(String::try_from(&value).as_deref(), Ok("owned 😃"));
    }

    #[test]
    fn array_shape_and_nested_array_are_rejected() {
        assert!(matches!(
            ExcelArray::new(2, 2, vec![ExcelValue::Number(1.0)]),
            Err(OwnedValueError::InvalidArrayShape { .. })
        ));
        assert!(matches!(
            ExcelArray::new(usize::MAX, 2, Vec::<ExcelValue>::new()),
            Err(OwnedValueError::ArrayShapeOverflow { .. })
        ));

        let inner = ExcelArray::new(0, 0, Vec::<ExcelValue>::new()).unwrap();
        assert!(inner.is_empty());
        assert_eq!(
            ExcelArray::new(1, 1, vec![ExcelValue::Array(inner)]),
            Err(OwnedValueError::NestedArrayUnsupported)
        );
    }

    #[test]
    fn array_indexing_rows_columns_and_iteration_are_row_major() {
        let array =
            ExcelArray::new(2, 3, (0..6).map(ExcelValue::Integer).collect::<Vec<_>>()).unwrap();

        assert_eq!(array.len(), 6);
        assert!(!array.is_empty());
        assert_eq!(array.get(1, 1), Some(&ExcelValue::Integer(4)));
        assert_eq!(array.get(2, 0), None);
        assert_eq!(
            array.row(1),
            Some(
                [
                    ExcelValue::Integer(3),
                    ExcelValue::Integer(4),
                    ExcelValue::Integer(5),
                ]
                .as_slice()
            )
        );
        assert_eq!(
            array.column(1).unwrap().collect::<Vec<_>>(),
            vec![&ExcelValue::Integer(1), &ExcelValue::Integer(4)]
        );
        assert_eq!(array.iter().count(), 6);
    }

    #[test]
    fn owned_types_are_static_send_and_sync() {
        fn assert_owned<T: Send + Sync + 'static>() {}
        assert_owned::<ExcelString>();
        assert_owned::<ExcelValue>();
        assert_owned::<ExcelArray>();

        let value = ExcelValue::Text(ExcelString::from("worker"));
        let returned = std::thread::spawn(move || value).join().unwrap();
        assert_eq!(
            returned,
            ExcelValue::Text(ExcelString::from_utf16_units(
                "worker".encode_utf16().collect::<Vec<_>>()
            ))
        );
    }
}
