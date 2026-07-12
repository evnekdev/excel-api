use crate::ExcelError;

/// Borrowed value received from Excel.
///
/// Borrowed string and array variants will be introduced after the raw ABI has
/// been validated against the Excel SDK.
#[derive(Clone, Copy, Debug, PartialEq)]
pub enum ExcelValueRef<'a> {
    Number(f64),
    Boolean(bool),
    Error(ExcelError),
    Missing,
    Empty,
    Text(&'a str),
}

/// Ordinary owned Rust representation of an Excel value.
#[derive(Clone, Debug, PartialEq)]
pub enum ExcelValue {
    Number(f64),
    Boolean(bool),
    Error(ExcelError),
    Missing,
    Empty,
    Text(String),
    Array(ExcelArray),
}

/// Owned rectangular array.
#[derive(Clone, Debug, PartialEq)]
pub struct ExcelArray {
    rows: usize,
    columns: usize,
    values: Vec<ExcelValue>,
}

impl ExcelArray {
    pub fn new(
        rows: usize,
        columns: usize,
        values: Vec<ExcelValue>,
    ) -> Result<Self, ArrayShapeError> {
        let expected = rows.checked_mul(columns).ok_or(ArrayShapeError)?;
        if values.len() != expected {
            return Err(ArrayShapeError);
        }

        Ok(Self {
            rows,
            columns,
            values,
        })
    }

    pub const fn rows(&self) -> usize {
        self.rows
    }

    pub const fn columns(&self) -> usize {
        self.columns
    }

    pub fn values(&self) -> &[ExcelValue] {
        &self.values
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct ArrayShapeError;

/// Exact optional-value semantics for callers that must distinguish omitted
/// arguments from blank cells.
#[derive(Clone, Debug, PartialEq)]
pub enum OptionalValue<T> {
    Missing,
    Empty,
    Value(T),
}

impl ExcelValueRef<'_> {
    pub const fn kind_name(self) -> &'static str {
        match self {
            Self::Number(_) => "number",
            Self::Boolean(_) => "boolean",
            Self::Error(_) => "error",
            Self::Missing => "missing",
            Self::Empty => "empty",
            Self::Text(_) => "text",
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn rejects_incorrect_array_shape() {
        assert_eq!(
            ExcelArray::new(2, 2, vec![ExcelValue::Number(1.0)]),
            Err(ArrayShapeError)
        );
    }
}
