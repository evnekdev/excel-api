use crate::ExcelError;

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
