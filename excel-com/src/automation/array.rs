use super::AutomationValue;
use crate::{ConversionError, ExcelComError};

/// A zero-based, row-major rectangular Automation array.
#[derive(Clone, Debug, PartialEq)]
pub struct AutomationArray {
    rows: usize,
    columns: usize,
    values: Vec<AutomationValue>,
}
impl AutomationArray {
    /// Creates a row-major rectangle with the supplied dimensions and values.
    pub fn new(
        rows: usize,
        columns: usize,
        values: Vec<AutomationValue>,
    ) -> Result<Self, ExcelComError> {
        (rows.checked_mul(columns) == Some(values.len()))
            .then_some(Self {
                rows,
                columns,
                values,
            })
            .ok_or(ExcelComError::Conversion(
                ConversionError::InvalidElementCount,
            ))
    }
    /// Creates a one-row Automation array.
    pub fn row(values: Vec<AutomationValue>) -> Result<Self, ExcelComError> {
        Self::new(1, values.len(), values)
    }
    /// Creates a one-column Automation array.
    pub fn column(values: Vec<AutomationValue>) -> Result<Self, ExcelComError> {
        Self::new(values.len(), 1, values)
    }
    /// Creates a rectangular Automation array from rows of equal length.
    pub fn from_rows(rows: Vec<Vec<AutomationValue>>) -> Result<Self, ExcelComError> {
        let columns = rows.first().map_or(0, Vec::len);
        if rows.iter().any(|row| row.len() != columns) {
            return Err(ExcelComError::Conversion(
                ConversionError::InvalidElementCount,
            ));
        }
        let row_count = rows.len();
        let values = rows.into_iter().flatten().collect();
        Self::new(row_count, columns, values)
    }
    /// Returns the number of rows.
    pub const fn rows(&self) -> usize {
        self.rows
    }
    /// Returns the number of columns.
    pub const fn columns(&self) -> usize {
        self.columns
    }
    /// Returns an element by zero-based row and column, if it is in bounds.
    pub fn get(&self, row: usize, column: usize) -> Option<&AutomationValue> {
        (row < self.rows && column < self.columns)
            .then(|| &self.values[row * self.columns + column])
    }
    /// Returns the values in zero-based, row-major order.
    pub fn values(&self) -> &[AutomationValue] {
        &self.values
    }
}
