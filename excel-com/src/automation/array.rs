use super::AutomationValue;
use crate::ExcelComError;

/// A zero-based, row-major rectangular Automation array.
#[derive(Clone, Debug, PartialEq)]
pub struct AutomationArray {
    rows: usize,
    columns: usize,
    values: Vec<AutomationValue>,
}
impl AutomationArray {
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
            .ok_or(ExcelComError::Conversion {
                detail: "array shape does not match element count",
            })
    }
    pub fn row(values: Vec<AutomationValue>) -> Result<Self, ExcelComError> {
        Self::new(1, values.len(), values)
    }
    pub fn column(values: Vec<AutomationValue>) -> Result<Self, ExcelComError> {
        Self::new(values.len(), 1, values)
    }
    pub const fn rows(&self) -> usize {
        self.rows
    }
    pub const fn columns(&self) -> usize {
        self.columns
    }
    pub fn get(&self, row: usize, column: usize) -> Option<&AutomationValue> {
        (row < self.rows && column < self.columns)
            .then(|| &self.values[row * self.columns + column])
    }
    pub fn values(&self) -> &[AutomationValue] {
        &self.values
    }
}
