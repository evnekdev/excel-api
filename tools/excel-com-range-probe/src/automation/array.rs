//! Exact rectangular Automation-array shape and row-major values.

use super::{AutomationValue, ConversionError};

#[derive(Clone, Debug, PartialEq)]
pub(crate) struct AutomationArray {
    pub(super) rows: usize,
    pub(super) columns: usize,
    pub(super) values: Vec<AutomationValue>,
}

impl AutomationArray {
    pub(crate) fn new(rows: usize, columns: usize, values: Vec<AutomationValue>) -> Result<Self, ConversionError> {
        (rows.checked_mul(columns) == Some(values.len()))
            .then_some(Self { rows, columns, values })
            .ok_or(ConversionError::InvalidElementCount)
    }

    pub(crate) fn row(values: Vec<AutomationValue>) -> Result<Self, ConversionError> { Self::new(1, values.len(), values) }
    pub(crate) fn column(values: Vec<AutomationValue>) -> Result<Self, ConversionError> { Self::new(values.len(), 1, values) }

    pub(crate) fn from_rows(rows: Vec<Vec<AutomationValue>>) -> Result<Self, ConversionError> {
        let row_count = rows.len();
        let columns = rows.first().map_or(0, Vec::len);
        if rows.iter().any(|row| row.len() != columns) { return Err(ConversionError::InvalidElementCount); }
        Self::new(row_count, columns, rows.into_iter().flatten().collect())
    }

    #[allow(dead_code)]
    pub(crate) const fn rows(&self) -> usize { self.rows }
    #[allow(dead_code)]
    pub(crate) const fn columns(&self) -> usize { self.columns }
    #[allow(dead_code)]
    pub(crate) const fn len(&self) -> usize { self.values.len() }
    #[allow(dead_code)]
    pub(crate) const fn is_empty(&self) -> bool { self.values.is_empty() }
    pub(crate) fn get(&self, row: usize, column: usize) -> Option<&AutomationValue> { self.index(row, column).and_then(|index| self.values.get(index)) }
    #[allow(dead_code)]
    pub(crate) fn get_mut(&mut self, row: usize, column: usize) -> Option<&mut AutomationValue> { self.index(row, column).and_then(|index| self.values.get_mut(index)) }
    pub(crate) fn values(&self) -> &[AutomationValue] { &self.values }
    #[allow(dead_code)]
    pub(crate) fn into_values(self) -> Vec<AutomationValue> { self.values }

    fn index(&self, row: usize, column: usize) -> Option<usize> {
        (row < self.rows && column < self.columns).then(|| row.checked_mul(self.columns)?.checked_add(column)).flatten()
    }
}
