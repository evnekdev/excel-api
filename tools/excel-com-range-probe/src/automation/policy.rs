//! Explicit conversion and Excel Range shape policies.

use super::{AutomationArray, ConversionError};

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub(crate) enum DateWritePolicy { DateVariant, Value2Serial }

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub(crate) enum ShapePolicy { Exact }

impl ShapePolicy {
    #[allow(dead_code)]
    pub(crate) fn validate(self, source: &AutomationArray, target_rows: usize, target_columns: usize) -> Result<(), ConversionError> {
        match self {
            Self::Exact if source.rows == target_rows && source.columns == target_columns => Ok(()),
            Self::Exact => Err(ConversionError::ShapeMismatch { source_rows: source.rows, source_columns: source.columns, target_rows, target_columns }),
        }
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub(crate) struct ConversionPolicy {
    pub(crate) date_write: DateWritePolicy,
    pub(crate) reject_non_finite_numbers: bool,
    pub(crate) reject_embedded_nul: bool,
    pub(crate) require_exact_integer_conversion: bool,
    pub(crate) shape: ShapePolicy,
}

impl Default for ConversionPolicy {
    fn default() -> Self {
        Self { date_write: DateWritePolicy::DateVariant, reject_non_finite_numbers: true, reject_embedded_nul: true, require_exact_integer_conversion: true, shape: ShapePolicy::Exact }
    }
}
