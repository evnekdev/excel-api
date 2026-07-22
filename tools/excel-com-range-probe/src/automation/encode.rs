//! Excel Range-write validation separate from generic BSTR encoding.

use super::{AutomationValue, ConversionError};

pub(crate) const EXCEL_CELL_STRING_LIMIT: usize = 32_767;

pub(crate) fn validate_excel_range_write(value: &AutomationValue) -> Result<(), ConversionError> {
    match value {
        AutomationValue::Text(value) if value.chars().count() > EXCEL_CELL_STRING_LIMIT => {
            Err(ConversionError::StringTooLong)
        }
        AutomationValue::Array(array) => array
            .values()
            .iter()
            .try_for_each(validate_excel_range_write),
        _ => Ok(()),
    }
}
