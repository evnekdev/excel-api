use std::path::Path;

use windows_sys::Win32::System::Com::SAFEARRAYBOUND;

use crate::automation::{AutomationArray, AutomationValue, OwnedVariant, SafeArray};
use crate::{ExcelComError, internal::path_bstr};

use super::{TextColumnSpec, TextDelimiter, TextParsingType};

type DelimiterArguments = (
    Option<bool>,
    Option<bool>,
    Option<bool>,
    Option<bool>,
    Option<bool>,
    Option<OwnedVariant>,
);

pub(super) fn text_path(path: &Path) -> Result<OwnedVariant, ExcelComError> {
    path_bstr(path)
}

pub(super) fn text_argument(value: &str) -> Result<OwnedVariant, ExcelComError> {
    if value.contains('\0') {
        return Err(ExcelComError::Unsupported {
            detail: "Excel Automation text cannot contain embedded NUL",
        });
    }
    OwnedVariant::bstr(value)
}

pub(super) fn one_based(value: usize, detail: &'static str) -> Result<i32, ExcelComError> {
    if value == 0 {
        return Err(ExcelComError::Unsupported { detail });
    }
    i32::try_from(value).map_err(|_| ExcelComError::Unsupported { detail })
}

pub(super) fn finite(
    value: Option<f64>,
    detail: &'static str,
) -> Result<Option<OwnedVariant>, ExcelComError> {
    match value {
        Some(value) if value.is_finite() => Ok(Some(OwnedVariant::f64(value))),
        Some(_) => Err(ExcelComError::Unsupported { detail }),
        None => Ok(None),
    }
}

pub(super) fn delimiter_arguments(
    delimiter: Option<&TextDelimiter>,
) -> Result<DelimiterArguments, ExcelComError> {
    let (tab, semicolon, comma, space, other_enabled, other) = match delimiter {
        None => (None, None, None, None, None, None),
        Some(TextDelimiter::Tab) => (
            Some(true),
            Some(false),
            Some(false),
            Some(false),
            Some(false),
            None,
        ),
        Some(TextDelimiter::Semicolon) => (
            Some(false),
            Some(true),
            Some(false),
            Some(false),
            Some(false),
            None,
        ),
        Some(TextDelimiter::Comma) => (
            Some(false),
            Some(false),
            Some(true),
            Some(false),
            Some(false),
            None,
        ),
        Some(TextDelimiter::Space) => (
            Some(false),
            Some(false),
            Some(false),
            Some(true),
            Some(false),
            None,
        ),
        Some(TextDelimiter::Other(value)) => (
            Some(false),
            Some(false),
            Some(false),
            Some(false),
            Some(true),
            Some(*value),
        ),
        Some(TextDelimiter::Custom {
            tab,
            semicolon,
            comma,
            space,
            other,
        }) => (
            Some(*tab),
            Some(*semicolon),
            Some(*comma),
            Some(*space),
            Some(other.is_some()),
            *other,
        ),
    };
    let other = match other {
        Some('\0') => {
            return Err(ExcelComError::Unsupported {
                detail: "TextDelimiter::Other cannot be NUL",
            });
        }
        Some(value) => Some(OwnedVariant::bstr(&value.to_string())?),
        None => None,
    };
    Ok((tab, semicolon, comma, space, other_enabled, other))
}

/// Encodes Excel's rank-two `FieldInfo` SAFEARRAY: `[start, column type]` rows.
pub(super) fn field_info(
    columns: &[TextColumnSpec],
    parsing: TextParsingType,
) -> Result<Option<OwnedVariant>, ExcelComError> {
    if columns.is_empty() {
        return Ok(None);
    }
    let mut values = Vec::with_capacity(columns.len() * 2);
    let mut previous: Option<usize> = None;
    for column in columns {
        let start = column
            .start
            .unwrap_or_else(|| previous.map_or(1, |value| value + 1));
        if parsing != TextParsingType::FIXED_WIDTH && start == 0 {
            return Err(ExcelComError::Unsupported {
                detail: "delimited FieldInfo columns are one-based and nonzero",
            });
        }
        if parsing == TextParsingType::FIXED_WIDTH && previous.is_some_and(|value| start <= value) {
            return Err(ExcelComError::Unsupported {
                detail: "fixed-width FieldInfo starts must strictly increase",
            });
        }
        previous = Some(start);
        let start = i32::try_from(start).map_err(|_| ExcelComError::Unsupported {
            detail: "FieldInfo start exceeds i32",
        })?;
        values.push(AutomationValue::Number(f64::from(start)));
        values.push(AutomationValue::Number(f64::from(column.column_type.raw())));
    }
    let array = AutomationArray::new(columns.len(), 2, values)?;
    crate::automation::encode_variant(
        &AutomationValue::Array(array),
        crate::ConversionPolicy::default(),
    )
    .map(Some)
}

pub(super) fn one_dimensional_i32(
    values: &[usize],
    detail: &'static str,
) -> Result<OwnedVariant, ExcelComError> {
    if values.is_empty() {
        return Err(ExcelComError::Unsupported { detail });
    }
    let count = u32::try_from(values.len()).map_err(|_| ExcelComError::Unsupported { detail })?;
    let array = SafeArray::create_variant(&[SAFEARRAYBOUND {
        cElements: count,
        lLbound: 1,
    }])
    .ok_or(ExcelComError::Unsupported {
        detail: "could not allocate Automation SAFEARRAY",
    })?;
    for (offset, value) in values.iter().copied().enumerate() {
        let encoded = OwnedVariant::i32(one_based(value, detail)?);
        let index = i32::try_from(offset + 1).map_err(|_| ExcelComError::Unsupported { detail })?;
        if !array.put_variant(&[index], &encoded) {
            return Err(ExcelComError::Unsupported {
                detail: "could not populate Automation SAFEARRAY",
            });
        }
    }
    Ok(OwnedVariant::array(array))
}

pub(super) fn array_argument(array: &AutomationArray) -> Result<OwnedVariant, ExcelComError> {
    crate::automation::encode_variant(
        &AutomationValue::Array(array.clone()),
        crate::ConversionPolicy::default(),
    )
}
