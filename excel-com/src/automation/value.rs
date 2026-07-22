use super::{
    AutomationArray, ConversionError, ConversionPolicy, Currency, ExcelError, OaDate, OwnedVariant,
    SafeArray,
};
use crate::ExcelComError;
use windows_sys::Win32::System::Com::SAFEARRAYBOUND;
use windows_sys::Win32::System::Variant::{
    VT_ARRAY, VT_BOOL, VT_BSTR, VT_CY, VT_DATE, VT_EMPTY, VT_ERROR, VT_I4, VT_NULL, VT_R8,
    VT_VARIANT,
};
/// Pointer-free Automation value used for Range values and formula text.
///
/// Arrays are zero-based row-major [`AutomationArray`] values. Date values
/// retain an OLE Automation serial, and errors retain the exact signed SCODE.
#[derive(Clone, Debug, PartialEq)]
pub enum AutomationValue {
    /// `VT_EMPTY`.
    Empty,
    /// `VT_NULL`.
    Null,
    /// `VT_BOOL`.
    Bool(bool),
    /// A finite numeric value, encoded as `VT_R8`.
    Number(f64),
    /// UTF-16 BSTR text.
    Text(String),
    /// A `VT_ERROR` value retaining its SCODE.
    Error(ExcelError),
    /// An OLE Automation `VT_DATE` serial.
    Date(OaDate),
    /// Exact `VT_CY` currency.
    Currency(Currency),
    /// A rank-two `VT_ARRAY|VT_VARIANT` rectangle.
    Array(AutomationArray),
}

/// Selects the Runtime representation of semantic dates for a Range setter.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub(crate) enum DateWriteMode {
    Value,
    Value2,
}

pub(crate) fn validate_range_shape(
    value: &AutomationValue,
    target_rows: usize,
    target_columns: usize,
) -> Result<(), ExcelComError> {
    match value {
        AutomationValue::Array(source)
            if source.rows() == target_rows && source.columns() == target_columns =>
        {
            Ok(())
        }
        AutomationValue::Array(source) => {
            Err(ExcelComError::Conversion(ConversionError::ShapeMismatch {
                source_rows: source.rows(),
                source_columns: source.columns(),
                target_rows,
                target_columns,
            }))
        }
        _ if target_rows == 1 && target_columns == 1 => Ok(()),
        _ => Err(ExcelComError::Conversion(ConversionError::ShapeMismatch {
            source_rows: 1,
            source_columns: 1,
            target_rows,
            target_columns,
        })),
    }
}

pub(crate) fn encode_variant(
    value: &AutomationValue,
    policy: ConversionPolicy,
) -> Result<OwnedVariant, ExcelComError> {
    match value {
        AutomationValue::Empty => Ok(OwnedVariant::empty()),
        AutomationValue::Null => Ok(OwnedVariant::null()),
        AutomationValue::Bool(value) => Ok(OwnedVariant::bool(*value)),
        AutomationValue::Number(value) if value.is_finite() => Ok(OwnedVariant::f64(*value)),
        AutomationValue::Number(_) => {
            Err(ExcelComError::Conversion(ConversionError::NonFiniteNumber))
        }
        AutomationValue::Text(value) => OwnedVariant::bstr(value),
        AutomationValue::Error(value) => Ok(OwnedVariant::error(value.scode())),
        AutomationValue::Date(value) => match policy.date_write {
            DateWriteMode::Value => Ok(OwnedVariant::date(value.serial())),
            DateWriteMode::Value2 => Ok(OwnedVariant::f64(value.serial())),
        },
        AutomationValue::Currency(value) => Ok(OwnedVariant::currency(value.scaled())),
        AutomationValue::Array(value) => encode_array(value, policy),
    }
}

fn encode_array(
    value: &AutomationArray,
    policy: ConversionPolicy,
) -> Result<OwnedVariant, ExcelComError> {
    let rows = u32::try_from(value.rows())
        .map_err(|_| ExcelComError::Conversion(ConversionError::InvalidElementCount))?;
    let columns = u32::try_from(value.columns())
        .map_err(|_| ExcelComError::Conversion(ConversionError::InvalidElementCount))?;
    if rows == 0 || columns == 0 {
        return Err(ExcelComError::Conversion(
            ConversionError::InvalidElementCount,
        ));
    }
    let array = SafeArray::create_variant(&[
        SAFEARRAYBOUND {
            cElements: rows,
            lLbound: 1,
        },
        SAFEARRAYBOUND {
            cElements: columns,
            lLbound: 1,
        },
    ])
    .ok_or(ExcelComError::Conversion(
        ConversionError::SafeArrayConstructionFailed,
    ))?;
    for row in 0..value.rows() {
        for column in 0..value.columns() {
            let item = value.get(row, column).ok_or(ExcelComError::Conversion(
                ConversionError::InvalidElementCount,
            ))?;
            if matches!(item, AutomationValue::Array(_)) {
                return Err(ExcelComError::Conversion(
                    ConversionError::UnsupportedVariantType {
                        vartype: VT_ARRAY | VT_VARIANT,
                    },
                ));
            }
            let encoded = encode_variant(item, policy).map_err(|_| {
                ExcelComError::Conversion(ConversionError::SafeArrayElementFailed { row, column })
            })?;
            let indices = [
                i32::try_from(row + 1)
                    .map_err(|_| ExcelComError::Conversion(ConversionError::InvalidElementCount))?,
                i32::try_from(column + 1)
                    .map_err(|_| ExcelComError::Conversion(ConversionError::InvalidElementCount))?,
            ];
            if !array.put_variant(&indices, &encoded) {
                return Err(ExcelComError::Conversion(
                    ConversionError::SafeArrayElementFailed { row, column },
                ));
            }
        }
    }
    Ok(OwnedVariant::array(array))
}

pub(crate) fn decode_variant(
    value: &OwnedVariant,
    policy: ConversionPolicy,
) -> Result<AutomationValue, ExcelComError> {
    if value.vt() == VT_ARRAY | VT_VARIANT {
        return decode_array(value, policy).map(AutomationValue::Array);
    }
    // SAFETY: each match arm reads only the union field selected by the checked VARTYPE.
    let result = unsafe {
        match value.vt() {
            VT_EMPTY => Ok(AutomationValue::Empty),
            VT_NULL => Ok(AutomationValue::Null),
            VT_BOOL => Ok(AutomationValue::Bool(
                value.0.Anonymous.Anonymous.Anonymous.boolVal != 0,
            )),
            VT_I4 => Ok(AutomationValue::Number(
                value.0.Anonymous.Anonymous.Anonymous.lVal as f64,
            )),
            VT_R8 => finite_number(value.0.Anonymous.Anonymous.Anonymous.dblVal),
            VT_BSTR => value.as_string().map(AutomationValue::Text),
            VT_ERROR => Ok(AutomationValue::Error(ExcelError::from_scode(
                value.0.Anonymous.Anonymous.Anonymous.scode,
            ))),
            VT_DATE => {
                OaDate::new(value.0.Anonymous.Anonymous.Anonymous.date).map(AutomationValue::Date)
            }
            VT_CY => Ok(AutomationValue::Currency(Currency::from_scaled(
                value.0.Anonymous.Anonymous.Anonymous.cyVal.int64,
            ))),
            vartype => Err(ExcelComError::Conversion(
                ConversionError::UnsupportedVariantType { vartype },
            )),
        }
    };
    let _ = policy;
    result
}

fn finite_number(value: f64) -> Result<AutomationValue, ExcelComError> {
    value
        .is_finite()
        .then_some(AutomationValue::Number(value))
        .ok_or(ExcelComError::Conversion(ConversionError::NonFiniteNumber))
}

fn decode_array(
    value: &OwnedVariant,
    policy: ConversionPolicy,
) -> Result<AutomationArray, ExcelComError> {
    // SAFETY: the caller established the VT_ARRAY|VT_VARIANT tag before reading parray.
    let raw = unsafe { value.0.Anonymous.Anonymous.Anonymous.parray };
    let metadata = SafeArray::metadata(raw).ok_or(ExcelComError::Conversion(
        ConversionError::SafeArrayConstructionFailed,
    ))?;
    if metadata.rank != 2 {
        return Err(ExcelComError::Conversion(
            ConversionError::UnsupportedSafeArrayRank {
                rank: metadata.rank,
            },
        ));
    }
    if metadata.element_vartype != Some(VT_VARIANT) {
        return Err(ExcelComError::Conversion(
            ConversionError::UnsupportedSafeArrayElementType {
                vartype: metadata.element_vartype.unwrap_or_default(),
            },
        ));
    }
    let [row_dimension, column_dimension]: &[_; 2] =
        metadata.dimensions.as_slice().try_into().map_err(|_| {
            ExcelComError::Conversion(ConversionError::UnsupportedSafeArrayRank {
                rank: metadata.rank,
            })
        })?;
    let rows = usize::try_from(row_dimension.element_count)
        .map_err(|_| ExcelComError::Conversion(ConversionError::InvalidElementCount))?;
    let columns = usize::try_from(column_dimension.element_count)
        .map_err(|_| ExcelComError::Conversion(ConversionError::InvalidElementCount))?;
    let mut values = Vec::with_capacity(rows.checked_mul(columns).ok_or(
        ExcelComError::Conversion(ConversionError::InvalidElementCount),
    )?);
    for row in 0..rows {
        for column in 0..columns {
            let indices = [
                array_index(row_dimension.lower_bound, row)?,
                array_index(column_dimension.lower_bound, column)?,
            ];
            let item = SafeArray::get_variant_borrowed(raw, &indices).ok_or(
                ExcelComError::Conversion(ConversionError::SafeArrayElementFailed { row, column }),
            )?;
            values.push(decode_variant(&item, policy).map_err(|_| {
                ExcelComError::Conversion(ConversionError::SafeArrayElementFailed { row, column })
            })?);
        }
    }
    AutomationArray::new(rows, columns, values)
}

fn array_index(lower_bound: i32, offset: usize) -> Result<i32, ExcelComError> {
    lower_bound
        .checked_add(
            i32::try_from(offset)
                .map_err(|_| ExcelComError::Conversion(ConversionError::InvalidElementCount))?,
        )
        .ok_or(ExcelComError::Conversion(
            ConversionError::InvalidElementCount,
        ))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn scalar_to_multi_cell_is_rejected_before_com() {
        assert!(matches!(
            validate_range_shape(&AutomationValue::Number(1.0), 2, 2),
            Err(ExcelComError::Conversion(
                ConversionError::ShapeMismatch { .. }
            ))
        ));
    }

    #[test]
    fn exact_array_shape_is_required() {
        let array = AutomationArray::new(2, 3, vec![AutomationValue::Empty; 6]).expect("shape");
        assert!(validate_range_shape(&AutomationValue::Array(array.clone()), 2, 3).is_ok());
        assert!(matches!(
            validate_range_shape(&AutomationValue::Array(array), 3, 2),
            Err(ExcelComError::Conversion(
                ConversionError::ShapeMismatch { .. }
            ))
        ));
    }
}
