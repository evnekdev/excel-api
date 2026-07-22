use super::{AutomationValue, ConversionError, ConversionPolicy, OwnedVariant, encode_variant};
use crate::{ExcelComError, excel::DispatchObject};
use windows_sys::Win32::Foundation::DISP_E_PARAMNOTFOUND;

/// Invocation-only distinction between a supplied Automation value and Missing.
#[derive(Clone, Debug, PartialEq)]
pub enum AutomationArgument {
    /// A supplied semantic Automation value.
    Value(AutomationValue),
    /// An explicit `VT_ERROR` / `DISP_E_PARAMNOTFOUND` optional argument.
    Missing,
}

impl AutomationArgument {
    pub(crate) fn encode(&self, policy: ConversionPolicy) -> Result<OwnedVariant, ExcelComError> {
        match self {
            Self::Value(value) => encode_variant(value, policy),
            Self::Missing => Ok(OwnedVariant::error(DISP_E_PARAMNOTFOUND)),
        }
    }
}

/// Logical-order argument builder for positional Automation calls.
///
/// The builder never removes trailing missing values.  The dispatch layer is
/// solely responsible for reversing the returned list into COM order.
pub(crate) struct PositionalArguments {
    values: Vec<OwnedVariant>,
}

impl PositionalArguments {
    pub(crate) const fn new() -> Self {
        Self { values: Vec::new() }
    }

    pub(crate) fn push_required(&mut self, value: OwnedVariant) {
        self.values.push(value);
    }

    pub(crate) fn push_optional(&mut self, value: Option<OwnedVariant>) {
        self.values
            .push(value.unwrap_or_else(|| OwnedVariant::error(DISP_E_PARAMNOTFOUND)));
    }

    pub(crate) fn push_argument(
        &mut self,
        value: AutomationArgument,
        policy: ConversionPolicy,
    ) -> Result<(), ExcelComError> {
        self.push_result(value.encode(policy))
    }

    pub(crate) fn push_result(
        &mut self,
        value: Result<OwnedVariant, ExcelComError>,
    ) -> Result<(), ExcelComError> {
        let position = self.values.len();
        self.values.push(value.map_err(|error| match error {
            ExcelComError::Conversion(source) => {
                ExcelComError::Conversion(ConversionError::ArgumentConversion {
                    position,
                    source: Box::new(source),
                })
            }
            other => other,
        })?);
        Ok(())
    }

    /// Safely passes a wrapper object as a `VT_DISPATCH` argument.
    pub(crate) fn push_object(&mut self, value: &DispatchObject) {
        self.values
            .push(OwnedVariant::dispatch_borrowed(&value.dispatch));
    }

    /// Safely passes an optional wrapper object, preserving `Missing` when absent.
    pub(crate) fn push_optional_object(&mut self, value: Option<&DispatchObject>) {
        match value {
            Some(value) => self.push_object(value),
            None => self.push_optional(None),
        }
    }

    pub(crate) fn into_inner(self) -> Vec<OwnedVariant> {
        self.values
    }
}

pub(crate) fn reverse_for_com(values: &mut [OwnedVariant]) {
    values.reverse();
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn missing_arguments_keep_interior_and_trailing_positions() {
        let mut arguments = PositionalArguments::new();
        arguments.push_optional(None);
        arguments.push_required(OwnedVariant::i32(7));
        arguments.push_optional(None);
        let mut values = arguments.into_inner();
        assert_eq!(values.len(), 3);
        assert_eq!(values[0].as_scode(), Some(DISP_E_PARAMNOTFOUND));
        assert_eq!(values[1].as_i32(), Some(7));
        assert_eq!(values[2].as_scode(), Some(DISP_E_PARAMNOTFOUND));
        reverse_for_com(&mut values);
        assert_eq!(values[0].as_scode(), Some(DISP_E_PARAMNOTFOUND));
        assert_eq!(values[1].as_i32(), Some(7));
        assert_eq!(values[2].as_scode(), Some(DISP_E_PARAMNOTFOUND));
    }

    #[test]
    fn conversion_error_reports_logical_argument_position() {
        let mut arguments = PositionalArguments::new();
        arguments.push_optional(None);
        let error = arguments
            .push_argument(
                AutomationArgument::Value(AutomationValue::Number(f64::NAN)),
                ConversionPolicy::default(),
            )
            .expect_err("NaN must not cross the COM boundary");
        assert!(matches!(
            error,
            ExcelComError::Conversion(ConversionError::ArgumentConversion { position: 1, .. })
        ));
    }
}
