//! Invocation arguments distinct from worksheet values.

use windows_sys::Win32::Foundation::DISP_E_PARAMNOTFOUND;

use super::{encode_variant, AutomationValue, ConversionError, ConversionPolicy};
use crate::raw::variant::OwnedVariant;

#[allow(dead_code)]
#[derive(Clone, Debug, PartialEq)]
pub(crate) enum AutomationArgument {
    Value(AutomationValue),
    Missing,
}

#[allow(dead_code)]
pub(crate) fn encode_argument(
    argument: &AutomationArgument,
    policy: &ConversionPolicy,
) -> Result<OwnedVariant, ConversionError> {
    match argument {
        AutomationArgument::Value(value) => encode_variant(value, policy),
        AutomationArgument::Missing => Ok(OwnedVariant::error(DISP_E_PARAMNOTFOUND)),
    }
}
