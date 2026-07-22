use super::{AutomationValue, ConversionPolicy, OwnedVariant, encode_variant};
use crate::ExcelComError;
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
