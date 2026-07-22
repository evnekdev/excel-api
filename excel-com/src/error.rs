use crate::automation::ConversionError;
use std::error::Error;
use std::fmt::{Display, Formatter};

/// Production-facing Automation failure without raw address disclosure.
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum ExcelComError {
    /// COM apartment initialization failed.
    Initialization {
        /// The failed HRESULT.
        hresult: i32,
    },
    /// Creating the local Excel Automation server failed.
    Activation {
        /// The failed HRESULT.
        hresult: i32,
    },
    /// Looking up an Automation member name failed.
    NameLookup {
        /// The requested member name.
        member: &'static str,
        /// The failed HRESULT.
        hresult: i32,
    },
    /// Calling an Automation member failed.
    Invocation {
        /// The invoked member name.
        member: &'static str,
        /// The resolved member DISPID.
        dispid: i32,
        /// The failed HRESULT.
        hresult: i32,
        /// The server-provided exception SCODE, if any.
        exception_scode: Option<i32>,
        /// The Automation argument index reported by COM, if any.
        argument_index: Option<u32>,
    },
    /// A value could not be converted before or after an Automation call.
    Conversion(ConversionError),
    /// A COM ownership invariant was violated.
    Ownership {
        /// A static description of the invariant.
        detail: &'static str,
    },
    /// The requested operation is intentionally outside this crate's slice.
    Unsupported {
        /// A static description of the unsupported operation.
        detail: &'static str,
    },
}

impl ExcelComError {
    pub(crate) const fn failed(hresult: i32) -> bool {
        hresult < 0
    }
}

impl Display for ExcelComError {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Initialization { hresult } => write!(
                formatter,
                "COM initialization failed (0x{:08X})",
                *hresult as u32
            ),
            Self::Activation { hresult } => write!(
                formatter,
                "Excel activation failed (0x{:08X})",
                *hresult as u32
            ),
            Self::NameLookup { member, hresult } => write!(
                formatter,
                "name lookup for {member} failed (0x{:08X})",
                *hresult as u32
            ),
            Self::Invocation {
                member,
                dispid,
                hresult,
                exception_scode,
                argument_index,
            } => write!(
                formatter,
                "invocation of {member} (DISPID {dispid}) failed (0x{:08X}, EXCEPINFO {:?}, argument {:?})",
                *hresult as u32,
                exception_scode.map(|value| format!("0x{:08X}", value as u32)),
                argument_index
            ),
            Self::Conversion(error) => write!(formatter, "Automation conversion failed: {error:?}"),
            Self::Ownership { detail } => write!(formatter, "COM ownership failure: {detail}"),
            Self::Unsupported { detail } => {
                write!(formatter, "unsupported Automation operation: {detail}")
            }
        }
    }
}

impl Error for ExcelComError {}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn formatting_keeps_hresult_and_omits_addresses() {
        let error = ExcelComError::Invocation {
            member: "Quit",
            dispid: 302,
            hresult: -1,
            exception_scode: Some(0x800A_03EC_u32 as i32),
            argument_index: None,
        };
        let text = error.to_string();
        assert!(text.contains("0xFFFFFFFF"));
        assert!(text.contains("Quit"));
        assert!(!text.contains("ptr="));
    }
}
