use std::error::Error;
use std::fmt::{Display, Formatter};

/// Production-facing Automation failure without raw address disclosure.
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum ExcelComError {
    Initialization {
        hresult: i32,
    },
    Activation {
        hresult: i32,
    },
    NameLookup {
        member: &'static str,
        hresult: i32,
    },
    Invocation {
        member: &'static str,
        dispid: i32,
        hresult: i32,
        exception_scode: Option<i32>,
        argument_index: Option<u32>,
    },
    Conversion {
        detail: &'static str,
    },
    Ownership {
        detail: &'static str,
    },
    Unsupported {
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
            Self::Conversion { detail } => {
                write!(formatter, "Automation conversion failed: {detail}")
            }
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
