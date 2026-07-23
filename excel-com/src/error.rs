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
    /// Querying a COM interface failed.
    QueryInterface {
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
    /// An `IEnumVARIANT` operation or yielded collection value was invalid.
    Enumeration {
        /// The collection being iterated.
        collection: &'static str,
        /// Zero-based item position requested from the enumerator.
        item_index: usize,
        /// The HRESULT reported by `IEnumVARIANT::Next`, if applicable.
        hresult: Option<i32>,
        /// A static description of the validated failure.
        detail: &'static str,
    },
    /// Calling an Automation member failed.
    Invocation {
        /// The Excel wrapper object receiving the call.
        object_type: &'static str,
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
        /// The `DISPATCH_*` flags supplied to `IDispatch::Invoke`.
        dispatch_flags: u16,
    },
    /// A value could not be converted before or after an Automation call.
    Conversion(ConversionError),
    /// A COM ownership invariant was violated.
    Ownership {
        /// A static description of the invariant.
        detail: &'static str,
    },
    /// A Windows path cannot be represented safely at the BSTR boundary.
    InvalidPath {
        /// A static description that intentionally excludes the caller path.
        detail: &'static str,
    },
    /// The requested operation is intentionally outside this crate's slice.
    Unsupported {
        /// A static description of the unsupported operation.
        detail: &'static str,
    },
    /// An operation and the state-restoration attempt both failed.
    ///
    /// Temporary Excel state guards return this variant so callers retain the
    /// original failure as well as the error that prevented restoration.
    StateRestoration {
        /// The error from the protected operation.
        operation: Box<ExcelComError>,
        /// The error from restoring the previous Excel state.
        restoration: Box<ExcelComError>,
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
            Self::QueryInterface { hresult } => write!(
                formatter,
                "COM QueryInterface failed (0x{:08X})",
                *hresult as u32
            ),
            Self::NameLookup { member, hresult } => write!(
                formatter,
                "name lookup for {member} failed (0x{:08X})",
                *hresult as u32
            ),
            Self::Enumeration {
                collection,
                item_index,
                hresult,
                detail,
            } => write!(
                formatter,
                "enumeration of {collection} at item {item_index} failed ({detail}, HRESULT {:?})",
                hresult.map(|value| format!("0x{:08X}", value as u32))
            ),
            Self::Invocation {
                object_type,
                member,
                dispid,
                hresult,
                exception_scode,
                argument_index,
                dispatch_flags,
            } => write!(
                formatter,
                "invocation of {object_type}.{member} (DISPID {dispid}, flags {dispatch_flags}) failed (0x{:08X}, EXCEPINFO {:?}, argument {:?})",
                *hresult as u32,
                exception_scode.map(|value| format!("0x{:08X}", value as u32)),
                argument_index
            ),
            Self::Conversion(error) => write!(formatter, "Automation conversion failed: {error:?}"),
            Self::Ownership { detail } => write!(formatter, "COM ownership failure: {detail}"),
            Self::InvalidPath { detail } => write!(formatter, "invalid Windows path: {detail}"),
            Self::Unsupported { detail } => {
                write!(formatter, "unsupported Automation operation: {detail}")
            }
            Self::StateRestoration {
                operation,
                restoration,
            } => write!(
                formatter,
                "Automation operation failed ({operation}); restoring Excel state also failed ({restoration})"
            ),
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
            object_type: "Application",
            member: "Quit",
            dispid: 302,
            hresult: -1,
            exception_scode: Some(0x800A_03EC_u32 as i32),
            argument_index: None,
            dispatch_flags: 1,
        };
        let text = error.to_string();
        assert!(text.contains("0xFFFFFFFF"));
        assert!(text.contains("Quit"));
        assert!(!text.contains("ptr="));
    }

    #[test]
    fn state_restoration_retains_both_errors() {
        let error = ExcelComError::StateRestoration {
            operation: Box::new(ExcelComError::Unsupported {
                detail: "operation failed",
            }),
            restoration: Box::new(ExcelComError::Unsupported {
                detail: "restore failed",
            }),
        };
        let text = error.to_string();
        assert!(text.contains("operation failed"));
        assert!(text.contains("restore failed"));
    }
}
