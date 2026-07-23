use crate::automation::ConversionError;
use crate::automation::{ComCallDisposition, InvocationRetrySafety};
use std::error::Error;
use std::fmt::{Display, Formatter};
use std::time::Duration;

/// Runtime-specific failure that remains separate from Automation conversion
/// and invocation errors.
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum ExcelRuntimeError {
    /// No Excel instance is registered in the Running Object Table.
    NoRunningInstance,
    /// The Running Object Table exposed more than one indistinguishable target.
    AmbiguousRunningInstances,
    /// Excel disappeared while an attachment or owned-session operation ran.
    SessionDisappeared,
    /// A bounded retry sequence observed a busy server without a safe result.
    ExcelBusy {
        /// Number of calls attempted.
        attempts: u32,
        /// Time spent in the retry sequence.
        elapsed: Duration,
    },
    /// The original operation was retryable but not safe to replay.
    RetryUnsafe {
        /// Excel wrapper receiving the call.
        object: &'static str,
        /// Excel member that was not replayed.
        member: &'static str,
    },
    /// Registering or restoring a thread-local message filter failed.
    MessageFilterRegistrationFailed {
        /// The HRESULT returned by `CoRegisterMessageFilter`.
        hresult: i32,
    },
    /// The owned Excel process did not naturally exit before the timeout.
    ProcessExitTimeout {
        /// The owned Excel process id, when it was observed from its own window.
        process_id: Option<u32>,
        /// Requested wait duration.
        timeout: Duration,
    },
    /// Accessing the Running Object Table or active-object registration failed.
    RotAccessFailed {
        /// The HRESULT returned by the COM API.
        hresult: i32,
    },
}

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
        /// Classifies the original HRESULT without replacing it.
        disposition: ComCallDisposition,
        /// Whether repeating this exact invocation is safe after ambiguity.
        retry_safety: InvocationRetrySafety,
        /// Number of attempts made for this invocation.
        attempts: u32,
        /// Time spent in the invocation, including retry delays.
        elapsed: Duration,
        /// Optional server-provided EXCEPINFO source.
        exception_source: Option<String>,
        /// Optional server-provided EXCEPINFO description.
        exception_description: Option<String>,
    },
    /// An attachment, ownership, message-filter, or process runtime failure.
    Runtime(ExcelRuntimeError),
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
                disposition,
                retry_safety,
                attempts,
                elapsed,
                exception_source,
                exception_description,
            } => write!(
                formatter,
                "invocation of {object_type}.{member} (DISPID {dispid}, flags {dispatch_flags}) failed (0x{:08X}, EXCEPINFO {:?}, source {:?}, description {:?}, argument {:?}, disposition {:?}, retry safety {:?}, attempts {attempts}, elapsed {:?})",
                *hresult as u32,
                exception_scode.map(|value| format!("0x{:08X}", value as u32)),
                exception_source,
                exception_description,
                argument_index,
                disposition,
                retry_safety,
                elapsed,
            ),
            Self::Runtime(error) => match error {
                ExcelRuntimeError::NoRunningInstance => write!(formatter, "no running Excel instance is registered"),
                ExcelRuntimeError::AmbiguousRunningInstances => write!(formatter, "multiple ambiguous Excel instances are registered"),
                ExcelRuntimeError::SessionDisappeared => write!(formatter, "Excel session disappeared during the operation"),
                ExcelRuntimeError::ExcelBusy { attempts, elapsed } => write!(formatter, "Excel remained busy after {attempts} attempts over {elapsed:?}"),
                ExcelRuntimeError::RetryUnsafe { object, member } => write!(formatter, "retry of {object}.{member} is unsafe because it may have mutated Excel"),
                ExcelRuntimeError::MessageFilterRegistrationFailed { hresult } => write!(formatter, "COM message-filter registration failed (0x{:08X})", *hresult as u32),
                ExcelRuntimeError::ProcessExitTimeout { process_id, timeout } => write!(formatter, "owned Excel process {:?} did not exit within {timeout:?}", process_id),
                ExcelRuntimeError::RotAccessFailed { hresult } => write!(formatter, "Excel active-object lookup failed (0x{:08X})", *hresult as u32),
            },
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
            disposition: ComCallDisposition::PermanentFailure,
            retry_safety: InvocationRetrySafety::NonIdempotentWrite,
            attempts: 1,
            elapsed: Duration::ZERO,
            exception_source: None,
            exception_description: None,
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
