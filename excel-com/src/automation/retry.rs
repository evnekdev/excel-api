use std::cell::RefCell;
use std::time::Duration;

use windows_sys::Win32::Foundation::{
    CO_E_SERVER_EXEC_FAILURE, RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_REJECTED,
    RPC_E_SERVERCALL_RETRYLATER,
};

/// Classifies a COM call result without discarding its original HRESULT.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum ComCallDisposition {
    /// Excel is temporarily busy and a safe call may be retried.
    RetryableBusy,
    /// Excel explicitly rejected the call and a safe call may be retried.
    RetryableRejected,
    /// The failure is known not to be transient.
    PermanentFailure,
    /// The HRESULT has no special retry classification.
    Unknown,
}

/// Describes whether repeating an Automation call can be safe after ambiguity.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum InvocationRetrySafety {
    /// A property read does not mutate the Excel object model.
    SafeRead,
    /// A clearly idempotent property assignment may be repeated.
    IdempotentWrite,
    /// Repeating a method could duplicate a visible Excel operation.
    NonIdempotentWrite,
    /// The wrapper cannot establish safe retry semantics.
    Unknown,
}

/// Bounded, opt-in retry settings for transient COM Automation failures.
///
/// The policy applies only while a [`crate::ComMessageFilterGuard`] is installed
/// for the current STA thread. Methods are deliberately not retried because a
/// rejected call can have completed at the server before COM reports failure.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct ComRetryPolicy {
    /// Maximum total invocations, including the first call.
    pub max_attempts: u32,
    /// Delay before the first retry.
    pub initial_delay: Duration,
    /// Upper bound for exponential retry delay.
    pub maximum_delay: Duration,
    /// Upper bound for the whole retry sequence.
    pub total_timeout: Duration,
    /// Whether `RPC_E_CALL_REJECTED` and `RPC_E_SERVERCALL_REJECTED` may retry.
    pub retry_call_rejected: bool,
    /// Whether `RPC_E_SERVERCALL_RETRYLATER` may retry.
    pub retry_server_busy: bool,
}

impl Default for ComRetryPolicy {
    fn default() -> Self {
        Self {
            max_attempts: 3,
            initial_delay: Duration::from_millis(50),
            maximum_delay: Duration::from_millis(250),
            total_timeout: Duration::from_secs(1),
            retry_call_rejected: true,
            retry_server_busy: true,
        }
    }
}

impl ComRetryPolicy {
    pub(crate) fn delay_for_retry(&self, completed_attempts: u32) -> Duration {
        let multiplier = 1_u32.checked_shl(completed_attempts.saturating_sub(1)).unwrap_or(u32::MAX);
        self.initial_delay
            .checked_mul(multiplier)
            .unwrap_or(self.maximum_delay)
            .min(self.maximum_delay)
    }
}

/// Returns the conservative classification for a COM Automation HRESULT.
pub const fn classify_com_hresult(hresult: i32) -> ComCallDisposition {
    const EXCEL_APPLICATION_ERROR: i32 = 0x800A_03EC_u32 as i32;
    match hresult {
        RPC_E_SERVERCALL_RETRYLATER => ComCallDisposition::RetryableBusy,
        RPC_E_CALL_REJECTED | RPC_E_SERVERCALL_REJECTED => ComCallDisposition::RetryableRejected,
        // Excel application errors (including 0x800A03EC) are application
        // failures, not a signal that replaying the call is safe.
        CO_E_SERVER_EXEC_FAILURE | EXCEL_APPLICATION_ERROR => ComCallDisposition::PermanentFailure,
        _ => ComCallDisposition::Unknown,
    }
}

thread_local! {
    static ACTIVE_POLICY: RefCell<Option<ComRetryPolicy>> = const { RefCell::new(None) };
}

pub(crate) fn active_policy() -> Option<ComRetryPolicy> {
    ACTIVE_POLICY.with(|policy| policy.borrow().clone())
}

pub(crate) fn replace_active_policy(policy: Option<ComRetryPolicy>) -> Option<ComRetryPolicy> {
    ACTIVE_POLICY.with(|active| std::mem::replace(&mut *active.borrow_mut(), policy))
}
