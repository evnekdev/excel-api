//! Public value types for external data and refresh operations.

use std::fmt::{Debug, Display, Formatter};
use std::path::Path;
use std::time::Duration;

use crate::excel::{Range, TextColumnSpec, TextDelimiter, TextParsingType, TextQualifier};
use crate::{AutomationValue, ExcelComError};

macro_rules! raw_type {
    ($(#[$docs:meta])* $name:ident { $($constant:ident = $value:expr => $constant_docs:literal;)* }) => {
        $(#[$docs])*
        #[derive(Clone, Copy, Debug, Eq, PartialEq, Hash)]
        #[repr(transparent)]
        pub struct $name(i32);
        impl $name {
            $(#[doc = $constant_docs] pub const $constant: Self = Self($value);)*
            /// Preserves an installed or future Excel value.
            pub const fn from_raw(value: i32) -> Self { Self(value) }
            /// Returns the exact Excel type-library value.
            pub const fn raw(self) -> i32 { self.0 }
        }
    };
}

raw_type! {
    /// A forward-compatible `XlConnectionType` value.
    ConnectionType {
        OLE_DB = 1 => "`xlConnectionTypeOLEDB`.";
        ODBC = 2 => "`xlConnectionTypeODBC`.";
        TEXT = 4 => "`xlConnectionTypeTEXT`.";
        WEB = 5 => "`xlConnectionTypeWEB`.";
        DATA_FEED = 6 => "`xlConnectionTypeDATAFEED`.";
        MODEL = 7 => "`xlConnectionTypeMODEL`.";
        WORKSHEET = 8 => "`xlConnectionTypeWORKSHEET`.";
    }
}
raw_type! {
    /// A forward-compatible Excel command type.
    CommandType {
        CUBE_NAME = 1 => "`xlCmdCube`.";
        SQL = 2 => "`xlCmdSql`.";
        TABLE = 3 => "`xlCmdTable`.";
        DEFAULT = 4 => "`xlCmdDefault`.";
        LIST = 5 => "`xlCmdList`.";
    }
}
raw_type! {
    /// A forward-compatible `XlCredentialsMethod` value.
    CredentialsMethod {
        INTEGRATED = 0 => "Integrated authentication.";
        NONE = 1 => "No credentials are supplied by Excel.";
        STORED = 2 => "Credentials are stored by the provider.";
    }
}
raw_type! {
    /// A forward-compatible Office connection-file type.
    ConnectionFileType {
        ODC = 0 => "Office Data Connection (`.odc`).";
        UDC = 1 => "Universal Data Connection (`.udcx`).";
    }
}

/// A connection string or formula that may contain credentials, endpoints, or tokens.
///
/// Its `Debug` and `Display` forms are deliberately redacted. Use
/// [`Self::expose_secret`] only when the caller has an explicit, local need to
/// inspect the value; never place that value in logs or evidence.
#[derive(Clone, Eq, PartialEq)]
pub struct SecretStringValue {
    value: String,
}
impl SecretStringValue {
    /// Creates a redacted string after rejecting embedded NUL.
    pub fn new(value: impl Into<String>) -> Result<Self, ExcelComError> {
        let value = value.into();
        if value.contains('\0') {
            return Err(ExcelComError::Unsupported {
                detail: "Excel Automation text cannot contain embedded NUL",
            });
        }
        Ok(Self { value })
    }
    /// Returns the sensitive string explicitly; callers must avoid logging it.
    pub fn expose_secret(&self) -> &str {
        &self.value
    }
}
impl Debug for SecretStringValue {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.write_str("SecretStringValue(REDACTED)")
    }
}
impl Display for SecretStringValue {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.write_str("REDACTED")
    }
}

/// Bounded options for adding an OLE DB connection without a password setter.
#[derive(Debug)]
pub struct OleDbConnectionAddOptions<'a> {
    /// Excel-visible connection name.
    pub name: &'a str,
    /// Optional non-sensitive description.
    pub description: Option<&'a str>,
    /// Redacted connection string; callers must own and review it.
    pub connection_string: &'a SecretStringValue,
    /// Provider command text, preserved as an Automation value.
    pub command_text: &'a AutomationValue,
    /// Optional provider command classification.
    pub command_type: Option<CommandType>,
    /// Requests a model connection only when Excel supports it.
    pub create_model_connection: Option<bool>,
    /// Requests relationship import only when Excel supports it.
    pub import_relationships: Option<bool>,
}

/// Bounded options for adding an ODBC connection without a password setter.
#[derive(Debug)]
pub struct OdbcConnectionAddOptions<'a> {
    /// Excel-visible connection name.
    pub name: &'a str,
    /// Optional non-sensitive description.
    pub description: Option<&'a str>,
    /// Redacted connection string; callers must own and review it.
    pub connection_string: &'a SecretStringValue,
    /// Provider command text, preserved as an Automation value.
    pub command_text: &'a AutomationValue,
    /// Optional provider command classification.
    pub command_type: Option<CommandType>,
}

/// Input for a persistent QueryTable backed only by an owned local text file.
#[derive(Debug)]
pub struct TextQueryAddOptions<'a> {
    /// Local path passed directly to Excel without canonicalization.
    pub path: &'a Path,
    /// QueryTable destination in the target worksheet.
    pub destination: &'a Range,
    /// Excel parsing mode.
    pub parsing_type: TextParsingType,
    /// Optional delimiter selection.
    pub delimiter: Option<TextDelimiter>,
    /// Optional text qualifier.
    pub text_qualifier: Option<TextQualifier>,
    /// Reused Excel `FieldInfo` column specifications.
    pub columns: Vec<TextColumnSpec>,
    /// Whether Excel refreshes this import while opening the workbook.
    pub refresh_on_file_open: bool,
    /// Whether Excel may refresh it in the background.
    pub background_query: bool,
}

/// Explicit STA-thread polling bounds for [`crate::Workbook::wait_for_refresh`].
#[derive(Clone, Copy, Debug)]
pub struct RefreshWaitOptions {
    /// Maximum elapsed time before returning a completed=false report.
    pub timeout: Duration,
    /// Time between refresh-state checks on the owning apartment thread.
    pub poll_interval: Duration,
}
impl RefreshWaitOptions {
    pub(crate) fn validate(self) -> Result<(), ExcelComError> {
        if self.timeout.is_zero() || self.poll_interval.is_zero() {
            return Err(ExcelComError::Unsupported {
                detail: "refresh timeout and poll interval must be nonzero",
            });
        }
        Ok(())
    }
}

/// Result of a bounded refresh wait; a timeout is not a COM failure.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct RefreshWaitReport {
    /// Whether all observable refreshes completed before the timeout.
    pub completed: bool,
    /// Time spent polling on the owning STA thread.
    pub elapsed: Duration,
    /// Number of QueryTables that still reported `Refreshing`.
    pub remaining_queries: usize,
}

/// Best-effort cancellation counts; unsupported refresh kinds are reported separately.
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct RefreshCancellationReport {
    /// QueryTables for which `CancelRefresh` was invoked.
    pub query_tables_cancelled: usize,
    /// Workbook connections for which `CancelRefresh` was invoked.
    pub connections_cancelled: usize,
    /// Refreshable objects without an exposed cancellation operation.
    pub unsupported_refreshes: usize,
}
