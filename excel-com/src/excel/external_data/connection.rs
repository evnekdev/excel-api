//! Typed connection-detail selection without exposing generic dispatch values.

use super::{ConnectionType, OdbcConnection, OleDbConnection, TextConnection, WebConnection};

/// The safe typed detail object selected from a [`super::WorkbookConnection`].
#[derive(Debug)]
pub enum ConnectionDetails {
    /// OLE DB provider metadata.
    OleDb(OleDbConnection),
    /// ODBC provider metadata.
    Odbc(OdbcConnection),
    /// Text connection metadata where exposed by Excel.
    Text(TextConnection),
    /// Web connection metadata where exposed by Excel; this crate does not contact it.
    Web(WebConnection),
    /// Worksheet-backed connection with no additional wrapper in this slice.
    Worksheet,
    /// Data Model connection with no DAX or model mutation support.
    Model,
    /// A future or unsupported Excel connection type.
    Unsupported(ConnectionType),
}
