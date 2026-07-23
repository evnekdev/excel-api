//! External data connections, persistent QueryTables, and bounded refresh control.
//!
//! The wrappers remain apartment-bound and never expose raw COM pointers.
//! Connection strings and Power Query formulas use [`SecretStringValue`],
//! whose diagnostic representations are redacted.
//!
//! Connections can be enumerated without exposing their implementation
//! pointers:
//!
//! ```no_run
//! # use excel_com::{ExcelComError, Workbook};
//! # fn example(workbook: &Workbook) -> Result<(), ExcelComError> {
//! for connection in workbook.connections()?.iter()? {
//!     let connection = connection?;
//!     println!("{}: {:?}", connection.name()?, connection.connection_type()?);
//! }
//! # Ok(())
//! # }
//! ```
//!
//! A local text QueryTable is a persistent, refreshable import; this is not
//! the same operation as the one-time `OpenText` import:
//!
//! ```no_run
//! # use std::path::Path;
//! # use excel_com::{ExcelComError, TextDelimiter, TextParsingType, TextQueryAddOptions, TextQualifier, Worksheet};
//! # fn example(worksheet: &Worksheet, csv_path: &Path) -> Result<(), ExcelComError> {
//! let query = worksheet.query_tables()?.add_from_local_text(&TextQueryAddOptions {
//!     path: csv_path,
//!     destination: &worksheet.range("A1")?,
//!     parsing_type: TextParsingType::DELIMITED,
//!     delimiter: Some(TextDelimiter::Comma),
//!     text_qualifier: Some(TextQualifier::DOUBLE_QUOTE),
//!     columns: vec![],
//!     refresh_on_file_open: false,
//!     background_query: false,
//! })?;
//! query.refresh(Some(false))?;
//! # Ok(())
//! # }
//! ```
//!
//! Refresh waiting is bounded and remains on the Excel STA; it never starts a
//! background Rust thread:
//!
//! ```no_run
//! # use std::time::Duration;
//! # use excel_com::{ExcelComError, RefreshWaitOptions, Workbook};
//! # fn example(workbook: &Workbook) -> Result<(), ExcelComError> {
//! let report = workbook.wait_for_refresh(RefreshWaitOptions {
//!     timeout: Duration::from_secs(30),
//!     poll_interval: Duration::from_millis(100),
//! })?;
//! assert!(report.completed || report.remaining_queries > 0);
//! # Ok(())
//! # }
//! ```

mod connection;
mod connections;
mod helpers;
mod odbc;
mod oledb;
mod query_table;
mod query_tables;
mod refresh;
mod text_connection;
mod types;
mod web_connection;
mod workbook_queries;

pub use connection::ConnectionDetails;
pub use connections::{Connections, ConnectionsIter, WorkbookConnection};
pub use odbc::OdbcConnection;
pub use oledb::OleDbConnection;
pub use query_table::QueryTable;
pub use query_tables::{QueryTables, QueryTablesIter};
pub use text_connection::TextConnection;
pub use types::*;
pub use web_connection::WebConnection;
pub use workbook_queries::{WorkbookQueries, WorkbookQueriesIter, WorkbookQuery};

#[cfg(test)]
mod tests;
