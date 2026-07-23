//! QueryTable collection access and safe local-text creation.

use std::ffi::OsString;
use std::fmt::{Debug, Formatter};

use crate::automation::{
    AutomationArray, AutomationValue, OwnedVariant, PositionalArguments, encode_variant, invoke,
    property_put,
};
use crate::excel::collection::CollectionDescriptor;
use crate::excel::{DispatchObject, Range, TextDelimiter, TextParsingType, Worksheet};
use crate::internal::{ComPtr, Dispatch, path_bstr};
use crate::object_model::{MemberId, member};
use crate::{ConversionPolicy, ExcelComError};

use super::helpers::{
    TypedIter, collection_count, collection_enum, collection_index, collection_name, object,
};
use super::{QueryTable, TextQueryAddOptions, WorkbookConnection};

const QUERY_TABLES: CollectionDescriptor = CollectionDescriptor {
    name: "QueryTables",
    count: MemberId::new("excel.querytables.count"),
    item: MemberId::new("excel.querytables.item"),
    new_enum: MemberId::new("excel.querytables.newenum"),
};

type DelimiterFlags = (bool, bool, bool, bool, Option<char>);

/// Apartment-bound collection of persistent worksheet QueryTables.
pub struct QueryTables {
    inner: DispatchObject,
}
impl Debug for QueryTables {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("QueryTables").field(&self.inner).finish()
    }
}
impl Clone for QueryTables {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl QueryTables {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "QueryTables",
            },
        }
    }
    /// Returns the number of persistent QueryTables on this sheet.
    pub fn count(&self) -> Result<usize, ExcelComError> {
        collection_count(&self.inner, QUERY_TABLES)
    }
    /// Returns a QueryTable by its one-based Excel index.
    pub fn item_by_index(&self, index: usize) -> Result<QueryTable, ExcelComError> {
        collection_index(&self.inner, QUERY_TABLES, index, QueryTable::from_dispatch)
    }
    /// Returns a QueryTable by its Excel-visible name.
    pub fn item_by_name(&self, name: &str) -> Result<QueryTable, ExcelComError> {
        collection_name(&self.inner, QUERY_TABLES, name, QueryTable::from_dispatch)
    }
    /// Iterates QueryTables in Excel's `_NewEnum` order on the owning apartment.
    pub fn iter(&self) -> Result<QueryTablesIter, ExcelComError> {
        Ok(QueryTablesIter {
            inner: TypedIter {
                enumerator: collection_enum(&self.inner, QUERY_TABLES)?,
                index: 0,
                terminal: false,
                kind: "QueryTables",
                make: QueryTable::from_dispatch,
            },
        })
    }
    /// Adds a QueryTable backed by an existing workbook connection.
    ///
    /// The supplied connection must belong to the same workbook; Excel
    /// performs that ownership and destination validation.
    pub fn add_from_connection(
        &self,
        connection: &WorkbookConnection,
        destination: &Range,
    ) -> Result<QueryTable, ExcelComError> {
        let mut args = PositionalArguments::new();
        args.push_object(connection.dispatch_object());
        args.push_object(destination.dispatch_object());
        args.push_optional(None); // SqlText
        let mut value = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.querytables.add"), false),
            args.into_inner(),
            false,
        )?;
        Ok(QueryTable::from_dispatch(value.take_dispatch()?))
    }
    /// Adds a persistent QueryTable backed by an owned local text file.
    ///
    /// Unlike `Workbooks.OpenText`, this import retains refresh settings and
    /// may create a workbook connection. It does not make any network request.
    pub fn add_from_local_text(
        &self,
        options: &TextQueryAddOptions<'_>,
    ) -> Result<QueryTable, ExcelComError> {
        validate_owned_local_file(options.path)?;
        let source = text_connection_source_path(options.path)?;
        let mut args = PositionalArguments::new();
        args.push_required(source);
        args.push_object(options.destination.dispatch_object());
        args.push_optional(None); // SqlText
        let mut value = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.querytables.add"), false),
            args.into_inner(),
            false,
        )?;
        let query = QueryTable::from_dispatch(value.take_dispatch()?);
        if let Err(error) = configure_local_text(&query, options) {
            let _ = query.clone().delete();
            return Err(error);
        }
        Ok(query)
    }
}

fn validate_owned_local_file(path: &std::path::Path) -> Result<(), ExcelComError> {
    if path.to_string_lossy().starts_with(r"\\") {
        return Err(ExcelComError::InvalidPath {
            detail: "network paths are excluded from local text QueryTable creation",
        });
    }
    if !path.is_file() {
        return Err(ExcelComError::InvalidPath {
            detail: "local text QueryTable source must be an existing regular file",
        });
    }
    Ok(())
}

/// Fallible, fused, apartment-bound iterator over [`QueryTables`].
pub struct QueryTablesIter {
    inner: TypedIter<QueryTable>,
}
impl Iterator for QueryTablesIter {
    type Item = Result<QueryTable, ExcelComError>;
    fn next(&mut self) -> Option<Self::Item> {
        self.inner.next()
    }
}
impl std::iter::FusedIterator for QueryTablesIter {}

impl Worksheet {
    /// Returns the persistent QueryTables associated with this worksheet.
    pub fn query_tables(&self) -> Result<QueryTables, ExcelComError> {
        object(
            self.dispatch_object(),
            "excel.worksheet.querytables",
            QueryTables::from_dispatch,
        )
    }
}

fn text_connection_source_path(path: &std::path::Path) -> Result<OwnedVariant, ExcelComError> {
    let mut source = OsString::from("TEXT;");
    source.push(path.as_os_str());
    path_bstr(std::path::Path::new(&source))
}

fn configure_local_text(
    query: &QueryTable,
    options: &TextQueryAddOptions<'_>,
) -> Result<(), ExcelComError> {
    put_i32(
        query,
        "excel.querytable.textfileparsetype",
        options.parsing_type.raw(),
    )?;
    if let Some(qualifier) = options.text_qualifier {
        put_i32(
            query,
            "excel.querytable.textfiletextqualifier",
            qualifier.raw(),
        )?;
    }
    let (tab, semicolon, comma, space, other_char) = delimiter_flags(options.delimiter.as_ref())?;
    put_bool(query, "excel.querytable.textfiletabdelimiter", tab)?;
    put_bool(
        query,
        "excel.querytable.textfilesemicolondelimiter",
        semicolon,
    )?;
    put_bool(query, "excel.querytable.textfilecommadelimiter", comma)?;
    put_bool(query, "excel.querytable.textfilespacedelimiter", space)?;
    if let Some(other_char) = other_char {
        put_text(
            query,
            "excel.querytable.textfileotherdelimiter",
            &other_char.to_string(),
        )?;
    }
    if let Some(columns) = field_info(&options.columns, options.parsing_type)? {
        put_variant(query, "excel.querytable.textfilecolumndatatypes", columns)?;
    }
    put_bool(
        query,
        "excel.querytable.refreshonfileopen",
        options.refresh_on_file_open,
    )?;
    put_bool(
        query,
        "excel.querytable.backgroundquery",
        options.background_query,
    )
}

fn delimiter_flags(delimiter: Option<&TextDelimiter>) -> Result<DelimiterFlags, ExcelComError> {
    let result = match delimiter {
        None => (false, false, false, false, None),
        Some(TextDelimiter::Tab) => (true, false, false, false, None),
        Some(TextDelimiter::Semicolon) => (false, true, false, false, None),
        Some(TextDelimiter::Comma) => (false, false, true, false, None),
        Some(TextDelimiter::Space) => (false, false, false, true, None),
        Some(TextDelimiter::Other(value)) => (false, false, false, false, Some(*value)),
        Some(TextDelimiter::Custom {
            tab,
            semicolon,
            comma,
            space,
            other,
        }) => (*tab, *semicolon, *comma, *space, *other),
    };
    if result.4 == Some('\0') {
        return Err(ExcelComError::Unsupported {
            detail: "TextDelimiter::Other cannot be NUL",
        });
    }
    Ok(result)
}

fn field_info(
    columns: &[crate::TextColumnSpec],
    parsing: TextParsingType,
) -> Result<Option<OwnedVariant>, ExcelComError> {
    if columns.is_empty() {
        return Ok(None);
    }
    let mut values = Vec::with_capacity(columns.len() * 2);
    let mut previous = None;
    for column in columns {
        let start = column
            .start
            .unwrap_or_else(|| previous.map_or(1, |value: usize| value + 1));
        if parsing != TextParsingType::FIXED_WIDTH && start == 0 {
            return Err(ExcelComError::Unsupported {
                detail: "delimited FieldInfo columns are one-based and nonzero",
            });
        }
        if parsing == TextParsingType::FIXED_WIDTH && previous.is_some_and(|value| start <= value) {
            return Err(ExcelComError::Unsupported {
                detail: "fixed-width FieldInfo starts must strictly increase",
            });
        }
        previous = Some(start);
        values.push(AutomationValue::Number(f64::from(
            i32::try_from(start).map_err(|_| ExcelComError::Unsupported {
                detail: "FieldInfo start exceeds i32",
            })?,
        )));
        values.push(AutomationValue::Number(f64::from(column.column_type.raw())));
    }
    Ok(Some(encode_variant(
        &AutomationValue::Array(AutomationArray::new(columns.len(), 2, values)?),
        ConversionPolicy::default(),
    )?))
}

fn put_bool(query: &QueryTable, id: &'static str, value: bool) -> Result<(), ExcelComError> {
    put_variant(query, id, OwnedVariant::bool(value))
}
fn put_i32(query: &QueryTable, id: &'static str, value: i32) -> Result<(), ExcelComError> {
    put_variant(query, id, OwnedVariant::i32(value))
}
fn put_text(query: &QueryTable, id: &'static str, value: &str) -> Result<(), ExcelComError> {
    put_variant(query, id, OwnedVariant::bstr(value)?)
}
fn put_variant(
    query: &QueryTable,
    id: &'static str,
    value: OwnedVariant,
) -> Result<(), ExcelComError> {
    let _ = property_put(
        &query.dispatch_object().dispatch,
        member(MemberId::new(id), true),
        value,
    )?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::path::Path;
    #[test]
    fn local_text_connection_is_not_a_network_url() {
        let path = Path::new("fixtures/local.csv");
        let text = text_connection_source_path(path)
            .expect("source")
            .as_string()
            .expect("text")
            .to_owned();
        assert!(text.starts_with("TEXT;"));
    }
    #[test]
    fn network_unc_source_is_rejected_before_excel() {
        assert!(validate_owned_local_file(Path::new(r"\\server\share\data.csv")).is_err());
    }
    #[test]
    fn field_info_reuses_delimited_one_based_encoding() {
        assert!(
            field_info(
                &[crate::TextColumnSpec {
                    start: Some(0),
                    column_type: crate::TextColumnType::GENERAL
                }],
                TextParsingType::DELIMITED
            )
            .is_err()
        );
    }
}
