//! Workbook connection collections and typed connection classification.

use std::fmt::{Debug, Formatter};

use crate::ExcelComError;
use crate::automation::{OwnedVariant, invoke, property_put};
use crate::excel::collection::CollectionDescriptor;
use crate::excel::{DispatchObject, Range, Workbook};
use crate::internal::{ComPtr, Dispatch};
use crate::object_model::{MemberId, member};

use super::connection::ConnectionDetails;
use super::helpers::{
    TypedIter, collection_count, collection_enum, collection_index, collection_name, integer,
    object, string,
};
use super::{ConnectionType, OdbcConnection, OleDbConnection, TextConnection};

const CONNECTIONS: CollectionDescriptor = CollectionDescriptor {
    name: "Connections",
    count: MemberId::new("excel.connections.count"),
    item: MemberId::new("excel.connections.item"),
    new_enum: MemberId::new("excel.connections.newenum"),
};
const RANGES: CollectionDescriptor = CollectionDescriptor {
    name: "Ranges",
    count: MemberId::new("excel.ranges.count"),
    item: MemberId::new("excel.ranges.item"),
    new_enum: MemberId::new("excel.ranges.newenum"),
};

/// Apartment-bound collection of a workbook's external and worksheet connections.
pub struct Connections {
    inner: DispatchObject,
}
impl Debug for Connections {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("Connections").field(&self.inner).finish()
    }
}
impl Clone for Connections {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl Connections {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "Connections",
            },
        }
    }
    /// Returns the number of workbook connections.
    pub fn count(&self) -> Result<usize, ExcelComError> {
        collection_count(&self.inner, CONNECTIONS)
    }
    /// Returns a connection by its one-based Excel index.
    pub fn item_by_index(&self, index: usize) -> Result<WorkbookConnection, ExcelComError> {
        collection_index(
            &self.inner,
            CONNECTIONS,
            index,
            WorkbookConnection::from_dispatch,
        )
    }
    /// Returns a connection by its Excel-visible name after rejecting embedded NUL.
    pub fn item_by_name(&self, name: &str) -> Result<WorkbookConnection, ExcelComError> {
        collection_name(
            &self.inner,
            CONNECTIONS,
            name,
            WorkbookConnection::from_dispatch,
        )
    }
    /// Iterates connections in Excel's `_NewEnum` order on the owning apartment.
    pub fn iter(&self) -> Result<ConnectionsIter, ExcelComError> {
        Ok(ConnectionsIter {
            inner: TypedIter {
                enumerator: collection_enum(&self.inner, CONNECTIONS)?,
                index: 0,
                terminal: false,
                kind: "Connections",
                make: WorkbookConnection::from_dispatch,
            },
        })
    }
}

/// Fallible, fused, apartment-bound iterator over [`Connections`].
pub struct ConnectionsIter {
    inner: TypedIter<WorkbookConnection>,
}
impl Iterator for ConnectionsIter {
    type Item = Result<WorkbookConnection, ExcelComError>;
    fn next(&mut self) -> Option<Self::Item> {
        self.inner.next()
    }
}
impl std::iter::FusedIterator for ConnectionsIter {}

/// Apartment-bound wrapper for Excel's `WorkbookConnection` object.
pub struct WorkbookConnection {
    inner: DispatchObject,
}
impl Debug for WorkbookConnection {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("WorkbookConnection")
            .field(&self.inner)
            .finish()
    }
}
impl Clone for WorkbookConnection {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl WorkbookConnection {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "WorkbookConnection",
            },
        }
    }
    pub(crate) fn dispatch_object(&self) -> &DispatchObject {
        &self.inner
    }
    /// Returns the Excel-visible connection name.
    pub fn name(&self) -> Result<String, ExcelComError> {
        string(&self.inner, "excel.workbookconnection.name")
    }
    /// Returns the non-sensitive Excel connection description.
    pub fn description(&self) -> Result<String, ExcelComError> {
        string(&self.inner, "excel.workbookconnection.description")
    }
    /// Changes the description after rejecting embedded NUL.
    pub fn set_description(&self, value: &str) -> Result<(), ExcelComError> {
        let _ = property_put(
            &self.inner.dispatch,
            member(MemberId::new("excel.workbookconnection.description"), true),
            OwnedVariant::bstr(value)?,
        )?;
        Ok(())
    }
    /// Returns Excel's forward-compatible connection type.
    pub fn connection_type(&self) -> Result<ConnectionType, ExcelComError> {
        Ok(ConnectionType::from_raw(integer(
            &self.inner,
            "excel.workbookconnection.type",
        )?))
    }
    /// Returns whether Excel includes this connection in `Workbook.RefreshAll`.
    pub fn refresh_with_refresh_all(&self) -> Result<bool, ExcelComError> {
        super::helpers::bool_value(
            &self.inner,
            "excel.workbookconnection.refreshwithrefreshall",
        )
    }
    /// Includes or excludes this connection from `Workbook.RefreshAll`.
    pub fn set_refresh_with_refresh_all(&self, value: bool) -> Result<(), ExcelComError> {
        let _ = property_put(
            &self.inner.dispatch,
            member(
                MemberId::new("excel.workbookconnection.refreshwithrefreshall"),
                true,
            ),
            OwnedVariant::bool(value),
        )?;
        Ok(())
    }
    /// Returns Ranges consuming this connection as a snapshot vector.
    pub fn ranges(&self) -> Result<Vec<Range>, ExcelComError> {
        let ranges = object(&self.inner, "excel.workbookconnection.ranges", |dispatch| {
            DispatchObject {
                dispatch,
                kind: "Ranges",
            }
        })?;
        let mut values = Vec::new();
        let mut iterator = TypedIter {
            enumerator: collection_enum(&ranges, RANGES)?,
            index: 0,
            terminal: false,
            kind: "Ranges",
            make: Range::from_dispatch,
        };
        for value in &mut iterator {
            values.push(value?);
        }
        Ok(values)
    }
    /// Requests an Excel-owned refresh; providers and host policy determine completion.
    pub fn refresh(&self) -> Result<(), ExcelComError> {
        let _ = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.workbookconnection.refresh"), false),
            vec![],
            false,
        )?;
        Ok(())
    }
    /// Deletes the connection and consumes its invalidated wrapper.
    pub fn delete(self) -> Result<(), ExcelComError> {
        let _ = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.workbookconnection.delete"), false),
            vec![],
            false,
        )?;
        Ok(())
    }
    /// Returns a typed, non-raw subtype wrapper where Excel exposes one.
    pub fn details(&self) -> Result<ConnectionDetails, ExcelComError> {
        let kind = self.connection_type()?;
        match kind {
            ConnectionType::OLE_DB => Ok(ConnectionDetails::OleDb(object(
                &self.inner,
                "excel.workbookconnection.oledbconnection",
                OleDbConnection::from_dispatch,
            )?)),
            ConnectionType::ODBC => Ok(ConnectionDetails::Odbc(object(
                &self.inner,
                "excel.workbookconnection.odbcconnection",
                OdbcConnection::from_dispatch,
            )?)),
            ConnectionType::TEXT => Ok(ConnectionDetails::Text(object(
                &self.inner,
                "excel.workbookconnection.textconnection",
                TextConnection::from_dispatch,
            )?)),
            // The installed Excel type library exposes `TextConnection` but
            // no `WorkbookConnection.WebConnection` detail member.
            ConnectionType::WEB => Ok(ConnectionDetails::Unsupported(kind)),
            ConnectionType::WORKSHEET => Ok(ConnectionDetails::Worksheet),
            ConnectionType::MODEL => Ok(ConnectionDetails::Model),
            _ => Ok(ConnectionDetails::Unsupported(kind)),
        }
    }
}

impl Workbook {
    /// Returns this workbook's typed external-data connections collection.
    pub fn connections(&self) -> Result<Connections, ExcelComError> {
        object(
            self.dispatch_object(),
            "excel.workbook.connections",
            Connections::from_dispatch,
        )
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn connection_constants_preserve_unknown_values() {
        assert_eq!(ConnectionType::from_raw(42).raw(), 42);
    }
    #[test]
    fn collection_descriptors_are_exact() {
        assert_eq!(CONNECTIONS.item.as_str(), "excel.connections.item");
        assert_eq!(RANGES.new_enum.as_str(), "excel.ranges.newenum");
    }
}
