//! Read-only Power Query (`WorkbookQuery`) inspection.

use std::fmt::{Debug, Formatter};

use crate::ExcelComError;
use crate::automation::{EnumVariant, OwnedVariant, enumerated_dispatch, invoke};
use crate::excel::collection::{CollectionDescriptor, count, enumerator};
use crate::excel::{DispatchObject, Workbook};
use crate::internal::{ComPtr, Dispatch};
use crate::object_model::{MemberId, member};

use super::SecretStringValue;
use super::helpers::{object, string};

const QUERIES: CollectionDescriptor = CollectionDescriptor {
    name: "WorkbookQueries",
    count: MemberId::new("excel.queries.count"),
    item: MemberId::new("excel.queries.item"),
    new_enum: MemberId::new("excel.queries.newenum"),
};

/// Apartment-bound workbook Power Query collection, where the installed Excel version exposes it.
pub struct WorkbookQueries {
    inner: DispatchObject,
}
impl Debug for WorkbookQueries {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("WorkbookQueries").field(&self.inner).finish()
    }
}
impl Clone for WorkbookQueries {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl WorkbookQueries {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "WorkbookQueries",
            },
        }
    }
    /// Returns the number of workbook queries.
    pub fn count(&self) -> Result<usize, ExcelComError> {
        count(&self.inner, QUERIES)
    }
    /// Returns a query by one-based index. Excel declares `Queries.Item` as a method.
    pub fn item_by_index(&self, index: usize) -> Result<WorkbookQuery, ExcelComError> {
        if index == 0 {
            return Err(ExcelComError::Unsupported {
                detail: "collection index is one-based",
            });
        }
        self.item(OwnedVariant::i32(i32::try_from(index).map_err(|_| {
            ExcelComError::Unsupported {
                detail: "query index exceeds i32",
            }
        })?))
    }
    /// Returns a query by its Excel-visible name after rejecting embedded NUL.
    pub fn item_by_name(&self, name: &str) -> Result<WorkbookQuery, ExcelComError> {
        self.item(OwnedVariant::bstr(name)?)
    }
    /// Iterates workbook queries on the owning apartment thread.
    pub fn iter(&self) -> Result<WorkbookQueriesIter, ExcelComError> {
        Ok(WorkbookQueriesIter {
            enumerator: enumerator(&self.inner, QUERIES)?,
            index: 0,
            terminal: false,
        })
    }
    fn item(&self, value: OwnedVariant) -> Result<WorkbookQuery, ExcelComError> {
        let mut value = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.queries.item"), false),
            vec![value],
            false,
        )?;
        Ok(WorkbookQuery::from_dispatch(value.take_dispatch()?))
    }
}

/// Fallible, fused, apartment-bound iterator over [`WorkbookQueries`].
pub struct WorkbookQueriesIter {
    enumerator: EnumVariant,
    index: usize,
    terminal: bool,
}
impl Iterator for WorkbookQueriesIter {
    type Item = Result<WorkbookQuery, ExcelComError>;
    fn next(&mut self) -> Option<Self::Item> {
        if self.terminal {
            return None;
        }
        match self.enumerator.next() {
            Ok(Some(mut value)) => {
                let index = self.index;
                self.index += 1;
                Some(
                    enumerated_dispatch(&mut value, "WorkbookQueries", index)
                        .map(WorkbookQuery::from_dispatch),
                )
            }
            Ok(None) => {
                self.terminal = true;
                None
            }
            Err(error) => {
                self.terminal = true;
                Some(Err(error))
            }
        }
    }
}
impl std::iter::FusedIterator for WorkbookQueriesIter {}

/// Apartment-bound Power Query definition. M formulas are redacted because
/// they can embed endpoints, provider options, and credentials.
pub struct WorkbookQuery {
    inner: DispatchObject,
}
impl Debug for WorkbookQuery {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        f.debug_tuple("WorkbookQuery").field(&self.inner).finish()
    }
}
impl Clone for WorkbookQuery {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl WorkbookQuery {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "WorkbookQuery",
            },
        }
    }
    /// Returns the Power Query name.
    pub fn name(&self) -> Result<String, ExcelComError> {
        string(&self.inner, "excel.workbookquery.name")
    }
    /// Returns Excel's non-secret query description.
    pub fn description(&self) -> Result<String, ExcelComError> {
        string(&self.inner, "excel.workbookquery.description")
    }
    /// Returns the M formula in a redacting wrapper.
    pub fn formula(&self) -> Result<SecretStringValue, ExcelComError> {
        SecretStringValue::new(string(&self.inner, "excel.workbookquery.formula")?)
    }
    /// Requests an Excel-owned query refresh; this crate does not edit M formulas.
    pub fn refresh(&self) -> Result<(), ExcelComError> {
        let _ = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.workbookquery.refresh"), false),
            vec![],
            false,
        )?;
        Ok(())
    }
    /// Deletes the query and consumes its invalidated wrapper.
    pub fn delete(self) -> Result<(), ExcelComError> {
        let _ = invoke(
            &self.inner.dispatch,
            member(MemberId::new("excel.workbookquery.delete"), false),
            vec![],
            false,
        )?;
        Ok(())
    }
}

impl Workbook {
    /// Returns workbook Power Queries when the installed Excel version exposes `Workbook.Queries`.
    pub fn queries(&self) -> Result<WorkbookQueries, ExcelComError> {
        object(
            self.dispatch_object(),
            "excel.workbook.queries",
            WorkbookQueries::from_dispatch,
        )
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn formula_debug_never_exposes_source() {
        let formula = SecretStringValue::new("let Source = \"secret\"").expect("secret");
        assert!(!format!("{formula:?}").contains("secret"));
    }
}
