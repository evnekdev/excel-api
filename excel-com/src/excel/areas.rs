use std::fmt::{Debug, Formatter};
use std::iter::FusedIterator;

use crate::ExcelComError;
use crate::automation::{EnumVariant, enumerated_dispatch};
use crate::excel::collection::{
    CollectionDescriptor, count as collection_count, enumerator, item_by_index,
};
use crate::excel::{DispatchObject, Range};
use crate::internal::{ComPtr, Dispatch};
use crate::object_model::MemberId;

const DESCRIPTOR: CollectionDescriptor = CollectionDescriptor {
    name: "Areas",
    count: MemberId::new("excel.areas.count"),
    item: MemberId::new("excel.areas.item"),
    new_enum: MemberId::new("excel.areas.newenum"),
};

/// A typed collection of contiguous ranges in a multi-area Excel range.
pub struct Areas {
    inner: DispatchObject,
}
impl Debug for Areas {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        formatter.debug_tuple("Areas").field(&self.inner).finish()
    }
}
impl Clone for Areas {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
        }
    }
}
impl Areas {
    pub(crate) fn from_dispatch(dispatch: ComPtr<Dispatch>) -> Self {
        Self {
            inner: DispatchObject {
                dispatch,
                kind: "Areas",
            },
        }
    }
    /// Returns the number of contiguous areas.
    pub fn count(&self) -> Result<usize, ExcelComError> {
        collection_count(&self.inner, DESCRIPTOR)
    }
    /// Returns the one-based contiguous area at `index`.
    pub fn item(&self, index: usize) -> Result<Range, ExcelComError> {
        Ok(Range::from_dispatch(item_by_index(
            &self.inner,
            DESCRIPTOR,
            index,
        )?))
    }
    /// Iterates ranges in Excel's `_NewEnum` area order.
    pub fn iter(&self) -> Result<AreasIter, ExcelComError> {
        Ok(AreasIter {
            enumerator: enumerator(&self.inner, DESCRIPTOR)?,
            next_index: 0,
            terminal: false,
        })
    }
}

/// Typed, single-pass Areas iterator.
pub struct AreasIter {
    enumerator: EnumVariant,
    next_index: usize,
    terminal: bool,
}
impl Iterator for AreasIter {
    type Item = Result<Range, ExcelComError>;
    fn next(&mut self) -> Option<Self::Item> {
        if self.terminal {
            return None;
        }
        match self.enumerator.next() {
            Ok(Some(mut value)) => {
                let index = self.next_index;
                self.next_index += 1;
                Some(enumerated_dispatch(&mut value, "Areas", index).map(Range::from_dispatch))
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
impl FusedIterator for AreasIter {}
