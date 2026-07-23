//! Private, shared external-data COM helpers.

use std::iter::FusedIterator;

use crate::automation::{EnumVariant, enumerated_dispatch, property_get};
use crate::excel::collection::{
    CollectionDescriptor, count, enumerator, item_by_index, item_by_name,
};
use crate::excel::{DispatchObject, Range};
use crate::internal::{ComPtr, Dispatch};
use crate::object_model::{MemberId, member};
use crate::{ConversionError, ExcelComError};

pub(crate) fn object<T>(
    target: &DispatchObject,
    id: &'static str,
    make: impl FnOnce(ComPtr<Dispatch>) -> T,
) -> Result<T, ExcelComError> {
    let mut value = property_get(&target.dispatch, member(MemberId::new(id), false), vec![])?;
    Ok(make(value.take_dispatch()?))
}
pub(crate) fn optional_object<T>(
    target: &DispatchObject,
    id: &'static str,
    make: impl FnOnce(ComPtr<Dispatch>) -> T,
) -> Result<Option<T>, ExcelComError> {
    let mut value = property_get(&target.dispatch, member(MemberId::new(id), false), vec![])?;
    Ok(value.take_optional_dispatch()?.map(make))
}
pub(crate) fn string(target: &DispatchObject, id: &'static str) -> Result<String, ExcelComError> {
    property_get(&target.dispatch, member(MemberId::new(id), false), vec![])?.as_string()
}
pub(crate) fn bool_value(target: &DispatchObject, id: &'static str) -> Result<bool, ExcelComError> {
    let value = property_get(&target.dispatch, member(MemberId::new(id), false), vec![])?;
    value.as_bool().ok_or(ExcelComError::Conversion(
        ConversionError::UnsupportedVariantType {
            vartype: value.vt(),
        },
    ))
}
pub(crate) fn integer(target: &DispatchObject, id: &'static str) -> Result<i32, ExcelComError> {
    let value = property_get(&target.dispatch, member(MemberId::new(id), false), vec![])?;
    value.as_i32().ok_or(ExcelComError::Conversion(
        ConversionError::UnsupportedVariantType {
            vartype: value.vt(),
        },
    ))
}
pub(crate) fn range(target: &DispatchObject, id: &'static str) -> Result<Range, ExcelComError> {
    object(target, id, Range::from_dispatch)
}
pub(crate) fn optional_range(
    target: &DispatchObject,
    id: &'static str,
) -> Result<Option<Range>, ExcelComError> {
    optional_object(target, id, Range::from_dispatch)
}
pub(crate) fn collection_count(
    target: &DispatchObject,
    descriptor: CollectionDescriptor,
) -> Result<usize, ExcelComError> {
    count(target, descriptor)
}
pub(crate) fn collection_index<T>(
    target: &DispatchObject,
    descriptor: CollectionDescriptor,
    index: usize,
    make: impl FnOnce(ComPtr<Dispatch>) -> T,
) -> Result<T, ExcelComError> {
    Ok(make(item_by_index(target, descriptor, index)?))
}
pub(crate) fn collection_name<T>(
    target: &DispatchObject,
    descriptor: CollectionDescriptor,
    name: &str,
    make: impl FnOnce(ComPtr<Dispatch>) -> T,
) -> Result<T, ExcelComError> {
    Ok(make(item_by_name(target, descriptor, name)?))
}
pub(crate) fn collection_enum(
    target: &DispatchObject,
    descriptor: CollectionDescriptor,
) -> Result<EnumVariant, ExcelComError> {
    enumerator(target, descriptor)
}

pub(crate) struct TypedIter<T> {
    pub(crate) enumerator: EnumVariant,
    pub(crate) index: usize,
    pub(crate) terminal: bool,
    pub(crate) kind: &'static str,
    pub(crate) make: fn(ComPtr<Dispatch>) -> T,
}
impl<T> Iterator for TypedIter<T> {
    type Item = Result<T, ExcelComError>;
    fn next(&mut self) -> Option<Self::Item> {
        if self.terminal {
            return None;
        }
        match self.enumerator.next() {
            Ok(Some(mut value)) => {
                let index = self.index;
                self.index += 1;
                Some(enumerated_dispatch(&mut value, self.kind, index).map(self.make))
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
impl<T> FusedIterator for TypedIter<T> {}
