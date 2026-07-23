//! Private common PivotTable helpers.

use crate::automation::{
    EnumVariant, OwnedVariant, enumerated_dispatch, invoke, property_get, property_put,
};
use crate::excel::collection::{CollectionDescriptor, count, enumerator, item_by_index};
use crate::excel::{DispatchObject, Range, RangeAddressOptions, ReferenceStyle};
use crate::internal::{ComPtr, Dispatch};
use crate::object_model::{MemberId, member};
use crate::{ConversionError, ExcelComError};
use std::iter::FusedIterator;

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
pub(crate) fn text(target: &DispatchObject, id: &'static str) -> Result<String, ExcelComError> {
    property_get(&target.dispatch, member(MemberId::new(id), false), vec![])?.as_string()
}
pub(crate) fn boolean(target: &DispatchObject, id: &'static str) -> Result<bool, ExcelComError> {
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
pub(crate) fn put(
    target: &DispatchObject,
    id: &'static str,
    value: OwnedVariant,
) -> Result<(), ExcelComError> {
    let _ = property_put(&target.dispatch, member(MemberId::new(id), true), value)?;
    Ok(())
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
pub(crate) fn source_reference(range: &Range) -> Result<String, ExcelComError> {
    range.address_with_options(&RangeAddressOptions {
        row_absolute: Some(true),
        column_absolute: Some(true),
        reference_style: ReferenceStyle::R1C1,
        external: Some(true),
        relative_to: None,
    })
}
pub(crate) fn count_collection(
    target: &DispatchObject,
    descriptor: CollectionDescriptor,
) -> Result<usize, ExcelComError> {
    count(target, descriptor)
}
pub(crate) fn item<T>(
    target: &DispatchObject,
    descriptor: CollectionDescriptor,
    index: usize,
    make: impl FnOnce(ComPtr<Dispatch>) -> T,
) -> Result<T, ExcelComError> {
    Ok(make(item_by_index(target, descriptor, index)?))
}
pub(crate) struct Iter<T> {
    pub(crate) enumerator: EnumVariant,
    pub(crate) index: usize,
    pub(crate) terminal: bool,
    pub(crate) kind: &'static str,
    pub(crate) make: fn(ComPtr<Dispatch>) -> T,
}
impl<T> Iterator for Iter<T> {
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
impl<T> FusedIterator for Iter<T> {}
pub(crate) fn iter<T>(
    target: &DispatchObject,
    descriptor: CollectionDescriptor,
    kind: &'static str,
    make: fn(ComPtr<Dispatch>) -> T,
) -> Result<Iter<T>, ExcelComError> {
    Ok(Iter {
        enumerator: enumerator(target, descriptor)?,
        index: 0,
        terminal: false,
        kind,
        make,
    })
}
pub(crate) fn method_iter<T>(
    target: &DispatchObject,
    id: &'static str,
    kind: &'static str,
    make: fn(ComPtr<Dispatch>) -> T,
) -> Result<Iter<T>, ExcelComError> {
    let mut value = invoke(
        &target.dispatch,
        member(MemberId::new(id), false),
        vec![],
        false,
    )?;
    Ok(Iter {
        enumerator: EnumVariant::from_variant(&mut value, kind)?,
        index: 0,
        terminal: false,
        kind,
        make,
    })
}
