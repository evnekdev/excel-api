use crate::automation::{EnumVariant, OwnedVariant, property_get};
use crate::excel::DispatchObject;
use crate::object_model::{MemberId, member};
use crate::{ConversionError, ExcelComError};

/// Private descriptor shared by the small, typed collection wrappers.
pub(crate) struct CollectionDescriptor {
    pub(crate) name: &'static str,
    pub(crate) count: MemberId,
    pub(crate) item: MemberId,
    pub(crate) new_enum: MemberId,
}

pub(crate) fn count(
    target: &DispatchObject,
    descriptor: CollectionDescriptor,
) -> Result<usize, ExcelComError> {
    let value = property_get(&target.dispatch, member(descriptor.count, false), vec![])?;
    let count = value.as_i32().ok_or(ExcelComError::Conversion(
        ConversionError::UnsupportedVariantType {
            vartype: value.vt(),
        },
    ))?;
    usize::try_from(count).map_err(|_| ExcelComError::Unsupported {
        detail: "collection Count was negative",
    })
}

pub(crate) fn item_by_index(
    target: &DispatchObject,
    descriptor: CollectionDescriptor,
    index: usize,
) -> Result<crate::internal::ComPtr<crate::internal::Dispatch>, ExcelComError> {
    if index == 0 {
        return Err(ExcelComError::Unsupported {
            detail: "collection index is one-based",
        });
    }
    let index = i32::try_from(index).map_err(|_| ExcelComError::Unsupported {
        detail: "collection index exceeds i32",
    })?;
    let mut result = property_get(
        &target.dispatch,
        member(descriptor.item, false),
        vec![OwnedVariant::i32(index)],
    )?;
    result.take_dispatch()
}

pub(crate) fn item_by_name(
    target: &DispatchObject,
    descriptor: CollectionDescriptor,
    name: &str,
) -> Result<crate::internal::ComPtr<crate::internal::Dispatch>, ExcelComError> {
    let mut result = property_get(
        &target.dispatch,
        member(descriptor.item, false),
        vec![OwnedVariant::bstr(name)?],
    )?;
    result.take_dispatch()
}

pub(crate) fn enumerator(
    target: &DispatchObject,
    descriptor: CollectionDescriptor,
) -> Result<EnumVariant, ExcelComError> {
    EnumVariant::from_new_enum(target, descriptor.new_enum, descriptor.name)
}
