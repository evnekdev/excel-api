//! Focused drawing implementation component.
#![allow(missing_docs)]
#[allow(unused_imports)]
use super::helpers::*;
#[allow(unused_imports)]
use super::types::*;
#[allow(unused_imports)]
use super::*;
pub(super) fn copy_picture(
    target: &DispatchObject,
    id: &'static str,
    options: &CopyPictureOptions,
) -> Result<(), ExcelComError> {
    let mut args = PositionalArguments::new();
    args.push_required(OwnedVariant::i32(options.appearance.raw()));
    args.push_required(OwnedVariant::i32(options.format.raw()));
    call(target, id, args.into_inner())
}
