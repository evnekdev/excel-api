//! Worksheet presentation and output APIs, organized by Excel feature area.
//!
//! Public types remain re-exported from this façade so the internal split does
//! not change the crate's public object-model surface.

use std::fmt::{Debug, Formatter};
use std::iter::FusedIterator;
use std::path::Path;

use windows_sys::Win32::System::Variant::VT_DISPATCH;

use crate::automation::{
    AutomationArgument, AutomationValue, ConversionPolicy, OwnedVariant, PositionalArguments,
    decode_variant, invoke, property_get, property_put,
};
use crate::excel::collection::{
    CollectionDescriptor, count as collection_count, enumerator, item_by_index,
};
use crate::excel::formatting::{map_mixed, mixed_bool, mixed_i32, property_mixed_get};
use crate::excel::text::text_bstr;
use crate::excel::{
    Application, DispatchObject, ExcelColor, MixedValue, Range, ReferenceStyle, Workbook,
    WorkbookOpenOptions, Workbooks, Worksheet,
};
use crate::internal::{ComPtr, Dispatch, path_bstr};
use crate::object_model::{MemberId, member};
use crate::{ConversionError, ExcelComError};

mod helpers;
use helpers::*;
mod types;
pub use types::*;
mod sheets;
pub use sheets::*;
mod window;
pub use window::*;
mod layout;
pub use layout::*;
mod output;
pub use output::*;
mod macro_runtime;
pub use macro_runtime::*;
mod conditional;
mod operations;
pub use conditional::*;
mod styles;
pub use styles::*;
mod comments;
pub use comments::*;
mod hyperlinks;
pub use hyperlinks::*;

#[cfg(test)]
mod tests;
