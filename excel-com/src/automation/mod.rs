mod argument;
mod array;
mod bstr;
mod conversion_error;
mod currency;
mod date;
mod dispatch;
mod enumerator;
mod excel_error;
mod invocation;
mod message_filter;
mod policy;
mod retry;
mod safearray;
mod value;
mod variant;

pub use argument::AutomationArgument;
pub(crate) use argument::{PositionalArguments, reverse_for_com};
pub use array::AutomationArray;
pub use conversion_error::ConversionError;
pub use currency::Currency;
pub use date::OaDate;
pub(crate) use dispatch::{activate_excel, invoke, property_get, property_put};
pub(crate) use enumerator::{EnumVariant, enumerated_dispatch};
pub use excel_error::ExcelError;
pub(crate) use invocation::{MemberDescriptor, MemberKind};
pub(crate) use message_filter::ComMessageFilterGuard;
pub use policy::ConversionPolicy;
pub use retry::ComRetryPolicy;
pub(crate) use retry::{
    ComCallDisposition, InvocationRetrySafety, active_policy, classify_com_hresult,
};
pub(crate) use safearray::SafeArray;
pub use value::AutomationValue;
pub(crate) use value::{DateWriteMode, decode_variant, encode_variant, validate_range_shape};
pub(crate) use variant::OwnedVariant;
