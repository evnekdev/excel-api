#![doc = "Procedural macro entry points for `excel-api`."]

use proc_macro::TokenStream;

/// Mark a Rust function for future worksheet-function registration.
///
/// In this initial outline the attribute is intentionally transparent. Thunk
/// generation will be added only after the raw ABI and conversion model are
/// validated.
#[proc_macro_attribute]
pub fn excel_function(_attribute: TokenStream, item: TokenStream) -> TokenStream {
    item
}

/// Mark a Rust function for future command registration.
///
/// In this initial outline the attribute is intentionally transparent.
#[proc_macro_attribute]
pub fn excel_command(_attribute: TokenStream, item: TokenStream) -> TokenStream {
    item
}
