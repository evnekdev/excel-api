#![doc = "Procedural macro entry points for `excel-api`."]

use std::collections::BTreeMap;

use proc_macro::TokenStream;
use proc_macro2::Span;
use quote::{format_ident, quote};
use syn::{
    Expr, FnArg, GenericArgument, ItemFn, Lit, LitStr, Meta, MetaNameValue, Pat, PathArguments,
    ReturnType, Token, Type, parse_macro_input, punctuated::Punctuated,
};

#[derive(Default)]
struct FunctionAttributes {
    name: Option<LitStr>,
    category: Option<LitStr>,
    description: Option<LitStr>,
    thunk: Option<LitStr>,
    return_type: Option<LitStr>,
    volatile: bool,
    thread_safe: bool,
    macro_type: bool,
    cluster_safe: bool,
    arguments: BTreeMap<String, LitStr>,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum ArgumentKind {
    Number,
    Boolean,
    Integer,
    GeneralValue,
    GeneralReference,
    CountedUtf16,
    NullTerminatedUtf16,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum ResultKind {
    Number,
    Boolean,
    Integer,
    Xloper12,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum ContextKind {
    Worksheet,
    ThreadSafe,
    Macro,
}

/// Generate deterministic worksheet-function registration metadata.
///
/// This milestone deliberately preserves the annotated Rust function and does
/// not generate an exported ABI thunk. The generated constant is named
/// `__EXCEL_FUNCTION_METADATA_<FUNCTION_NAME_IN_UPPERCASE>`.
#[proc_macro_attribute]
pub fn excel_function(attribute: TokenStream, item: TokenStream) -> TokenStream {
    let attributes = parse_macro_input!(
        attribute with Punctuated::<Meta, Token![,]>::parse_terminated
    );
    let function = parse_macro_input!(item as ItemFn);
    match expand_excel_function(attributes, function) {
        Ok(tokens) => tokens.into(),
        Err(error) => error.into_compile_error().into(),
    }
}

fn expand_excel_function(
    raw_attributes: Punctuated<Meta, Token![,]>,
    function: ItemFn,
) -> syn::Result<proc_macro2::TokenStream> {
    let attributes = parse_attributes(raw_attributes)?;
    validate_function_shape(&function)?;

    let excel_name = attributes.name.ok_or_else(|| {
        syn::Error::new(
            function.sig.ident.span(),
            "missing required `name = \"...\"`",
        )
    })?;
    let thunk = attributes.thunk.ok_or_else(|| {
        syn::Error::new(
            function.sig.ident.span(),
            "missing required future `thunk = \"...\"` symbol association",
        )
    })?;

    let mut excel_arguments = Vec::new();
    let mut argument_names = Vec::new();
    let mut argument_help = Vec::new();
    let mut context = None;

    for input in &function.sig.inputs {
        let FnArg::Typed(argument) = input else {
            return Err(syn::Error::new_spanned(input, "methods are not supported"));
        };
        let Pat::Ident(pattern) = argument.pat.as_ref() else {
            return Err(syn::Error::new_spanned(
                &argument.pat,
                "destructuring arguments are not supported",
            ));
        };
        if pattern.by_ref.is_some() || pattern.subpat.is_some() {
            return Err(syn::Error::new_spanned(
                &argument.pat,
                "argument patterns must be plain identifiers",
            ));
        }

        if let Some(kind) = context_kind(&argument.ty) {
            if context.replace(kind).is_some() {
                return Err(syn::Error::new_spanned(
                    &argument.ty,
                    "only one injected context is supported",
                ));
            }
            continue;
        }

        excel_arguments.push(map_argument_type(&argument.ty)?);
        let name = pattern.ident.to_string();
        let help = attributes.arguments.get(&name).ok_or_else(|| {
            syn::Error::new(
                pattern.ident.span(),
                format!("missing `arguments({name} = \"help\")` entry"),
            )
        })?;
        argument_names.push(LitStr::new(&name, pattern.ident.span()));
        argument_help.push(help.clone());
    }

    if attributes.arguments.len() != argument_names.len() {
        return Err(syn::Error::new(
            Span::call_site(),
            "argument help must name every Excel-visible argument exactly once",
        ));
    }
    if matches!(context, Some(ContextKind::ThreadSafe)) && !attributes.thread_safe {
        return Err(syn::Error::new(
            function.sig.ident.span(),
            "`ThreadSafeContext` requires the `thread_safe` flag",
        ));
    }
    if matches!(context, Some(ContextKind::Worksheet)) && attributes.thread_safe {
        return Err(syn::Error::new(
            function.sig.ident.span(),
            "thread-safe functions must inject `ThreadSafeContext`, not `WorksheetContext`",
        ));
    }
    if matches!(context, Some(ContextKind::Macro)) && !attributes.macro_type {
        return Err(syn::Error::new(
            function.sig.ident.span(),
            "`MacroContext` requires the `macro_type` flag",
        ));
    }
    if attributes.macro_type && (attributes.thread_safe || attributes.cluster_safe) {
        return Err(syn::Error::new(
            function.sig.ident.span(),
            "macro-sheet functions cannot be thread-safe or cluster-safe",
        ));
    }

    let inferred_result = map_result_type(&function.sig.output)?;
    let result = match attributes.return_type {
        Some(value) => parse_result_override(&value)?,
        None => inferred_result,
    };

    let function_ident = &function.sig.ident;
    let metadata_ident = format_ident!(
        "__EXCEL_FUNCTION_METADATA_{}",
        function_ident.to_string().to_uppercase(),
        span = function_ident.span()
    );
    let result = result_tokens(result);
    let arguments = excel_arguments.into_iter().map(argument_tokens);
    let category = attributes
        .category
        .map(|value| quote!(.category(#value)))
        .unwrap_or_default();
    let description = attributes
        .description
        .map(|value| quote!(.description(#value)))
        .unwrap_or_default();
    let volatile = attributes.volatile;
    let thread_safe = attributes.thread_safe;
    let macro_type = attributes.macro_type;
    let cluster_safe = attributes.cluster_safe;

    Ok(quote! {
        #function

        #[doc(hidden)]
        #[allow(non_upper_case_globals)]
        pub const #metadata_ident: ::excel_api::FunctionRegistration =
            ::excel_api::FunctionRegistration::new(
                #thunk,
                #excel_name,
                ::excel_api::FunctionSignature::new(#result, &[#(#arguments),*]),
            )
            #category
            #description
            .arguments(&[#(#argument_names),*], &[#(#argument_help),*])
            .flags(::excel_api::FunctionFlags {
                volatile: #volatile,
                thread_safe: #thread_safe,
                macro_type: #macro_type,
                cluster_safe: #cluster_safe,
            });
    })
}

fn parse_attributes(raw: Punctuated<Meta, Token![,]>) -> syn::Result<FunctionAttributes> {
    let mut parsed = FunctionAttributes::default();
    for meta in raw {
        match meta {
            Meta::NameValue(value) if value.path.is_ident("name") => {
                set_once(&mut parsed.name, string_value(&value)?, &value)?;
            }
            Meta::NameValue(value) if value.path.is_ident("category") => {
                set_once(&mut parsed.category, string_value(&value)?, &value)?;
            }
            Meta::NameValue(value) if value.path.is_ident("description") => {
                set_once(&mut parsed.description, string_value(&value)?, &value)?;
            }
            Meta::NameValue(value) if value.path.is_ident("thunk") => {
                set_once(&mut parsed.thunk, string_value(&value)?, &value)?;
            }
            Meta::NameValue(value) if value.path.is_ident("return_type") => {
                set_once(&mut parsed.return_type, string_value(&value)?, &value)?;
            }
            Meta::Path(path) if path.is_ident("volatile") => set_flag(&mut parsed.volatile, path)?,
            Meta::Path(path) if path.is_ident("thread_safe") => {
                set_flag(&mut parsed.thread_safe, path)?;
            }
            Meta::Path(path) if path.is_ident("macro_type") => {
                set_flag(&mut parsed.macro_type, path)?;
            }
            Meta::Path(path) if path.is_ident("cluster_safe") => {
                set_flag(&mut parsed.cluster_safe, path)?;
            }
            Meta::List(list) if list.path.is_ident("arguments") => {
                let values =
                    list.parse_args_with(Punctuated::<MetaNameValue, Token![,]>::parse_terminated)?;
                for value in values {
                    let Some(ident) = value.path.get_ident() else {
                        return Err(syn::Error::new_spanned(
                            value.path,
                            "argument help keys must be identifiers",
                        ));
                    };
                    let name = ident.to_string();
                    let help = string_value(&value)?;
                    if parsed.arguments.insert(name.clone(), help).is_some() {
                        return Err(syn::Error::new_spanned(
                            ident,
                            format!("duplicate help for argument `{name}`"),
                        ));
                    }
                }
            }
            unsupported => {
                return Err(syn::Error::new_spanned(
                    unsupported,
                    "unsupported `excel_function` attribute option",
                ));
            }
        }
    }
    Ok(parsed)
}

fn set_once<T>(slot: &mut Option<T>, value: T, source: &MetaNameValue) -> syn::Result<()> {
    if slot.replace(value).is_some() {
        Err(syn::Error::new_spanned(
            source,
            "duplicate attribute option",
        ))
    } else {
        Ok(())
    }
}

fn set_flag(slot: &mut bool, path: syn::Path) -> syn::Result<()> {
    if std::mem::replace(slot, true) {
        Err(syn::Error::new_spanned(path, "duplicate flag"))
    } else {
        Ok(())
    }
}

fn string_value(value: &MetaNameValue) -> syn::Result<LitStr> {
    if let Expr::Lit(expression) = &value.value
        && let Lit::Str(string) = &expression.lit
    {
        return Ok(string.clone());
    }
    Err(syn::Error::new_spanned(
        &value.value,
        "expected a string literal",
    ))
}

fn validate_function_shape(function: &ItemFn) -> syn::Result<()> {
    let signature = &function.sig;
    if signature
        .inputs
        .iter()
        .any(|input| matches!(input, FnArg::Receiver(_)))
    {
        return Err(syn::Error::new_spanned(
            signature,
            "methods are not supported",
        ));
    }
    if signature.variadic.is_some() {
        return Err(syn::Error::new_spanned(
            signature,
            "variadic functions are not supported",
        ));
    }
    if signature.asyncness.is_some() {
        return Err(syn::Error::new_spanned(
            signature,
            "async functions are not supported",
        ));
    }
    if signature.unsafety.is_some() {
        return Err(syn::Error::new_spanned(
            signature,
            "unsafe functions are not supported",
        ));
    }
    if signature.abi.is_some() {
        return Err(syn::Error::new_spanned(
            signature,
            "ABI functions are not supported; M9A does not generate thunks",
        ));
    }
    if !signature.generics.params.is_empty() || signature.generics.where_clause.is_some() {
        return Err(syn::Error::new_spanned(
            signature,
            "generic functions are not supported",
        ));
    }
    Ok(())
}

fn context_kind(ty: &Type) -> Option<ContextKind> {
    let Type::Reference(reference) = ty else {
        return None;
    };
    if reference.mutability.is_some() {
        return None;
    }
    let Type::Path(path) = reference.elem.as_ref() else {
        return None;
    };
    match path.path.segments.last()?.ident.to_string().as_str() {
        "WorksheetContext" => Some(ContextKind::Worksheet),
        "ThreadSafeContext" => Some(ContextKind::ThreadSafe),
        "MacroContext" => Some(ContextKind::Macro),
        _ => None,
    }
}

fn map_argument_type(ty: &Type) -> syn::Result<ArgumentKind> {
    let Type::Path(path) = ty else {
        return Err(unsupported_type(ty));
    };
    let segment = path
        .path
        .segments
        .last()
        .ok_or_else(|| unsupported_type(ty))?;
    let name = segment.ident.to_string();
    match name.as_str() {
        "f64" => Ok(ArgumentKind::Number),
        "bool" => Ok(ArgumentKind::Boolean),
        "i16" | "i32" | "u16" => Ok(ArgumentKind::Integer),
        "ExcelString" | "String" | "ExcelArray" | "ExcelValue" | "ExcelValueArg" => {
            Ok(ArgumentKind::GeneralValue)
        }
        "ExcelReferenceArg" => Ok(ArgumentKind::GeneralReference),
        "CountedUtf16Arg" => Ok(ArgumentKind::CountedUtf16),
        "NullTerminatedUtf16Arg" => Ok(ArgumentKind::NullTerminatedUtf16),
        "ExcelValueRef" => Err(syn::Error::new_spanned(
            ty,
            "ambiguous `ExcelValueRef` input; use `ExcelValueArg` (Q) or `ExcelReferenceArg` (U)",
        )),
        "Option" | "OptionalValue" => {
            let inner = one_type_argument(segment, ty)?;
            match map_argument_type(inner)? {
                ArgumentKind::GeneralReference => Ok(ArgumentKind::GeneralReference),
                ArgumentKind::CountedUtf16 | ArgumentKind::NullTerminatedUtf16 => Err(
                    syn::Error::new_spanned(ty, "direct UTF-16 arguments cannot be optional"),
                ),
                _ => Ok(ArgumentKind::GeneralValue),
            }
        }
        _ => Err(unsupported_type(ty)),
    }
}

fn map_result_type(output: &ReturnType) -> syn::Result<ResultKind> {
    let ReturnType::Type(_, ty) = output else {
        return Err(syn::Error::new_spanned(
            output,
            "worksheet functions must return a supported value",
        ));
    };
    map_result_type_inner(ty)
}

fn map_result_type_inner(ty: &Type) -> syn::Result<ResultKind> {
    if matches!(ty, Type::Reference(_)) {
        return Err(syn::Error::new_spanned(
            ty,
            "borrowed returns are not supported",
        ));
    }
    let Type::Path(path) = ty else {
        return Err(unsupported_type(ty));
    };
    let segment = path
        .path
        .segments
        .last()
        .ok_or_else(|| unsupported_type(ty))?;
    match segment.ident.to_string().as_str() {
        "f64" => Ok(ResultKind::Number),
        "bool" => Ok(ResultKind::Boolean),
        "i16" | "i32" | "u16" => Ok(ResultKind::Integer),
        "ExcelString" | "String" | "ExcelArray" | "ExcelValue" | "ExcelReturnValue"
        | "ExcelError" => Ok(ResultKind::Xloper12),
        "Result" => {
            let (success, error) = two_type_arguments(segment, ty)?;
            validate_result_error(error)?;
            map_result_type_inner(success)
        }
        "ExcelStr" | "CountedUtf16Arg" | "NullTerminatedUtf16Arg" => Err(syn::Error::new_spanned(
            ty,
            "direct dynamic-string returns are not supported",
        )),
        _ => Err(unsupported_type(ty)),
    }
}

fn validate_result_error(ty: &Type) -> syn::Result<()> {
    let Type::Path(path) = ty else {
        return Err(syn::Error::new_spanned(
            ty,
            "unsupported Result error; use a documented excel-api error type",
        ));
    };
    let Some(segment) = path.path.segments.last() else {
        return Err(unsupported_type(ty));
    };
    match segment.ident.to_string().as_str() {
        "ExcelError"
        | "ConversionError"
        | "ReturnError"
        | "ReturnMaterializationError"
        | "ThunkError" => Ok(()),
        _ => Err(syn::Error::new_spanned(
            ty,
            "unsupported Result error; use ExcelError, ConversionError, ReturnError, ReturnMaterializationError, or ThunkError",
        )),
    }
}

fn one_type_argument<'a>(segment: &'a syn::PathSegment, ty: &Type) -> syn::Result<&'a Type> {
    let PathArguments::AngleBracketed(arguments) = &segment.arguments else {
        return Err(unsupported_type(ty));
    };
    let mut types = arguments.args.iter().filter_map(|argument| match argument {
        GenericArgument::Type(ty) => Some(ty),
        _ => None,
    });
    let first = types.next().ok_or_else(|| unsupported_type(ty))?;
    if types.next().is_some() {
        return Err(unsupported_type(ty));
    }
    Ok(first)
}

fn two_type_arguments<'a>(
    segment: &'a syn::PathSegment,
    ty: &Type,
) -> syn::Result<(&'a Type, &'a Type)> {
    let PathArguments::AngleBracketed(arguments) = &segment.arguments else {
        return Err(unsupported_type(ty));
    };
    let types = arguments
        .args
        .iter()
        .filter_map(|argument| match argument {
            GenericArgument::Type(ty) => Some(ty),
            _ => None,
        })
        .collect::<Vec<_>>();
    if types.len() != 2 {
        return Err(unsupported_type(ty));
    }
    Ok((types[0], types[1]))
}

fn parse_result_override(value: &LitStr) -> syn::Result<ResultKind> {
    match value.value().as_str() {
        "number" => Ok(ResultKind::Number),
        "boolean" => Ok(ResultKind::Boolean),
        "integer" => Ok(ResultKind::Integer),
        "xloper12" => Ok(ResultKind::Xloper12),
        _ => Err(syn::Error::new_spanned(
            value,
            "return_type must be `number`, `boolean`, `integer`, or `xloper12`",
        )),
    }
}

fn unsupported_type(ty: &Type) -> syn::Error {
    syn::Error::new_spanned(
        ty,
        "type is not supported by the closed Excel metadata mapping",
    )
}

fn argument_tokens(kind: ArgumentKind) -> proc_macro2::TokenStream {
    let variant = match kind {
        ArgumentKind::Number => quote!(Number),
        ArgumentKind::Boolean => quote!(Boolean),
        ArgumentKind::Integer => quote!(Integer),
        ArgumentKind::GeneralValue => quote!(GeneralValue),
        ArgumentKind::GeneralReference => quote!(GeneralReference),
        ArgumentKind::CountedUtf16 => quote!(CountedUtf16),
        ArgumentKind::NullTerminatedUtf16 => quote!(NullTerminatedUtf16),
    };
    quote!(::excel_api::ExcelArgumentType::#variant)
}

fn result_tokens(kind: ResultKind) -> proc_macro2::TokenStream {
    let variant = match kind {
        ResultKind::Number => quote!(Number),
        ResultKind::Boolean => quote!(Boolean),
        ResultKind::Integer => quote!(Integer),
        ResultKind::Xloper12 => quote!(Xloper12),
    };
    quote!(::excel_api::ExcelReturnType::#variant)
}

/// Mark a Rust function for future command registration.
///
/// Command metadata and thunks remain deferred to M12.
#[proc_macro_attribute]
pub fn excel_command(_attribute: TokenStream, item: TokenStream) -> TokenStream {
    item
}

#[cfg(test)]
mod tests {
    use super::*;
    use syn::parse_quote;

    fn shape_error(function: ItemFn) -> String {
        validate_function_shape(&function).unwrap_err().to_string()
    }

    #[test]
    fn rejects_unsupported_function_shapes() {
        assert!(
            shape_error(parse_quote!(
                async fn f() -> i32 {
                    0
                }
            ))
            .contains("async")
        );
        assert!(
            shape_error(parse_quote!(
                fn f<T>() -> i32 {
                    0
                }
            ))
            .contains("generic")
        );
        assert!(
            shape_error(parse_quote!(
                fn f(&self) -> i32 {
                    0
                }
            ))
            .contains("methods")
        );
        assert!(
            shape_error(parse_quote!(
                unsafe extern "C" fn f(_: i32, ...) -> i32 {
                    0
                }
            ))
            .contains("variadic")
        );
    }

    #[test]
    fn rejects_unsupported_and_ambiguous_types() {
        let impl_trait: Type = parse_quote!(impl Iterator<Item = i32>);
        let ambiguous: Type = parse_quote!(ExcelValueRef<'_>);
        let borrowed: ReturnType = parse_quote!(-> &str);
        let direct_string: ReturnType = parse_quote!(-> CountedUtf16Arg<'_>);
        assert!(map_argument_type(&impl_trait).is_err());
        assert!(
            map_argument_type(&ambiguous)
                .unwrap_err()
                .to_string()
                .contains("ambiguous")
        );
        assert!(
            map_result_type(&borrowed)
                .unwrap_err()
                .to_string()
                .contains("borrowed")
        );
        assert!(
            map_result_type(&direct_string)
                .unwrap_err()
                .to_string()
                .contains("dynamic-string")
        );
    }

    #[test]
    fn rejects_destructuring_and_unlisted_result_errors() {
        let function: ItemFn = parse_quote!(
            fn f((left, right): (i32, i32)) -> i32 {
                left + right
            }
        );
        let attributes: Punctuated<Meta, Token![,]> =
            parse_quote!(name = "F", thunk = "f_thunk", arguments(left = "left"));
        assert!(
            expand_excel_function(attributes, function)
                .unwrap_err()
                .to_string()
                .contains("destructuring")
        );

        let unsupported: ReturnType = parse_quote!(-> Result<i32, String>);
        assert!(map_result_type(&unsupported).is_err());
    }

    #[test]
    fn deterministic_metadata_name_and_no_thunk_are_generated() {
        let function: ItemFn = parse_quote!(
            fn add(x: f64, y: f64) -> f64 {
                x + y
            }
        );
        let attributes: Punctuated<Meta, Token![,]> = parse_quote!(
            name = "RUST.ADD",
            category = "Rust",
            description = "Adds",
            thunk = "rust_add",
            return_type = "xloper12",
            thread_safe,
            arguments(x = "x", y = "y")
        );
        let expanded = expand_excel_function(attributes, function)
            .unwrap()
            .to_string();
        assert!(expanded.contains("__EXCEL_FUNCTION_METADATA_ADD"));
        assert!(expanded.contains("FunctionRegistration"));
        assert!(!expanded.contains("no_mangle"));
        assert!(!expanded.contains("extern \"system\""));
        assert!(!expanded.contains("unsafe"));
    }
}
