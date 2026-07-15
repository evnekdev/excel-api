#![doc = "Procedural macro entry points for `excel-api`."]

use std::collections::BTreeMap;

use proc_macro::TokenStream;
use proc_macro2::Span;
use quote::{format_ident, quote};
use syn::{
    Expr, FnArg, GenericArgument, Ident, ItemFn, Lit, LitStr, Meta, MetaNameValue, Pat,
    PathArguments, ReturnType, Token, Type, parse_macro_input, punctuated::Punctuated,
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

enum InvocationArgument {
    Excel(Ident),
    Context,
}

struct ThunkArgument {
    ident: Ident,
    ty: Type,
    kind: ArgumentKind,
}

/// Generate deterministic worksheet-function registration metadata and an ABI thunk.
///
/// The ordinary function is preserved. The generated metadata constant is
/// named `__EXCEL_FUNCTION_METADATA_<FUNCTION_NAME_IN_UPPERCASE>` and the
/// Rust-visible thunk item is named
/// `__excel_function_thunk_<function_name_in_lowercase>`.
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
    validate_export_name(&thunk)?;

    let mut excel_arguments = Vec::new();
    let mut thunk_arguments = Vec::new();
    let mut invocation_arguments = Vec::new();
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
            invocation_arguments.push(InvocationArgument::Context);
            continue;
        }

        let kind = map_argument_type(&argument.ty)?;
        excel_arguments.push(kind);
        let name = pattern.ident.to_string();
        let help = attributes.arguments.get(&name).ok_or_else(|| {
            syn::Error::new(
                pattern.ident.span(),
                format!("missing `arguments({name} = \"help\")` entry"),
            )
        })?;
        argument_names.push(LitStr::new(&name, pattern.ident.span()));
        argument_help.push(help.clone());
        thunk_arguments.push(ThunkArgument {
            ident: pattern.ident.clone(),
            ty: (*argument.ty).clone(),
            kind,
        });
        invocation_arguments.push(InvocationArgument::Excel(pattern.ident.clone()));
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
        Some(value) => {
            let result = parse_result_override(&value)?;
            if result != inferred_result && result != ResultKind::Xloper12 {
                return Err(syn::Error::new_spanned(
                    value,
                    "a scalar return_type override must match the Rust result; use `xloper12` for a general Excel return",
                ));
            }
            result
        }
        None => inferred_result,
    };

    let function_ident = &function.sig.ident;
    let metadata_ident = format_ident!(
        "__EXCEL_FUNCTION_METADATA_{}",
        function_ident.to_string().to_uppercase(),
        span = function_ident.span()
    );
    let thunk_ident = format_ident!(
        "__excel_function_thunk_{}",
        function_ident.to_string().to_lowercase(),
        span = function_ident.span()
    );
    let result_metadata = result_tokens(result);
    let arguments = excel_arguments.iter().copied().map(argument_tokens);
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
    let raw_arguments = thunk_arguments.iter().map(raw_argument_tokens);
    let conversions = thunk_arguments.iter().map(conversion_tokens);
    let invoke_arguments = invocation_arguments.iter().map(|argument| match argument {
        InvocationArgument::Excel(ident) => quote!(#ident),
        InvocationArgument::Context => quote!(__excel_context),
    });
    let invoke = quote!(#function_ident(#(#invoke_arguments),*));
    let invoke = match context {
        Some(ContextKind::Worksheet) => {
            quote!(__excel_scope.with_worksheet_context(|__excel_context| #invoke))
        }
        Some(ContextKind::ThreadSafe) => {
            quote!(__excel_scope.with_thread_safe_context(|__excel_context| #invoke))
        }
        Some(ContextKind::Macro) => {
            quote!(__excel_scope.with_macro_context(|__excel_context| #invoke))
        }
        None => invoke,
    };
    let unwrap_result = if result_error_type(&function.sig.output).is_some() {
        quote! {
            let __excel_result = __excel_result
                .map_err(::excel_api::thunk::function_error)?;
        }
    } else {
        proc_macro2::TokenStream::new()
    };
    let converted_result = match result {
        ResultKind::Number => quote!(Ok(__excel_result)),
        ResultKind::Boolean => quote!(Ok(if __excel_result { 1_i16 } else { 0_i16 })),
        ResultKind::Integer => quote!(Ok(i32::from(__excel_result))),
        ResultKind::Xloper12 => quote!(::excel_api::thunk::IntoThunkReturn::into_thunk_return(
            __excel_result
        )),
    };
    let thunk_body = quote! {
        ::excel_api::thunk::with_callback(|__excel_scope| {
            #(#conversions)*
            let __excel_result = #invoke;
            #unwrap_result
            #converted_result
        })
    };
    let thunk_return = match result {
        ResultKind::Number => quote!(f64),
        ResultKind::Boolean => quote!(i16),
        ResultKind::Integer => quote!(i32),
        ResultKind::Xloper12 => quote!(::excel_api::thunk::RawXloper12),
    };
    let thunk_execution = match result {
        ResultKind::Number => quote!(::excel_api::thunk::scalar_thunk(0.0_f64, || #thunk_body)),
        ResultKind::Boolean => quote!(::excel_api::thunk::scalar_thunk(0_i16, || #thunk_body)),
        ResultKind::Integer => quote!(::excel_api::thunk::scalar_thunk(0_i32, || #thunk_body)),
        ResultKind::Xloper12 => quote!(::excel_api::thunk::xloper12_thunk(|| #thunk_body)),
    };

    Ok(quote! {
        #function

        #[doc(hidden)]
        #[doc = "# Safety"]
        #[doc = "Pointer arguments must be the live callback-owned values described by the generated registration signature."]
        #[unsafe(export_name = #thunk)]
        pub unsafe extern "system" fn #thunk_ident(#(#raw_arguments),*) -> #thunk_return {
            #thunk_execution
        }

        #[doc(hidden)]
        #[allow(non_upper_case_globals)]
        pub const #metadata_ident: ::excel_api::FunctionRegistration =
            ::excel_api::FunctionRegistration::new(
                #thunk,
                #excel_name,
                ::excel_api::FunctionSignature::new(#result_metadata, &[#(#arguments),*]),
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

fn raw_argument_tokens(argument: &ThunkArgument) -> proc_macro2::TokenStream {
    let ident = &argument.ident;
    let ty = match argument.kind {
        ArgumentKind::Number => quote!(f64),
        ArgumentKind::Boolean => quote!(i16),
        ArgumentKind::Integer => quote!(i32),
        ArgumentKind::GeneralValue | ArgumentKind::GeneralReference => {
            quote!(::excel_api::thunk::RawXloper12)
        }
        ArgumentKind::CountedUtf16 | ArgumentKind::NullTerminatedUtf16 => {
            quote!(*mut ::excel_api::thunk::RawXchar)
        }
    };
    quote!(#ident: #ty)
}

fn conversion_tokens(argument: &ThunkArgument) -> proc_macro2::TokenStream {
    let ident = &argument.ident;
    let ty = &argument.ty;
    match argument.kind {
        ArgumentKind::Number => quote!(let #ident: #ty = #ident;),
        ArgumentKind::Boolean => quote!(let #ident: #ty = #ident != 0;),
        ArgumentKind::Integer => quote! {
            let #ident: #ty = ::excel_api::thunk::from_excel(
                ::excel_api::ExcelValueRef::Integer(#ident)
            )?;
        },
        ArgumentKind::GeneralValue | ArgumentKind::GeneralReference => quote! {
            // SAFETY: forwarded from the generated exported thunk contract.
            let #ident = unsafe { __excel_scope.decode(#ident) }?;
            let #ident: #ty = ::excel_api::thunk::from_excel(#ident)?;
        },
        ArgumentKind::CountedUtf16 => quote! {
            // SAFETY: forwarded from the generated exported thunk contract.
            let #ident: #ty = unsafe { __excel_scope.counted_utf16(#ident) }?;
        },
        ArgumentKind::NullTerminatedUtf16 => quote! {
            // SAFETY: forwarded from the generated exported thunk contract.
            let #ident: #ty = unsafe { __excel_scope.null_terminated_utf16(#ident) }?;
        },
    }
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

fn validate_export_name(value: &LitStr) -> syn::Result<()> {
    let name = value.value();
    let mut characters = name.chars();
    let valid_start = characters
        .next()
        .is_some_and(|character| character == '_' || character.is_ascii_alphabetic());
    if !valid_start
        || !characters.all(|character| character == '_' || character.is_ascii_alphanumeric())
    {
        return Err(syn::Error::new_spanned(
            value,
            "thunk must be a non-empty ASCII x64 export identifier",
        ));
    }
    Ok(())
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

fn result_error_type(output: &ReturnType) -> Option<&Type> {
    let ReturnType::Type(_, ty) = output else {
        return None;
    };
    let Type::Path(path) = ty.as_ref() else {
        return None;
    };
    let segment = path.path.segments.last()?;
    (segment.ident == "Result")
        .then(|| two_type_arguments(segment, ty).ok().map(|(_, error)| error))
        .flatten()
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
    fn rejects_invalid_exports_and_incompatible_scalar_overrides() {
        let function: ItemFn = parse_quote!(
            fn value() -> f64 {
                1.0
            }
        );
        let invalid_export: Punctuated<Meta, Token![,]> =
            parse_quote!(name = "VALUE", thunk = "not-an-export");
        assert!(
            expand_excel_function(invalid_export, function.clone())
                .unwrap_err()
                .to_string()
                .contains("ASCII x64 export")
        );

        let invalid_override: Punctuated<Meta, Token![,]> =
            parse_quote!(name = "VALUE", thunk = "value", return_type = "boolean");
        assert!(
            expand_excel_function(invalid_override, function)
                .unwrap_err()
                .to_string()
                .contains("must match")
        );
    }

    #[test]
    fn deterministic_metadata_and_exact_xloper12_thunk_are_generated() {
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
        assert!(expanded.contains("__excel_function_thunk_add"));
        assert!(expanded.contains("FunctionRegistration"));
        assert!(expanded.contains("export_name = \"rust_add\""));
        assert!(expanded.contains("unsafe extern \"system\""));
        assert!(expanded.contains("f64 , y : f64"));
        assert!(expanded.contains("thunk :: RawXloper12"));
        assert!(expanded.contains("thunk :: xloper12_thunk"));
        assert!(expanded.contains("thunk :: with_callback"));
    }

    #[test]
    fn scalar_and_pointer_abi_signatures_come_from_the_registration_model() {
        let function: ItemFn = parse_quote!(
            fn probe(flag: bool, integer: i16, value: ExcelValueArg<'_>) -> bool {
                let _ = (integer, value);
                flag
            }
        );
        let attributes: Punctuated<Meta, Token![,]> = parse_quote!(
            name = "PROBE",
            thunk = "probe_export",
            arguments(flag = "flag", integer = "integer", value = "value")
        );
        let expanded = expand_excel_function(attributes, function)
            .unwrap()
            .to_string();
        assert!(expanded.contains("flag : i16"));
        assert!(expanded.contains("integer : i32"));
        assert!(expanded.contains("value : :: excel_api :: thunk :: RawXloper12"));
        assert!(expanded.contains("-> i16"));
        assert!(expanded.contains("ExcelArgumentType :: Boolean"));
        assert!(expanded.contains("ExcelArgumentType :: Integer"));
        assert!(expanded.contains("ExcelArgumentType :: GeneralValue"));
    }
}
