use core::fmt;

use excel_api_sys::{
    XLL_MODIFIER_CLUSTER_SAFE, XLL_MODIFIER_MACRO_SHEET, XLL_MODIFIER_THREAD_SAFE,
    XLL_MODIFIER_VOLATILE, XLL_TYPE_ASYNC_HANDLE, XLL_TYPE_ASYNC_VOID, XLL_TYPE_BOOL,
    XLL_TYPE_DOUBLE, XLL_TYPE_I32, XLL_TYPE_XCHAR_COUNTED, XLL_TYPE_XCHAR_NULL_TERMINATED,
    XLL_TYPE_XLOPER12_REFERENCE, XLL_TYPE_XLOPER12_VALUE,
};

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
/// One Excel registration type code for a generated function parameter.
pub enum ExcelArgumentType {
    /// `B`: 64-bit floating-point number.
    Number,
    /// `A`: Boolean.
    Boolean,
    /// `J`: 32-bit integer.
    Integer,
    /// `Q`: value-only `XLOPER12` argument.
    GeneralValue,
    /// `U`: reference-preserving `XLOPER12` argument.
    GeneralReference,
    /// `D%`: counted direct UTF-16 argument.
    CountedUtf16,
    /// `C%`: NUL-terminated direct UTF-16 argument.
    NullTerminatedUtf16,
    /// `X`: internal asynchronous completion handle.
    AsyncHandle,
}

impl ExcelArgumentType {
    const fn code(self) -> &'static str {
        match self {
            Self::Number => XLL_TYPE_DOUBLE,
            Self::Boolean => XLL_TYPE_BOOL,
            Self::Integer => XLL_TYPE_I32,
            Self::GeneralValue => XLL_TYPE_XLOPER12_VALUE,
            Self::GeneralReference => XLL_TYPE_XLOPER12_REFERENCE,
            Self::CountedUtf16 => XLL_TYPE_XCHAR_COUNTED,
            Self::NullTerminatedUtf16 => XLL_TYPE_XCHAR_NULL_TERMINATED,
            Self::AsyncHandle => XLL_TYPE_ASYNC_HANDLE,
        }
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
/// Excel registration type code for a generated function result.
pub enum ExcelReturnType {
    /// `B`: 64-bit floating-point number.
    Number,
    /// `A`: Boolean.
    Boolean,
    /// `J`: 32-bit integer.
    Integer,
    /// `Q`: an `XLOPER12` return root.
    Xloper12,
    /// `>`: asynchronous void result completed through `xlAsyncReturn`.
    AsyncVoid,
}

impl ExcelReturnType {
    const fn code(self) -> &'static str {
        match self {
            Self::Number => XLL_TYPE_DOUBLE,
            Self::Boolean => XLL_TYPE_BOOL,
            Self::Integer => XLL_TYPE_I32,
            Self::Xloper12 => XLL_TYPE_XLOPER12_VALUE,
            Self::AsyncVoid => XLL_TYPE_ASYNC_VOID,
        }
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
/// Result and argument type metadata used to construct registration type text.
pub struct FunctionSignature {
    /// Excel-visible result type.
    pub result: ExcelReturnType,
    /// Excel-visible and internal argument types in declaration order.
    pub arguments: &'static [ExcelArgumentType],
}

impl FunctionSignature {
    /// Creates a static function signature.
    pub const fn new(result: ExcelReturnType, arguments: &'static [ExcelArgumentType]) -> Self {
        Self { result, arguments }
    }
}

#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
/// Optional Excel registration modifiers for a worksheet function.
pub struct FunctionFlags {
    /// Adds the volatile modifier.
    pub volatile: bool,
    /// Requests Excel's thread-safe registration modifier.
    pub thread_safe: bool,
    /// Marks the function as macro-sheet equivalent.
    pub macro_type: bool,
    /// Requests the cluster-safe modifier when otherwise legal.
    pub cluster_safe: bool,
}

#[derive(Clone, Copy, Debug)]
/// Complete static registration metadata for one worksheet function.
pub struct FunctionRegistration {
    /// Unmangled Rust-export thunk name.
    pub rust_symbol: &'static str,
    /// Excel-visible worksheet-function name.
    pub excel_name: &'static str,
    /// Result and parameter type codes.
    pub signature: FunctionSignature,
    /// Optional Excel category.
    pub category: Option<&'static str>,
    /// Optional Excel-visible function description.
    pub description: Option<&'static str>,
    /// Excel-visible parameter names excluding the internal async handle.
    pub argument_names: &'static [&'static str],
    /// Excel-visible parameter help, parallel to [`Self::argument_names`].
    pub argument_descriptions: &'static [&'static str],
    /// Registration modifiers.
    pub flags: FunctionFlags,
}

/// Registration metadata for an XLL command, which is deliberately distinct
/// from a worksheet-function descriptor.
#[derive(Clone, Copy, Debug)]
pub struct CommandRegistration {
    /// Unmangled Rust-export command thunk name.
    pub rust_symbol: &'static str,
    /// Excel-visible command name.
    pub excel_name: &'static str,
    /// Optional Excel-visible command description.
    pub description: Option<&'static str>,
    /// Optional command shortcut text.
    pub shortcut: Option<&'static str>,
}

impl CommandRegistration {
    /// Creates command metadata with no optional help or shortcut.
    pub const fn new(rust_symbol: &'static str, excel_name: &'static str) -> Self {
        Self {
            rust_symbol,
            excel_name,
            description: None,
            shortcut: None,
        }
    }

    /// Adds Excel-visible command help text.
    pub const fn description(mut self, description: &'static str) -> Self {
        self.description = Some(description);
        self
    }

    /// Adds the Excel shortcut text.
    pub const fn shortcut(mut self, shortcut: &'static str) -> Self {
        self.shortcut = Some(shortcut);
        self
    }

    /// `I` is the documented 16-bit command return ABI; commands have no
    /// Excel-visible arguments and use macro type 2 during registration.
    pub const fn type_text(&self) -> &'static str {
        "I"
    }

    /// Validates command metadata before lifecycle registration.
    pub fn validate(&self) -> Result<(), RegistrationError> {
        if self.rust_symbol.is_empty() {
            return Err(RegistrationError::EmptyRustSymbol);
        }
        if self.excel_name.is_empty() {
            return Err(RegistrationError::EmptyExcelName);
        }
        Ok(())
    }
}

impl FunctionRegistration {
    /// Returns whether this descriptor has the verified async `>`/`X` shape.
    pub const fn is_asynchronous(&self) -> bool {
        matches!(self.signature.result, ExcelReturnType::AsyncVoid)
    }

    /// Creates function metadata with no optional category, help, or modifiers.
    pub const fn new(
        rust_symbol: &'static str,
        excel_name: &'static str,
        signature: FunctionSignature,
    ) -> Self {
        Self {
            rust_symbol,
            excel_name,
            signature,
            category: None,
            description: None,
            argument_names: &[],
            argument_descriptions: &[],
            flags: FunctionFlags {
                volatile: false,
                thread_safe: false,
                macro_type: false,
                cluster_safe: false,
            },
        }
    }

    /// Adds the Excel category.
    pub const fn category(mut self, category: &'static str) -> Self {
        self.category = Some(category);
        self
    }

    /// Adds Excel-visible function help text.
    pub const fn description(mut self, description: &'static str) -> Self {
        self.description = Some(description);
        self
    }

    /// Adds parameter names and matching Excel-visible help text.
    pub const fn arguments(
        mut self,
        names: &'static [&'static str],
        descriptions: &'static [&'static str],
    ) -> Self {
        self.argument_names = names;
        self.argument_descriptions = descriptions;
        self
    }

    /// Sets registration modifiers after validating them with [`Self::validate`].
    pub const fn flags(mut self, flags: FunctionFlags) -> Self {
        self.flags = flags;
        self
    }

    /// Builds canonical Excel registration type text after validation.
    pub fn type_text(&self) -> Result<String, RegistrationError> {
        self.validate()?;
        let mut text = String::from(self.signature.result.code());
        for argument in self.signature.arguments {
            text.push_str(argument.code());
        }
        // Microsoft specifies modifiers after the parameter codes. This
        // canonical order makes descriptors deterministic.
        if self.flags.volatile {
            text.push(XLL_MODIFIER_VOLATILE);
        }
        if self.flags.macro_type {
            text.push(XLL_MODIFIER_MACRO_SHEET);
        }
        if self.flags.thread_safe {
            text.push(XLL_MODIFIER_THREAD_SAFE);
        }
        if self.flags.cluster_safe {
            text.push(XLL_MODIFIER_CLUSTER_SAFE);
        }
        Ok(text)
    }

    /// Validates names, argument metadata, modifiers, and async shape.
    pub fn validate(&self) -> Result<(), RegistrationError> {
        if self.rust_symbol.is_empty() {
            return Err(RegistrationError::EmptyRustSymbol);
        }
        if self.excel_name.is_empty() {
            return Err(RegistrationError::EmptyExcelName);
        }
        let async_handles = self
            .signature
            .arguments
            .iter()
            .filter(|argument| **argument == ExcelArgumentType::AsyncHandle)
            .count();
        let visible_arguments = self.signature.arguments.len() - async_handles;
        if visible_arguments != self.argument_names.len() {
            return Err(RegistrationError::SignatureArgumentLengthMismatch);
        }
        if self.argument_names.len() != self.argument_descriptions.len() {
            return Err(RegistrationError::ArgumentMetadataLengthMismatch);
        }
        if self.flags.macro_type && (self.flags.thread_safe || self.flags.cluster_safe) {
            return Err(RegistrationError::IncompatibleFlags);
        }
        let asynchronous = self.signature.result == ExcelReturnType::AsyncVoid;
        if asynchronous != (async_handles == 1) || (asynchronous && self.flags.cluster_safe) {
            return Err(RegistrationError::InvalidAsyncSignature);
        }
        Ok(())
    }
}

#[derive(Clone, Copy, Debug)]
/// Static metadata for one XLL and its generated registrations.
pub struct AddInDescriptor {
    /// Excel-visible add-in name.
    pub name: &'static str,
    /// Excel-visible add-in description.
    pub description: &'static str,
    /// Worksheet function registrations.
    pub functions: &'static [FunctionRegistration],
    /// Command registrations.
    pub commands: &'static [CommandRegistration],
}

impl AddInDescriptor {
    /// Creates add-in metadata without commands.
    pub const fn new(
        name: &'static str,
        description: &'static str,
        functions: &'static [FunctionRegistration],
    ) -> Self {
        Self {
            name,
            description,
            functions,
            commands: &[],
        }
    }

    /// Adds static command metadata to this descriptor.
    pub const fn commands(mut self, commands: &'static [CommandRegistration]) -> Self {
        self.commands = commands;
        self
    }

    /// Validates the add-in and every contained registration descriptor.
    pub fn validate(&self) -> Result<(), RegistrationError> {
        if self.name.is_empty() {
            return Err(RegistrationError::EmptyAddInName);
        }
        for function in self.functions {
            function.validate()?;
        }
        for command in self.commands {
            command.validate()?;
        }
        Ok(())
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
/// Static registration metadata is inconsistent or exceeds Excel limits.
pub enum RegistrationError {
    /// The add-in name was empty.
    EmptyAddInName,
    /// The exported Rust thunk symbol was empty.
    EmptyRustSymbol,
    /// The Excel-visible name was empty.
    EmptyExcelName,
    /// Visible signature parameters and supplied names have different lengths.
    SignatureArgumentLengthMismatch,
    /// Parameter names and descriptions have different lengths.
    ArgumentMetadataLengthMismatch,
    /// Macro-sheet registration conflicts with thread-safe or cluster-safe flags.
    IncompatibleFlags,
    /// The `>` result and single `X` handle contract was not satisfied.
    InvalidAsyncSignature,
    /// Generated registration text exceeds Excel's counted-string limit.
    StringTooLong,
}

impl fmt::Display for RegistrationError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::EmptyAddInName => "add-in name must not be empty",
            Self::EmptyRustSymbol => "Rust symbol must not be empty",
            Self::EmptyExcelName => "Excel function name must not be empty",
            Self::SignatureArgumentLengthMismatch => {
                "signature and argument metadata lengths differ"
            }
            Self::ArgumentMetadataLengthMismatch => {
                "argument names and descriptions must have the same length"
            }
            Self::IncompatibleFlags => {
                "a function cannot be both thread-safe and macro-sheet equivalent"
            }
            Self::InvalidAsyncSignature => {
                "an asynchronous function requires `>` plus exactly one `X` argument and cannot be cluster-safe"
            }
            Self::StringTooLong => "registration text exceeds Excel's counted-string limit",
        })
    }
}

impl std::error::Error for RegistrationError {}

#[cfg(test)]
mod tests {
    use super::*;

    const TWO_NUMBERS: &[ExcelArgumentType] =
        &[ExcelArgumentType::Number, ExcelArgumentType::Number];

    #[test]
    fn exact_type_text_is_generated_from_the_signature() {
        let function = FunctionRegistration::new(
            "rust_add",
            "RUST.ADD",
            FunctionSignature::new(ExcelReturnType::Xloper12, TWO_NUMBERS),
        )
        .arguments(&["x", "y"], &["x", "y"])
        .flags(FunctionFlags {
            volatile: true,
            thread_safe: true,
            macro_type: false,
            cluster_safe: true,
        });
        assert_eq!(function.type_text().as_deref(), Ok("QBB!$&"));
    }

    #[test]
    fn q_and_u_and_direct_wide_strings_remain_distinct() {
        const ARGS: &[ExcelArgumentType] = &[
            ExcelArgumentType::GeneralValue,
            ExcelArgumentType::GeneralReference,
            ExcelArgumentType::CountedUtf16,
            ExcelArgumentType::NullTerminatedUtf16,
        ];
        let function = FunctionRegistration::new(
            "probe",
            "PROBE",
            FunctionSignature::new(ExcelReturnType::Xloper12, ARGS),
        )
        .arguments(&["q", "u", "d", "c"], &["", "", "", ""]);
        assert_eq!(function.type_text().as_deref(), Ok("QQUD%C%"));
    }

    #[test]
    fn incompatible_flags_and_mismatched_metadata_are_rejected() {
        let mismatch = FunctionRegistration::new(
            "add",
            "RUST.ADD",
            FunctionSignature::new(ExcelReturnType::Number, TWO_NUMBERS),
        )
        .arguments(&["x"], &["x"]);
        assert_eq!(
            mismatch.validate(),
            Err(RegistrationError::SignatureArgumentLengthMismatch)
        );

        let incompatible = FunctionRegistration::new(
            "f",
            "F",
            FunctionSignature::new(ExcelReturnType::Number, &[]),
        )
        .flags(FunctionFlags {
            volatile: false,
            thread_safe: true,
            macro_type: true,
            cluster_safe: false,
        });
        assert_eq!(
            incompatible.validate(),
            Err(RegistrationError::IncompatibleFlags)
        );
    }

    #[test]
    fn async_registration_requires_exact_void_handle_pair() {
        let valid = FunctionRegistration::new(
            "async_probe",
            "ASYNC.PROBE",
            FunctionSignature::new(
                ExcelReturnType::AsyncVoid,
                &[ExcelArgumentType::Number, ExcelArgumentType::AsyncHandle],
            ),
        )
        .arguments(&["value"], &["value"])
        .flags(FunctionFlags {
            thread_safe: true,
            ..FunctionFlags::default()
        });
        assert_eq!(valid.type_text().as_deref(), Ok(">BX$"));

        let missing = FunctionRegistration::new(
            "bad",
            "BAD",
            FunctionSignature::new(ExcelReturnType::AsyncVoid, &[ExcelArgumentType::Number]),
        )
        .arguments(&["value"], &["value"]);
        assert_eq!(
            missing.validate(),
            Err(RegistrationError::InvalidAsyncSignature)
        );

        let stray = FunctionRegistration::new(
            "bad",
            "BAD",
            FunctionSignature::new(ExcelReturnType::Number, &[ExcelArgumentType::AsyncHandle]),
        );
        assert_eq!(
            stray.validate(),
            Err(RegistrationError::InvalidAsyncSignature)
        );
    }
}
