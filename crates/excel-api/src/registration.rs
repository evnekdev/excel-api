use core::fmt;

use excel_api_sys::{
    XLL_MODIFIER_CLUSTER_SAFE, XLL_MODIFIER_MACRO_SHEET, XLL_MODIFIER_THREAD_SAFE,
    XLL_MODIFIER_VOLATILE, XLL_TYPE_ASYNC_HANDLE, XLL_TYPE_ASYNC_VOID, XLL_TYPE_BOOL,
    XLL_TYPE_DOUBLE, XLL_TYPE_I32, XLL_TYPE_XCHAR_COUNTED, XLL_TYPE_XCHAR_NULL_TERMINATED,
    XLL_TYPE_XLOPER12_REFERENCE, XLL_TYPE_XLOPER12_VALUE,
};

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum ExcelArgumentType {
    Number,
    Boolean,
    Integer,
    GeneralValue,
    GeneralReference,
    CountedUtf16,
    NullTerminatedUtf16,
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
pub enum ExcelReturnType {
    Number,
    Boolean,
    Integer,
    Xloper12,
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
pub struct FunctionSignature {
    pub result: ExcelReturnType,
    pub arguments: &'static [ExcelArgumentType],
}

impl FunctionSignature {
    pub const fn new(result: ExcelReturnType, arguments: &'static [ExcelArgumentType]) -> Self {
        Self { result, arguments }
    }
}

#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct FunctionFlags {
    pub volatile: bool,
    pub thread_safe: bool,
    pub macro_type: bool,
    pub cluster_safe: bool,
}

#[derive(Clone, Copy, Debug)]
pub struct FunctionRegistration {
    pub rust_symbol: &'static str,
    pub excel_name: &'static str,
    pub signature: FunctionSignature,
    pub category: Option<&'static str>,
    pub description: Option<&'static str>,
    pub argument_names: &'static [&'static str],
    pub argument_descriptions: &'static [&'static str],
    pub flags: FunctionFlags,
}

/// Registration metadata for an XLL command, which is deliberately distinct
/// from a worksheet-function descriptor.
#[derive(Clone, Copy, Debug)]
pub struct CommandRegistration {
    pub rust_symbol: &'static str,
    pub excel_name: &'static str,
    pub description: Option<&'static str>,
    pub shortcut: Option<&'static str>,
}

impl CommandRegistration {
    pub const fn new(rust_symbol: &'static str, excel_name: &'static str) -> Self {
        Self {
            rust_symbol,
            excel_name,
            description: None,
            shortcut: None,
        }
    }

    pub const fn description(mut self, description: &'static str) -> Self {
        self.description = Some(description);
        self
    }

    pub const fn shortcut(mut self, shortcut: &'static str) -> Self {
        self.shortcut = Some(shortcut);
        self
    }

    /// `I` is the documented 16-bit command return ABI; commands have no
    /// Excel-visible arguments and use macro type 2 during registration.
    pub const fn type_text(&self) -> &'static str {
        "I"
    }

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
    pub const fn is_asynchronous(&self) -> bool {
        matches!(self.signature.result, ExcelReturnType::AsyncVoid)
    }

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

    pub const fn category(mut self, category: &'static str) -> Self {
        self.category = Some(category);
        self
    }

    pub const fn description(mut self, description: &'static str) -> Self {
        self.description = Some(description);
        self
    }

    pub const fn arguments(
        mut self,
        names: &'static [&'static str],
        descriptions: &'static [&'static str],
    ) -> Self {
        self.argument_names = names;
        self.argument_descriptions = descriptions;
        self
    }

    pub const fn flags(mut self, flags: FunctionFlags) -> Self {
        self.flags = flags;
        self
    }

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
pub struct AddInDescriptor {
    pub name: &'static str,
    pub description: &'static str,
    pub functions: &'static [FunctionRegistration],
    pub commands: &'static [CommandRegistration],
}

impl AddInDescriptor {
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

    pub const fn commands(mut self, commands: &'static [CommandRegistration]) -> Self {
        self.commands = commands;
        self
    }

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
pub enum RegistrationError {
    EmptyAddInName,
    EmptyRustSymbol,
    EmptyExcelName,
    SignatureArgumentLengthMismatch,
    ArgumentMetadataLengthMismatch,
    IncompatibleFlags,
    InvalidAsyncSignature,
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
