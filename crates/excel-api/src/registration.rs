use core::fmt;

/// Registration behavior flags.
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct FunctionFlags {
    pub volatile: bool,
    pub thread_safe: bool,
    pub macro_type: bool,
    pub cluster_safe: bool,
}

/// Static description of one worksheet function.
#[derive(Clone, Copy, Debug)]
pub struct FunctionRegistration {
    pub rust_symbol: &'static str,
    pub excel_name: &'static str,
    pub category: Option<&'static str>,
    pub description: Option<&'static str>,
    pub argument_names: &'static [&'static str],
    pub argument_descriptions: &'static [&'static str],
    pub flags: FunctionFlags,
}

impl FunctionRegistration {
    pub const fn new(rust_symbol: &'static str, excel_name: &'static str) -> Self {
        Self {
            rust_symbol,
            excel_name,
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

    pub fn validate(&self) -> Result<(), RegistrationError> {
        if self.rust_symbol.is_empty() {
            return Err(RegistrationError::EmptyRustSymbol);
        }
        if self.excel_name.is_empty() {
            return Err(RegistrationError::EmptyExcelName);
        }
        if self.argument_names.len() != self.argument_descriptions.len() {
            return Err(RegistrationError::ArgumentMetadataLengthMismatch);
        }
        if self.flags.thread_safe && self.flags.macro_type {
            return Err(RegistrationError::IncompatibleFlags);
        }
        Ok(())
    }
}

/// Static descriptor for one complete add-in.
#[derive(Clone, Copy, Debug)]
pub struct AddInDescriptor {
    pub name: &'static str,
    pub description: &'static str,
    pub functions: &'static [FunctionRegistration],
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
        }
    }

    pub fn validate(&self) -> Result<(), RegistrationError> {
        if self.name.is_empty() {
            return Err(RegistrationError::EmptyAddInName);
        }

        for function in self.functions {
            function.validate()?;
        }

        Ok(())
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum RegistrationError {
    EmptyAddInName,
    EmptyRustSymbol,
    EmptyExcelName,
    ArgumentMetadataLengthMismatch,
    IncompatibleFlags,
}

impl fmt::Display for RegistrationError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::EmptyAddInName => "add-in name must not be empty",
            Self::EmptyRustSymbol => "Rust symbol must not be empty",
            Self::EmptyExcelName => "Excel function name must not be empty",
            Self::ArgumentMetadataLengthMismatch => {
                "argument names and descriptions must have the same length"
            }
            Self::IncompatibleFlags => {
                "a function cannot initially be both thread-safe and macro-type"
            }
        })
    }
}

impl std::error::Error for RegistrationError {}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn rejects_mismatched_argument_metadata() {
        let function = FunctionRegistration::new("add", "RUST.ADD")
            .arguments(&["x", "y"], &["First argument"]);
        assert_eq!(
            function.validate(),
            Err(RegistrationError::ArgumentMetadataLengthMismatch)
        );
    }
}
