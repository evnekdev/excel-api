use excel_api::{
    AddInDescriptor, FunctionFlags, FunctionRegistration, IntoExcel,
};

pub fn add(x: f64, y: f64) -> f64 {
    x + y
}

pub static FUNCTIONS: &[FunctionRegistration] = &[
    FunctionRegistration::new("add", "RUST.ADD")
        .category("Rust")
        .description("Adds two numbers")
        .arguments(&["x", "y"], &["First number", "Second number"])
        .flags(FunctionFlags {
            volatile: false,
            thread_safe: true,
            macro_type: false,
            cluster_safe: false,
        }),
];

pub static ADD_IN: AddInDescriptor = AddInDescriptor::new(
    "excel-api minimal example",
    "Minimal Rust-native Excel add-in",
    FUNCTIONS,
);

/// Initial lifecycle placeholder.
///
/// Actual registration through `Excel12v` belongs to the next milestone.
#[unsafe(no_mangle)]
pub extern "system" fn xlAutoOpen() -> i32 {
    match ADD_IN.validate() {
        Ok(()) => 1,
        Err(_) => 0,
    }
}

#[unsafe(no_mangle)]
pub extern "system" fn xlAutoClose() -> i32 {
    1
}

#[unsafe(no_mangle)]
pub extern "system" fn xlAutoFree12(_value: *mut excel_api_sys::XLOPER12) {
    // No heap-backed Excel return values exist in this milestone.
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn example_descriptor_is_valid() {
        assert_eq!(ADD_IN.validate(), Ok(()));
    }

    #[test]
    fn add_converts_to_an_excel_value() {
        assert!(add(2.0, 3.0).into_excel().is_ok());
    }
}
