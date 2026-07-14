use excel_api::{AddInDescriptor, FunctionFlags, FunctionRegistration};

pub fn add(x: f64, y: f64) -> f64 {
    x + y
}

pub static FUNCTIONS: &[FunctionRegistration] = &[FunctionRegistration::new("add", "RUST.ADD")
    .category("Rust")
    .description("Adds two numbers")
    .arguments(&["x", "y"], &["First number", "Second number"])
    .flags(FunctionFlags {
        volatile: false,
        thread_safe: true,
        macro_type: false,
        cluster_safe: false,
    })];

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
/// # Safety
///
/// Excel must pass null or a unique pointer produced by
/// `ExcelReturn::into_raw_for_excel` from this loaded XLL and call this export
/// exactly once for that pointer.
pub unsafe extern "system" fn xlAutoFree12(value: *mut excel_api_sys::XLOPER12) {
    // SAFETY: this export has the same contract and exact WINAPI ABI as the
    // reusable callback body.
    unsafe { excel_api::xl_auto_free12(value) };
}

const _: excel_api_sys::XlAutoFree12Fn = xlAutoFree12;

#[cfg(test)]
mod tests {
    use super::*;
    use excel_api::IntoExcel;

    #[test]
    fn example_descriptor_is_valid() {
        assert_eq!(ADD_IN.validate(), Ok(()));
    }

    #[test]
    fn add_converts_to_an_excel_value() {
        assert!(add(2.0, 3.0).into_excel().is_ok());
    }

    #[test]
    fn exported_auto_free_delegates_to_the_core_owner_cleanup() {
        let pointer = excel_api::ExcelReturnValue::from("owned by XLL")
            .plan()
            .unwrap()
            .materialize()
            .unwrap()
            .into_raw_for_excel();
        // SAFETY: this is the unique fresh handoff pointer and this test calls
        // the exported callback exactly once.
        unsafe { xlAutoFree12(pointer) };
    }
}
