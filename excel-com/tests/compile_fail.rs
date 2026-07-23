//! Public API misuse cases that must remain rejected at compile time.

#[test]
fn apartment_and_ownership_misuse_is_rejected() {
    let cases = trybuild::TestCases::new();
    cases.compile_fail("tests/ui/*.rs");
}
