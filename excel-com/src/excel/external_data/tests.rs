//! Deterministic public-value tests for external-data wrappers.

use super::*;

#[test]
fn secret_debug_and_display_are_redacted() {
    let value = SecretStringValue::new("confidential-test-value").expect("secret");
    assert_eq!(format!("{value}"), "REDACTED");
    assert!(!format!("{value:?}").contains("confidential-test-value"));
}

#[test]
fn forward_compatible_connection_types_round_trip() {
    assert_eq!(ConnectionType::from_raw(-123).raw(), -123);
    assert_eq!(CommandType::SQL.raw(), 2);
}
