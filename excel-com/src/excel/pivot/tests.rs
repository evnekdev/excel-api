//! Deterministic PivotTable value and descriptor tests.

use super::*;

#[test]
fn orientations_and_versions_preserve_raw_typelib_values() {
    assert_eq!(PivotFieldOrientation::DATA.raw(), 4);
    assert_eq!(PivotTableVersion::CURRENT.raw(), -1);
    assert_eq!(PivotSourceType::from_raw(77).raw(), 77);
}

#[test]
fn filters_are_deliberately_bounded() {
    assert!(valid_filter(PivotFilterType::LABEL_CONTAINS).is_ok());
    assert!(valid_filter(PivotFilterType::from_raw(60)).is_err());
}
