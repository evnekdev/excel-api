mod support;

use support::Fixture;

#[test]
fn repository_owned_fixtures_are_unchanged_and_safe() {
    for fixture in [
        Fixture::BlankXlsx,
        Fixture::BlankXlsm,
        Fixture::ChartSource,
        Fixture::PivotSource,
        Fixture::LocalQuerySource,
    ] {
        fixture.verify().expect("fixture integrity");
    }
}

#[test]
fn fixture_copy_is_temporary_and_separate_from_the_tracked_source() {
    let temporary = Fixture::BlankXlsx
        .copy_for_test()
        .expect("temporary fixture copy");
    assert_ne!(temporary.path(), Fixture::BlankXlsx.path());
    assert!(temporary.path().exists());
}
