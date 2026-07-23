//! Compile-time thread-affinity assertions for representative public wrappers.

use excel_com::drawing::{Chart, Shape};
use excel_com::external_data::QueryTable;
use excel_com::pivot::PivotTable;
use excel_com::{
    Application, AttachedApplication, ComApartment, OwnedApplication, Range, Workbook, Worksheet,
};
use static_assertions::assert_not_impl_any;

assert_not_impl_any!(ComApartment: Send, Sync);
assert_not_impl_any!(Application: Send, Sync);
assert_not_impl_any!(Workbook: Send, Sync);
assert_not_impl_any!(Worksheet: Send, Sync);
assert_not_impl_any!(Range: Send, Sync);
assert_not_impl_any!(Chart: Send, Sync);
assert_not_impl_any!(Shape: Send, Sync);
assert_not_impl_any!(QueryTable: Send, Sync);
assert_not_impl_any!(PivotTable: Send, Sync);
assert_not_impl_any!(OwnedApplication<'static>: Send, Sync);
assert_not_impl_any!(AttachedApplication<'static>: Send, Sync);

#[test]
fn representative_wrappers_remain_apartment_bound() {
    // The assertions above deliberately fail compilation if a wrapper becomes
    // transferable or shareable across threads.
}
