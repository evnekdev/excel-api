# Excel Object Model inventory

This maintained inventory is generated from the locally registered Excel type library plus explicit policy metadata. It is an implementation guide for the experimental `excel-com` crate, not a claim of complete wrapper coverage.

Every object has independent `surface_class` (what the typelib exposes) and `roadmap_class` (the wrapper plan) fields. Standard IUnknown and IDispatch entries are retained structurally but excluded from human Excel-member coverage. The experimental crate implements a bounded `Application -> Workbooks -> Workbook -> Worksheets -> Worksheet -> Range` slice. See [STATUS](STATUS.md) for coverage and the indexes directory for objects, members, events, enums, and deferred surface area. Historical runtime research remains in `docs/research/excel-com/`.
