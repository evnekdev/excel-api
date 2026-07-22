# Excel Object Model inventory

This maintained inventory is generated from the locally registered Excel type library plus explicit policy metadata. It is an implementation guide for the experimental `excel-com` crate, not a claim of complete wrapper coverage.

The experimental crate implements a bounded `Application -> Workbooks -> Workbook -> Worksheets -> Worksheet -> Range` slice. See [STATUS](STATUS.md) for coverage and the indexes directory for objects, members, events, enums, and deferred surface area. Status values and surface classes are controlled and machine-readable. Historical runtime research remains in `docs/research/excel-com/`.
