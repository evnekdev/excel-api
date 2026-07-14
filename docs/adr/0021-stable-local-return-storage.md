# ADR-0021: Stable local return storage owns backing allocations by Rust fields

## Status

Accepted and implemented in M5.

## Decision

`ExcelReturn` owns a boxed `repr(C)` `ReturnAllocation` whose `XLOPER12` root
is the first field at byte offset zero. The remaining fields own an optional
single boxed `XLOPER12` element block and a table of independent boxed counted
UTF-16 buffers. No raw ABI pointer is used as cleanup authority.

String boxes reach final addresses before array elements receive pointers.
The element box reaches its final address before the root receives `lparray`.
No pointer targets a field of the outer owner, so moving the `ExcelReturn` box
handle cannot invalidate any address. The root is exposed read-only and only
after the root-first owner is boxed.

Materialization rechecks the full Prompt 04 storage totals before root creation.
The Rust buffer-owner table is allocator-dependent container bookkeeping and
is excluded from ABI payload bytes and backing-object allocation counts.
Fallible reservations are used for string, element, and metadata buffers;
normal field drop cleans every partial construction path.

The sole production unsafe block zero-initializes the complete raw union before
an active member is selected. This is valid because every union field consists
only of integer, floating-point, or raw-pointer values with valid all-zero bit
patterns. Zeroing defines otherwise unused union bytes.

## Consequences

M5 supports safe local ownership only. No ownership bit is set, no raw pointer
is consumed, and no Excel callback can free the allocation. M6 must add a
one-way handoff state and post-handoff reconstruction without changing these
backing-storage invariants.
