# ADR-0016: References are not values

Keep `xltypeRef`/`xltypeSRef` as separate reference types; explicit coercion is
required to obtain values.

## Status

Implemented through M3. Callback references remain distinct borrowed types;
conversion to `ExcelValue` returns `UnsupportedReference`. Owned references and
explicit coercion require a later approved contract.
