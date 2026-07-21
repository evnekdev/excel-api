# Codex Prompt 11 — Implement Two-Dimensional Range Value Transport

## Objective

Implement safe, efficient scalar and rectangular value transfer between Rust and Excel ranges through `VARIANT` and `SAFEARRAY`.

Prioritize correctness and ownership over convenience integrations.

## Repository and branch

Repository: `evnekdev/excel-api`

Branch:

```text
feature/excel-com-11-range-arrays
```

## Required implementation

### `ExcelValue`

Implement the approved cell-value model covering empty, Boolean, number, text, Excel error, date if approved, and unknown/unsupported values without undefined behavior.

Do not confuse COM operation errors with worksheet cell errors.

### `RangeValues`

Implement:

```rust
pub struct RangeValues {
    rows: usize,
    columns: usize,
    values: Vec<ExcelValue>,
}
```

Required invariants:

- `values.len() == rows * columns`;
- documented row-major Rust indexing;
- checked dimension multiplication;
- no jagged representation;
- explicit empty-shape policy;
- safe indexing APIs and useful iterators.

### COM conversion

Implement:

- scalar `VARIANT` to/from `ExcelValue`;
- two-dimensional `SAFEARRAY` of variants to `RangeValues`;
- `RangeValues` to two-dimensional `SAFEARRAY`;
- non-zero lower bounds;
- correct row/column mapping;
- cleanup on partial failure;
- overflow checks;
- unsupported dimension and element-type errors.

Do not assume every returned SAFEARRAY contains variants unless evidence establishes it.

### Range API

Add `values2()` and `set_values2(...)` or approved equivalents. Define scalar versus matrix behavior precisely. Preserve shape for multi-cell ranges. Produce structured shape mismatch errors.

### Convenience conversions

Implement only low-cost, well-defined conversions such as rectangular `Vec<Vec<T>>`, const-generic arrays, or row slices. Validate rectangularity.

Optional `ndarray` support may be added behind a feature only if approved. Do not add Polars to the core crate.

### Formula arrays

Add rectangular formula reading/writing only if it safely reuses proven transport. Keep `Formula` and `Formula2` distinct.

## Tests

Cover:

- 1×1, 1×N, N×1, N×M;
- mixed types;
- empty and error cells;
- non-zero lower bounds;
- shape mismatch;
- overflow;
- partial conversion failure;
- exactly-once cleanup;
- row/column orientation.

Real-Excel tests must compare representative matrices with a control implementation.

## Performance evidence

Add a basic ignored benchmark or measurement tool comparing bulk 2D transfer with repeated per-cell calls. Do not create unstable CI timing thresholds.

Document bulk range transfer as preferred.

## Acceptance workflow

A workflow equivalent to this must succeed:

```rust
let values = RangeValues::from_rows([
    [1.0, 2.0],
    [3.0, 4.0],
])?;

sheet.range("A1:B2")?.set_values2(&values)?;
let roundtrip = sheet.range("A1:B2")?.values2()?;
assert_eq!(roundtrip, values);
```

Commit, push, and open a draft PR.

Do not merge. Stop after reporting the draft PR.
