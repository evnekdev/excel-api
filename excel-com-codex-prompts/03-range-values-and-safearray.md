# Codex Prompt 03 — Research `VARIANT`, `SAFEARRAY`, and Excel Range Semantics

## Objective

Determine the exact value-conversion and array-transport requirements for reading and writing Excel ranges through COM Automation.

This prompt may add small research probes, but it must not add production `excel-com` APIs.

## Repository and branch

Repository: `evnekdev/excel-api`

Branch:

```text
research/excel-com-03-range-values
```

## Required preparation

Read prompts 01 and 02 research outputs and official Microsoft documentation for:

- `Range.Value`;
- `Range.Value2`;
- `Range.Formula`;
- `Range.Formula2`;
- `Range.HasFormula`;
- Excel error values;
- Automation `VARIANT`;
- `SAFEARRAY` APIs.

Inspect the locally available Excel type library if supported by the environment.

## Deliverables

Create:

```text
docs/research/excel-com/03-variant-safearray-and-range-values.md
```

Optional research code may be added under:

```text
tools/excel-com-probes/
```

Research code must be unpublished, clearly experimental, Windows-only where necessary, and free of production API commitments.

## Required cases

### Range shapes

Establish documented and observed behavior for:

- one cell;
- one row with multiple columns;
- one column with multiple rows;
- rectangular multi-cell range;
- entire row or column where practical;
- noncontiguous multi-area range;
- edge cases Excel permits.

### Cell values

Test and document:

- empty cell;
- blank string;
- text;
- Boolean;
- integer-looking number;
- floating-point number;
- date;
- Currency;
- Excel error;
- formula result;
- formula text;
- dynamic-array formula;
- unsupported or object-valued cells where applicable.

### `Value` versus `Value2`

Document:

- returned `VARTYPE`;
- Date handling;
- Currency handling;
- error handling;
- scalar versus array behavior;
- recommended Rust default;
- whether explicit `value()` and `value2()` escape methods are justified.

### `SAFEARRAY` structure

For every tested multi-cell result record:

- dimension count;
- lower and upper bounds per dimension;
- element type;
- indexing order;
- mapping between Excel row/column coordinates and SAFEARRAY indices;
- ownership and destruction requirements.

Do not assume zero-based bounds.

### Writing arrays

Test and document:

- matching two-dimensional arrays;
- scalar to multiple cells;
- row vectors;
- column vectors;
- shape mismatch;
- jagged arrays;
- empty arrays;
- mixed types;
- missing, null, empty, and error values;
- scalar and rectangular formulas.

### Error values

Map common Excel errors separately from COM failures:

```text
#NULL!
#DIV/0!
#VALUE!
#REF!
#NAME?
#NUM!
#N/A
#SPILL!
#CALC!
```

Determine which are represented as `VT_ERROR`, the returned error codes, and how unknown/newer values should be preserved.

## Control implementations

Where possible, implement equivalent probes in raw Rust COM, Python pywin32, and C# Interop or VBA. Use controls to distinguish Excel behavior from Rust conversion bugs.

Do not commit machine-specific paths or private data.

## Proposed value model

End with a reasoned proposal for:

```rust
pub enum AutomationValue { /* ... */ }

pub enum ExcelValue { /* ... */ }

pub struct RangeValues {
    rows: usize,
    columns: usize,
    values: Vec<ExcelValue>,
}
```

Analyze:

- whether Automation and Excel values should be separate;
- scalar versus matrix return types;
- row-major Rust representation;
- unknown `VARTYPE` preservation;
- date representation;
- error representation;
- optional `chrono`, `ndarray`, and `serde` features;
- why Polars should or should not be in the core dependency graph.

## Evidence format

For each experiment record:

```text
Environment
Excel version/build
Operation
Expected result
Observed HRESULT
Observed VARIANT type
Observed SAFEARRAY dimensions and bounds
Control result
Conclusion
Remaining uncertainty
```

## Validation and completion

Run all applicable workspace checks and new probe tests. Do not claim a real-Excel result that was not actually executed.

Commit, push, and open a draft PR containing documented contracts, evidence, conversion risks, and unresolved cases.

Do not merge. Stop after reporting the draft PR.
