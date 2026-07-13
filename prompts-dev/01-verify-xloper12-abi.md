# Prompt 01: verify the Excel 12 ABI comprehensively

## Required reading

Read all root architecture documents relevant to ABI, memory, strings, arrays,
threading, registration, and testing, plus:

- `tools/Excel2013XLLSDK/INCLUDE/XLCALL.H`
- `tools/Excel2013XLLSDK/SRC/XLCALL.CPP`
- SDK samples
- existing `excel-api-sys`
- existing `tools/abi-check`

Run the current workspace tests first.

## Objective

Make `excel-api-sys` a faithful, auditable representation of the Excel 12+ ABI
for Windows x64 MSVC. This milestone is raw ABI only.

## Source priority

1. official checked-in `XLCALL.H`;
2. official SDK source/docs;
3. Microsoft public docs;
4. the uploaded book for interpretation;
5. Excel-DNA only as a secondary cross-check.

## Implement and verify

- primitive aliases: `XCHAR`, Boolean, `RW`, `COL`, `IDSHEET`;
- `XLREF12`, variable-length `XLMREF12`, `FP12`;
- complete `XLOPER12` and every nested union/struct member;
- all `xltype*`, ownership bits, errors, flow values, return codes;
- row, column, and string limits;
- function IDs needed by milestones 01–08;
- registration type-text codes/modifiers needed by milestone 08;
- exact `Excel12`, `Excel12v`, and lifecycle callback signatures.

Keep flexible-array C layouts faithful. Do not invent safe abstractions in
`excel-api-sys`.

## ABI checker

Expand `tools/abi-check` to compare C and Rust sizes, alignments, offsets, and
constants for every supported raw definition. Normal workspace tests must not
require the SDK; the standalone checker may require Windows/MSVC/SDK.

## Restrictions

- no `ExcelStr`, `ExcelValueRef`, `ExcelOwnedValue`, or `ExcelReturn`;
- no pointer dereferencing;
- no ownership policy in the raw crate;
- no placeholder union member presented as authoritative;
- document uncertainties instead of guessing.

## Validation

```text
cargo fmt --all --check
cargo clippy --workspace --all-targets --all-features -- -D warnings
cargo test --workspace --all-features
cargo run --manifest-path tools/abi-check/Cargo.toml
```

## Acceptance criteria

1. supported modern definitions match `XLCALL.H`;
2. `xltypeRef` and `xltypeInt` cannot be confused;
3. complete supported union/layout coverage exists;
4. standalone ABI checker passes on Windows x64 MSVC;
5. ordinary workspace tests pass without SDK dependency;
6. no safe-wrapper implementation has started;
7. remaining unsupported legacy definitions are explicit.
