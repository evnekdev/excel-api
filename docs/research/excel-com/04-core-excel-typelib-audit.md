# Core Excel Automation type-library audit

**Status:** installed-type-library evidence only; no Excel Automation runtime
implementation was introduced.
**Date:** 2026-07-21

## Objective and evidence boundary

This audit adds a reproducible evidence layer from the locally installed Excel
Automation type library. It verifies declared Automation metadata needed before
later Range runtime work: coclass and interface identities, `DISPID`s, invoke
kinds, signatures, optional/default flags, collection members, and selected
enum values. The new layer is adjacent to, not a replacement for, the
documentation-derived knowledge base:

- Canonical documentation facts remain in `knowledge/excel-object-model/data/`
  and were not changed.
- Installed-type-library facts are in
  `knowledge/excel-object-model/typelib/` and the deterministic reports in
  `knowledge/excel-object-model/generated/typelib/`.
- Runtime behavior remains unverified. In particular, this audit does not
  establish collection index bases, name lookup, enumeration lifetime, omitted
  argument effects, `VARIANT` runtime values, `SAFEARRAY` dimensions/bounds,
  activation, or shutdown behavior.

The audit added no `excel-com` crate, no `IDispatch::Invoke` path, no Excel
activation, no COM registration modification, no production dependency, no
public wrapper, and no generic Windows constant definitions.

## Tool and repeatability

`tools/excel-com-typelib-audit` is a standalone unpublished Rust tool with its
own lockfile. Its `audit` and `check` commands accept either an explicit
type-library path or the registered Excel library, and only read type
information. It calls `LoadTypeLibEx(..., REGKIND_NONE)` for an explicit file
or `LoadRegTypeLib` for Excel 1.9; neither API creates an Excel application.

The inspector uses the official `windows` crate (`windows` and `windows-core`
0.62.2), initializes only a COM apartment for reflection, and releases every
acquired `TYPEATTR`, `FUNCDESC`, `VARDESC`, and `TLIBATTR` through scoped
guards. Its output is sorted by stable record ID, UTF-8, LF-terminated, and
has a final newline. `check` performs two independent inspections and compares
their generated artifact bytes with the checked-in evidence.

The committed source manifest records only portable input identity: GUID,
version, LCID, SYSKIND, registration category, `EXCEL.EXE` basename and hash,
Excel/Windows version strings, Office bitness, tool versions, date, and target
set. It contains no raw local path.

## Inspected environment and library identity

The registered Win64 category resolved to the locally installed
`EXCEL.EXE` (the basename and SHA-256 are recorded in
[`SOURCE_MANIFEST.toml`](../../../knowledge/excel-object-model/typelib/SOURCE_MANIFEST.toml)).
The audit inspected:

| Field | Observed value |
| --- | --- |
| Library GUID | `{00020813-0000-0000-C000-000000000046}` |
| Version / LCID | 1.9 / 0 |
| Declared SYSKIND | `SYS_WIN32` |
| Registration category | `HKCR\\TypeLib\\{00020813-0000-0000-C000-000000000046}\\1.9\\0\\Win64` |
| Excel file version | 16.0.20131.20154 |
| Office bitness | 64-bit |
| Windows | Windows 10 Enterprise 25H2, build 26200.8875 |
| Type infos | 1,036 |

`SYS_WIN32` is a declared library attribute, while the registration selected
the Win64 Excel installation. It is evidence to retain, not a basis for
inferring an ABI or a wrapper architecture.

## Architectural spine and Application coclass

The audited coclass is `Application`, CLSID
`{00024500-0000-0000-C000-000000000046}`. Its default Automation interface is
`_Application`, IID `{000208D5-0000-0000-C000-000000000046}`; the coclass also
declares the source event interface `AppEvents`, IID
`{00024413-0000-0000-C000-000000000046}`. This is declared coclass metadata,
not an authorization to create an application or implement event sinks.

The generated [architectural spine report](../../../knowledge/excel-object-model/generated/typelib/architectural-spine.md)
records the core interface identities and audited member counts: `_Application`
(17), `Workbooks` (7), `_Workbook` (10), `Worksheets` (6), `_Worksheet` (7),
and `Range` (32). The audited record set preserves each reflected `DISPID`,
`INVOKEKIND`, calling convention, return type, ordered parameter list, parameter
flags, declared defaults, and interface GUID. It does not design a Rust API.

## Candidate, optional-argument, and Range findings

All 50 Prompt 03 candidate members were found in the installed library. The
[candidate signature report](../../../knowledge/excel-object-model/generated/typelib/candidate-0.1-signatures.md)
is the precise source for their declared `DISPID`s and signatures.

The [optional-arguments report](../../../knowledge/excel-object-model/generated/typelib/optional-arguments.md)
covers `Workbooks.Open`, `Workbook.SaveAs`, `Workbook.Close`, `Worksheets.Add`,
`Range.Find`, `Range.Sort`, `Application.Run`, and both
`ExportAsFixedFormat` members. It records declaration order and optional/default
flags only. It intentionally does not choose an omission representation or an
options-builder API.

The [Range declarations report](../../../knowledge/excel-object-model/generated/typelib/range-contracts.md)
captures the get/put declarations for `Value`, `Value2`, `Formula`, and
`Formula2`, plus `Text`, `HasFormula`, `Item`, `Cells`, `Rows`, `Columns`,
`Offset`, `Resize`, `Address`, `ClearContents`, `Find`, `Sort`, and the
Worksheet Range entry points. Reflected `VARIANT` declarations are not a
runtime `VARTYPE` observation and do not reveal `SAFEARRAY` orientation,
dimension count, lower bounds, or ownership.

## Collections, enums, aliases, and unresolved work

The [collection report](../../../knowledge/excel-object-model/generated/typelib/collections.md)
records `Count`, `Item`, `_Default`, `_NewEnum`, `Add`, and relevant mutation
members for `Workbooks`, `Worksheets`, `Sheets`, `Windows`, `Names`,
`ListObjects`, `ChartObjects`, and `Shapes`. It confirms declared default
members at `DISPID 0` and `_NewEnum` at `DISPID -4` where reflected; neither
fact determines runtime indexing, enumeration, or COM lifetime behavior.

Twenty referenced or required enums, including the minimum requested
calculation/file-format/reference/direction/find/search/sort/sheet/fixed-format
families, were reflected with their values in
[enum values](../../../knowledge/excel-object-model/generated/typelib/enum-values.md).
No aliases were required by the selected signatures. The current target set
emitted no unresolved entries; that is a scope result, not a statement that the
full Excel object model is complete. Runtime validation of the targets remains
the outstanding next gate.

## Documentation differences

The [documentation-differences report](../../../knowledge/excel-object-model/generated/typelib/documentation-differences.md)
compares the documentation-derived parameter count and member kind with the
compatible reflected member without rewriting either evidence source. It found
these declared-count differences:

- `Application.Run`: documentation has 2 parameters; the typelib declares 31.
- `Workbook.SaveAs`: documentation has 12 parameters; the typelib declares 13.
- `Range.ClearContents`: documentation has 0 parameters; the typelib declares
  one optional parameter.
- `Range.Sort`: documentation has 15 parameters; the typelib declares 16.

The report retains all other candidate comparisons and their reflected
`DISPID`s. These are compatibility observations for later design/research, not
claims that any argument may safely be omitted or that the documentation is
incorrect at runtime.

## Reproduction and validation

The audit used the following observed machine values, solely to regenerate this
specific local evidence:

```powershell
cargo run --offline --manifest-path tools/excel-com-typelib-audit/Cargo.toml -- audit `
    --root knowledge/excel-object-model `
    --typelib <installed-EXCEL.EXE> `
    --windows-version "Windows 10 Enterprise 25H2 build 26200.8875" `
    --excel-file-version "16.0.20131.20154" `
    --office-bitness "64-bit"
```

The audit did not create or automate Excel; no Excel process was started by the
tool. Before merge, the following checks are required and were run for this
change:

```text
cargo fmt --all --check
cargo clippy --workspace --all-targets --all-features -- -D warnings
cargo test --workspace --all-features
cargo test --workspace --doc
cargo fmt --manifest-path tools/excel-com-kb/Cargo.toml --check
cargo clippy --offline --manifest-path tools/excel-com-kb/Cargo.toml --all-targets -- -D warnings
cargo test --offline --manifest-path tools/excel-com-kb/Cargo.toml
cargo run --offline --manifest-path tools/excel-com-kb/Cargo.toml -- check --root knowledge/excel-object-model
cargo fmt --manifest-path tools/excel-com-typelib-audit/Cargo.toml --check
cargo clippy --offline --manifest-path tools/excel-com-typelib-audit/Cargo.toml --all-targets -- -D warnings
cargo test --offline --manifest-path tools/excel-com-typelib-audit/Cargo.toml
cargo run --offline --manifest-path tools/excel-com-typelib-audit/Cargo.toml -- check --root knowledge/excel-object-model <same input values>
git diff --check
```

Final review also checks deterministic regeneration, Markdown links, LF/final
newlines, evidence schema shape, source-manifest portability, absence of raw
absolute local paths, and that `knowledge/excel-object-model/data/` remains
unchanged.
