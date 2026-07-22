# excel-com

`excel-com` is an experimental, unpublished foundation for safe Excel COM
Automation. Its semantic and wrapper APIs may change before a first release.

The implemented path is `Application -> Workbooks -> Workbook -> Worksheets
-> Worksheet -> Range`. It supports creating a local Excel instance,
inspecting and setting visibility, controlling `DisplayAlerts` with restoration,
creating or opening a workbook, saving it or a copy, closing it with typed
options, navigating worksheets, and reading or writing a bounded Range
value/formula surface. It does not claim complete Excel object-model support.

The crate is layered as Excel wrappers, object-model member descriptors,
Automation values and dispatch invocation, then private `windows-sys` COM
ownership. Research tools exercise this crate but it does not depend on those
tools or their evidence formats.

Excel wrappers are apartment-bound and are neither `Send` nor `Sync`. Callers
create an explicit `ComApartment::sta()` and pass it to `Application::new`.
`Drop` releases COM references but never calls `Quit`; application shutdown is
an explicit operation. Raw COM pointers, `VARIANT`, and `SAFEARRAY` values are
not exposed by the ordinary API.

`AutomationValue` preserves Automation scalar distinctions and rectangular
arrays. `ExcelComError` preserves HRESULT and invocation context without
recording pointer addresses.

## File lifecycle

`Workbooks::open` sends the complete 15-position Excel signature and
`Workbook::save_as` sends all 13 positions. Omitted positions are retained as
`VT_ERROR` / `DISP_E_PARAMNOTFOUND`; argument order is reversed once at the
private dispatch boundary. `WorkbookOpenOptions` and `WorkbookSaveAsOptions`
are typed, preserve their logical argument positions, and redact passwords in
their `Debug` implementations. File paths accept `Path`/`OsStr` input directly
as Windows UTF-16 units: the wrapper neither canonicalizes nor performs a
lossy string conversion, and rejects embedded NULs before COM.

Use `Application::display_alerts_guard` for temporary alert suppression. The
guard restores the prior setting on drop, and `restore` exposes an explicit
restoration result. `Workbook::close` consumes the wrapper and uses
`SaveChanges::{Prompt, Save, Discard}`; `close_without_saving` delegates to
the explicit discard form.

See `../docs/excel-object-model/README.md` for the generated inventory and
`../docs/architecture/excel-com-project-layout.md` for repository boundaries.

Live tests are opt-in because they launch a new Excel process:

```powershell
cargo test -p excel-com --test live -- --ignored --test-threads=1
cargo test -p excel-com --test workbook_file_live -- --ignored --test-threads=1
```

Events, charts, macros, existing-session attachment, marshaling, generic
collections, formatting, and a stable public API are intentionally out of
scope for this first crate slice.

## Typed collections and Range navigation

`Workbooks`, `Worksheets`, and `Areas` expose fallible `iter()` methods backed
by owned `IEnumVARIANT` cursors. The cursors are single-pass and apartment-bound;
each item is a `Result`, an error fuses the cursor, and early drop releases the
COM enumerator. Excel controls collection order and mutation-during-iteration
semantics.

`Workbook`, `Worksheet`, and `Range` expose fallible `is_same_object` methods
based on canonical `IUnknown` identity, not names, paths, or addresses. The
crate deliberately does not implement `PartialEq`: Excel may return separate
COM objects for logically equivalent Range lookups.

Range navigation includes one-based `item`/`cell`, `cells`, signed `offset`,
strict nonzero `resize`, `rows`, `columns`, `areas`, `entire_row`, and
`entire_column`. `Application::union2` is the deliberately narrow helper used
to construct a multi-area Range. These APIs do not materialize Rust cell
collections or generalize multi-area assignment.

## Selecting, naming, converting, and evaluating ranges

The concise selection default is A1. R1C1 selection is explicit and is
converted by Excel rather than parsed by Rust. Numeric coordinates are
one-based. `ReferenceStyle` is global Excel state, so temporary changes use a
restoring guard. Workbook and worksheet Name collections have distinct Excel
scope behavior; a valid Name does not necessarily resolve to a Range.

| Need | API |
|---|---|
| A1 selection | `Worksheet::range` |
| Two A1 corners | `Worksheet::range_between` |
| R1C1 selection | `Worksheet::range_r1c1` |
| Numeric cell | `Worksheet::cell` |
| Numeric rectangle | `Worksheet::range_from_cells` |
| A1 output | `Range::address_a1` |
| R1C1 output | `Range::address_r1c1` |
| Customized address | `Range::address_with_options` |
| Formula/reference conversion | `Application::convert_formula` |
| Workbook names | `Workbook::names` |
| Worksheet names | `Worksheet::names` |
| Add name | `Names::add` |
| Resolve name to Range | `Name::range` |
| Evaluate scalar | `Application::evaluate_value` |
| Evaluate Range | `Application::evaluate_range` |

`Worksheet::range` changed from its earlier `AutomationArgument` plus optional
second-argument form to `range(&str)`. Use `range_between` for separate A1
corners instead.

## API documentation

Rustdoc describes the public wrapper and Automation-value contracts; the
generated object-model inventory describes the much larger Excel type library
and is not API documentation. Build the local crate documentation with:

```powershell
cargo doc -p excel-com --all-features --no-deps --open
```
