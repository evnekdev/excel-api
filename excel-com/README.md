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

Events, charts, existing-session attachment, marshaling, generic collections,
and a stable public API remain intentionally out of scope for this first crate
slice.

## Conditional formatting, styles, Notes, and hyperlinks

`Range::format_conditions` exposes typed conditional-format rule families;
Excel continues to evaluate formulas, ordering, priority, and `StopIfTrue`.
`Range::display_format` is a read-only view of effective displayed formatting,
not an alias for ordinary Range formatting. `Workbook::styles` provides
workbook Styles, and Range assignment supports both names and a Style object.

Theme colours use `ThemeColor`; direct `ExcelColor` values and theme values are
both left to Excel's own precedence rules. Tint and shade values are checked in
the inclusive `-1.0..=1.0` range before COM.

Modern Excel calls its legacy `Comment` objects Notes. The crate keeps them
separate from read-only, account-dependent threaded comments. Hyperlinks carry
an external `Address` and/or internal workbook `SubAddress`; the crate does
not validate or follow targets, and live coverage creates no external link.
The visible acceptance tests are opt-in and their partial/runtime-blocked
observations are recorded under `knowledge/excel-object-model/`.

## Structured data operations

`Worksheet::list_objects` provides typed `ListObjects`, `ListObject`,
`ListColumns`, and `ListRows` wrappers. Table and filter indexes are one-based;
an empty table has no `DataBodyRange`, represented as `Option<Range>`. Excel
continues to own structured-reference parsing, table-name validation, totals,
calculated-column propagation, and resize semantics.

`Range::apply_auto_filter`, `AutoFilter`, `Filters`, `Sort`, `SortFields`, and
`Validation` expose a bounded stateful filter/sort/validation surface. Range
sorting, duplicate removal, insertion, deletion, clear operations, and hidden
dimensions modify Excel state in place. Copy, cut, and PasteSpecial use only
Excel's cut/copy state; unattended callers should prefer explicit copy/cut
destinations and should not rely on arbitrary system clipboard contents.

The opt-in table integration test launches one visible Excel process and exits
it normally:

```powershell
cargo test -p excel-com --test structured_data_live -- --ignored --test-threads=1
```

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

## Range formatting

Core Range formatting uses apartment-bound `Font`, `Interior`, `Borders`, and
`Border` wrappers. Formatting getters return `MixedValue<T>` so an Excel mixed
selection is never silently coerced to a scalar default. `ExcelColor` follows
Excel's COLORREF byte order: `ExcelColor::from_rgb(12, 34, 56).raw()` is
`3678732`.

| Need | API |
|---|---|
| Font | `Range::font` |
| Fill | `Range::interior` |
| Borders | `Range::borders` |
| Number format | `Range::set_number_format` |
| Horizontal alignment | `Range::set_horizontal_alignment` |
| Vertical alignment | `Range::set_vertical_alignment` |
| Wrap text | `Range::set_wrap_text` |
| Row height | `Range::set_row_height` |
| Column width | `Range::set_column_width` |
| AutoFit | `Range::auto_fit` |

Use `range.entire_column()?.auto_fit()?` or
`range.entire_row()?.auto_fit()?` for Excel's supported AutoFit shapes. Column
widths are Excel character-width units based on the Normal style font, and row
heights are in points. The crate exposes invariant `NumberFormat`, not
`NumberFormatLocal`.

## Formulas, calculation, and auditing

Formula text is deliberately distinct from evaluated `AutomationValue` data.
`FormulaValue` reports a scalar formula string, an exact rectangular formula
array, `Mixed`, or `Empty`; it never silently turns a mixed selection into a
scalar. `set_formula`, `set_formula2`, `set_formula_r1c1`, and the local
variants accept scalar text only on a 1x1 Range. The corresponding array
setters require an `AutomationArray` whose dimensions exactly match the Range
before COM is called.

`Formula2` and `Formula2R1C1` are the dynamic-array-aware Excel members.
`has_spill`, `spilling_to_range`, and `spill_parent` are transparent Excel
queries. `FormulaArray` remains the separate legacy array-formula operation:
the crate does not emulate Ctrl+Shift+Enter or manufacture its historical
length and partial-edit restrictions. Excel errors remain structured results.

| Need | API |
|---|---|
| A1 formula | `Range::formula` / `set_formula` |
| Dynamic-array formula | `Range::formula2` / `set_formula2` |
| R1C1 formula | `Range::formula_r1c1` / `set_formula_r1c1` |
| Legacy array formula | `Range::formula_array` / `set_formula_array` |
| Formula presence | `Range::has_formula` |
| Spill ownership | `Range::has_spill`, `spilling_to_range`, `spill_parent` |
| Scoped calculation mode | `Application::calculation_mode_guard` |
| Recalculate | `Application`, `Worksheet`, or `Range` `calculate` |
| Precedents/dependents | `Range::direct_precedents`, `direct_dependents`, `precedents`, `dependents` |
| Cell discovery | `Range::special_cells` |
| Deterministic search | `Range::find`, `find_all`, `find_next`, `find_previous` |
| Explicit replacement | `Range::replace` |

Excel remains the parser and calculation engine. In particular, the crate does
not implement a formula parser, a Rust calculation engine, or any event-based
recalculation model. Find and Replace defaults send concrete remembered search
settings so they do not inherit state from an Excel Find dialog or prior call.
The `find_all` iterator detects Excel's wraparound with normalized external
addresses rather than COM identity.

## API documentation

Rustdoc describes the public wrapper and Automation-value contracts; the
generated object-model inventory describes the much larger Excel type library
and is not API documentation. Build the local crate documentation with:

```powershell
cargo doc -p excel-com --all-features --no-deps --open
```
