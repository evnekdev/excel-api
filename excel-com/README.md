# excel-com

`excel-com` is an experimental Windows-only Excel COM Automation crate. Version
0.1 establishes the initial public API, but breaking changes remain possible
before 1.0. It supports Windows desktop Excel through COM; it does not support
macOS Excel, Excel Online, Linux-native Excel, LibreOffice, Office Scripts, or
Microsoft Graph workbook APIs.

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
create an explicit `ComApartment::sta()` and start a crate-owned session with
`OwnedApplication::new`. `Drop` releases COM references but never calls
`Quit`; application shutdown is an explicit operation. Raw COM pointers,
`VARIANT`, and `SAFEARRAY` values are not exposed by the ordinary API.

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

Events, COM marshaling, VBA source editing, ActiveX/Form controls, Data
Model/DAX mutation, and a stable pre-1.0 API remain out of scope.

## Charts, drawings, pictures, and sparklines

`Worksheet::chart_objects` creates embedded charts from an Excel `Range` with
explicit point bounds; `Workbook::charts` exposes chart sheets. `Chart` covers
source data, series, axes, titles, legends, data labels, trendlines, error
bars, and export through an installed Excel filter. Excel remains the chart
calculation and rendering engine.

`Worksheet::shapes` covers a deliberately bounded Office drawing surface:
AutoShapes, lines, local file-backed pictures, text boxes, placement, z-order,
and deletion. Shape grouping is intentionally unavailable pending controlled
runtime evidence. `Range::copy_picture` uses Excel's cut/copy state and does
not read the operating-system clipboard; use `Application::cut_copy_mode` and
`clear_cut_copy_mode` to manage that state. `Range::sparkline_groups` creates
line, column, and win/loss groups from Excel ranges.

| Need | API |
|---|---|
| Embedded chart | `Worksheet::add_chart` / `Worksheet::chart_objects` |
| Chart sheet | `Workbook::charts` |
| Series and axes | `Chart::series_collection` / `Chart::axes` |
| Shapes and pictures | `Worksheet::shapes` |
| Excel-native picture copy | `Range::copy_picture` |
| Sparklines | `Range::sparkline_groups` |

The visible tests below each start and quit a local Excel server. They never
attach to an existing interactive session:

```powershell
cargo test -p excel-com --test charts_live -- --ignored --test-threads=1
cargo test -p excel-com --test shapes_images_live -- --ignored --test-threads=1
cargo test -p excel-com --test sparklines_live -- --ignored --test-threads=1
```

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

## Text interchange, analysis, and controlled links

`Workbooks::open_text` delegates CSV, TSV, and fixed-width import directly to
Excel: delimiter handling, quoted fields, type inference, and locale behaviour
remain Excel-owned. `Workbook::save_as_text` uses `SaveAs` with exact text
formats; Excel normally exports only its active worksheet. `Range` provides
Text to Columns, directional fills, AutoFill, DataSeries, Flash Fill,
Excel-native transpose, Advanced Filter, Subtotal, Consolidate, Goal Seek,
What-if Data Tables, and `SpecialCells` convenience lookups.

Scenarios are typed Excel collections. External-link APIs preserve exact source
strings and require callers to name operations explicitly. `break_link` is
destructive: Excel replaces link formulas with their current values.
`AskToUpdateLinksGuard` controls link prompts only; macro security is a
separate setting. The Prompt 18 live suites are intentionally ignored while a
fresh `Workbooks.Add` baseline is runtime-blocked.

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

## Quick start: copied fixture

Use a workbook supplied by your application or a copied fixture; do not make
`Workbooks.Add` part of unattended startup. The current development host has a
machine-specific `0x800A03EC` Add limitation, so fixture opening is the
release smoke-test path.

```no_run
use excel_com::{
    AutomationValue, ComApartment, OwnedApplication, SaveChanges, WorkbookCloseOptions,
    WorkbookOpenOptions,
};
use std::{env, time::Duration};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let path = env::args().nth(1).expect("pass a copied .xlsx path");
    let apartment = ComApartment::sta()?;
    let excel = OwnedApplication::new(&apartment)?;
    let workbook = excel.workbooks()?.open(path, WorkbookOpenOptions::new())?;
    {
        let sheet = workbook.worksheets()?.item_by_index(1)?;
        let cell = sheet.range("A1")?;
        cell.set_value(AutomationValue::Text("hello from excel-com".into()))?;
    }
    workbook.close(WorkbookCloseOptions { save_changes: SaveChanges::Discard, ..WorkbookCloseOptions::new() })?;
    excel.quit_and_wait(Duration::from_secs(30))?;
    Ok(())
}
```

## Ownership, errors, and security

`OwnedApplication` represents a server created by this crate and is the only
type with `quit`. `AttachedApplication::attach` borrows an existing Excel
session and never quits it or closes unrelated workbooks. Both require the
creating STA thread to remain alive and, where Excel needs it, to pump messages.

All operations return `ExcelComError`; invocation errors preserve object,
member, DISPID, HRESULT, and available EXCEPINFO through stable accessors.
Passwords and `SecretStringValue` diagnostics are redacted. External data,
links, PDF export, file replacement prompts, and printer-dependent PageSetup
are Excel/provider dependent. The optional `macro-runtime` feature enables
`Application::run_macro`; only enable it for trusted workbooks.

Specialized APIs are under `drawing`, `formatting`, `tables`, `presentation`,
`data`, `external_data`, and `pivot`. Excel versions are expected to vary in
their support for newer chart, provider, and workbook features. See the
release notes and known limitations before relying on unattended automation.
