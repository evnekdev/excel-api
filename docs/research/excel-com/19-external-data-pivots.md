# Prompt 19: external data and PivotTables

## Baseline and scope

The branch starts at `1f1f15935c6817682de36dad37af0191e39c0544`. The
fresh-process baseline had zero `EXCEL.EXE` processes, activated Excel, and
failed exactly once at `Workbooks.Add` with outer `0x80020009` and Excel SCODE
`0x800A03EC`. The transient Excel instance exited naturally. There was no
repository-owned blank-workbook fixture suitable for a fallback, so all
workbook-dependent live observations are runtime-blocked rather than retried.

The subsystem is source-derived from the installed Excel type library and uses
the existing apartment-bound `DispatchObject` model. It neither contacts a
public service nor accesses a non-owned database. No provider, DSN, credential,
Office security setting, or registration setting was installed or changed.

## Architecture and security

`external_data` contains typed `Connections`, `WorkbookConnection`, OLE DB,
ODBC, text and web wrappers; `QueryTables`; `WorkbookQueries`; and refresh
orchestration. `pivot` contains `PivotCaches`, `PivotTable`, fields, items,
filters, and read-only slicer inspection. Thin accessors are retained on
`Workbook`, `Worksheet`, and `Application`; public wrappers never expose raw
COM pointers and remain apartment-bound.

`SecretStringValue` represents connection strings and M formulas. Its `Debug`
and `Display` implementations redact the contents. Reading requires the
deliberately explicit `expose_secret()` method. Diagnostic errors and evidence
record member names, HRESULTs, and classifications but never connection strings,
formula text, provider credentials, pointers, PIDs, or HWNDs.

The OLE DB and ODBC option structs are bounded descriptors only. Connection
creation is deliberately not exposed: no deterministic installed local provider
was available for safe validation, so this area is provider-blocked rather than
guessing at `Connections.Add` or `Add2` semantics.

## QueryTables and refresh

`QueryTables` supports Count, one-based or name Item access, and fallible fused
`IEnumVARIANT` iteration. Creation supports an existing workbook connection or
an owned local text path. The local-text route supplies `TEXT;` followed by the
path, uses the established FieldInfo SAFEARRAY encoding, and writes Excel's
typed `TextFile*` properties. In particular, `TextFileOtherDelimiter` is
treated as its string delimiter character rather than a Boolean flag.

`QueryTable::refresh` preserves Excel's Boolean return. `Refreshing`,
`CancelRefresh`, result range, destination, connection, formatting, and
refresh-on-open members are typed. `Workbook::refresh_all` dispatches Excel's
workbook-wide operation. `is_refreshing` deliberately reports only observable
QueryTables because Excel exposes no universal workbook refresh-state flag.

Cancellation is best effort: it invokes `CancelRefresh` only for objects that
currently report `Refreshing`; OLE DB and ODBC connections participate, while
unsupported kinds are counted. `Application::calculate_until_async_queries_done`
uses Excel's own async wait. `Workbook::wait_for_refresh` is bounded by a
caller-supplied nonzero timeout and poll interval, runs on the owning STA, does
not create a Rust background thread, and returns a timeout report rather than
inventing a COM failure.

No live local text refresh, cancellation, or async-wait behavior was observed:
the host failed before a workbook could be created. The ignored live test is
retained for a future healthy host and uses only a temporary local CSV.

## Workbook queries

The installed type library exposes `Workbook.Queries`, `WorkbookQuery`, Count,
Item, and `_NewEnum`. The wrapper supports read-only name, description, formula
(redacted), refresh, and deletion. Editing M formulas is intentionally excluded.
There is no reviewed repository-owned local Power Query fixture, and the current
host cannot create a workbook, so workbook-query runtime behavior is both
fixture-blocked and runtime-blocked.

## PivotCaches and PivotTables

`PivotCaches` supports typed count, item, iteration, and source creation from a
Range, a ListObject's full range, or a pre-existing `WorkbookConnection`.
Range-backed source creation asks Excel for an external R1C1 address, avoiding
an unsafe attempt to pass a Range dispatch where the type library expects text.
`source_data` and command text remain `AutomationValue` so their physical
variant representations are not erased; cache connection text is redacted.

The installed `PivotCache.CreatePivotTable` type-library signature has four
arguments: destination, name, read-data, and `DefaultVersion`. The public
`version` option is a compatibility alias for `DefaultVersion`; conflicting
values are rejected instead of silently choosing one. A created PivotTable
provides its ranges, cache, typed field collections, refresh, update, style,
source-data change, and cache-change operations.

`apply_layout` applies typed field orientations and one-based positions, then
adds typed data fields with aggregation and optional number formatting. Its
private `ManualUpdate` guard captures state, restores explicitly, also attempts
non-panicking Drop restoration, prevents double restoration, and returns both
the operation and restoration errors when both occur.

`PivotItems` preserve Excel's ability to reject hiding the final visible item.
`PivotFilters` offers a limited typed set of label and value predicates,
including equality, text comparisons, ranges, and top-count/top-percent; unknown
filter types are rejected before COM invocation. `SlicerCaches`, `SlicerCache`,
`Slicers`, and `Slicer` are read-only inspection wrappers where present in the
installed version. Slicer creation remains excluded.

No cache source variant, PivotTable refresh return, absent range behavior,
orientation membership change, filter acceptance, or persistence behavior was
observed live because `Workbooks.Add` failed. The ignored Pivot live test uses
only an owned range and closes the unsaved workbook.

## Inventory, documentation, and refactor

Every member invoked by the new wrappers has an exact registry ID. The inventory
now recognizes typed collection metadata for connections, QueryTables, workbook
queries, PivotCaches, PivotTables, PivotFields, PivotItems, PivotFilters,
SlicerCaches, and Slicers. Controlled external-data and pivot capability groups
are attached to Workbook metadata and generated into the object documentation.

The historical 2,958-line drawing implementation was split into focused modules
for types, helpers, chart objects, charts, chart sheets, series, axes, labels,
trendlines, shapes, Office formatting, sparklines, and export. The legacy file
is now a negligible compatibility placeholder; public names and re-exports stay
in `drawing/mod.rs`.

The normalized evidence set under
`knowledge/excel-object-model/external-data-pivots/` records source-derived
facts separately from runtime-blocked, fixture-blocked, and provider-blocked
observations. It does not rewrite historical evidence.
