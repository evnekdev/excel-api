# Codex Prompt 10 — Implement Workbook, Worksheet, and Range Wrappers

## Objective

Implement the first complete typed Excel Automation vertical slice:

```text
ExcelApp → Workbooks → Workbook → Worksheets → Worksheet → Range
```

Support scalar values and formulas. Rectangular transport remains for the next prompt.

## Repository and branch

Repository: `evnekdev/excel-api`

Branch:

```text
feature/excel-com-10-core-object-model
```

## Required wrappers

Implement `Workbooks`, `Workbook`, `Worksheets`, `Worksheet`, and `Range`.

Each wrapper must own or safely share a `DispatchObject`, remain apartment-bound, expose a documented raw-dispatch escape hatch, avoid claiming stronger lifetimes than COM provides, and return structured errors.

## `ExcelApp`

Add:

```rust
app.workbooks()
app.run_macro(...)
```

Macro execution may remain scalar-only while array semantics are incomplete.

## `Workbooks`

Implement:

- `count`;
- `add`;
- `open` through a structured options builder or minimal explicit API;
- `item` where useful;
- indexed lookup;
- name lookup if officially supported.

Do not expose a huge positional `open` signature with loosely typed parameters.

## `Workbook`

Implement:

- `name`;
- `full_name`;
- `path`;
- `worksheets`;
- `save`;
- `save_as` with structured options;
- `close`;
- `saved`;
- raw dispatch access.

Document created and user-owned workbook lifecycles.

## `Worksheets`

Implement count, lookup by explicit 1-based index, lookup by name, add, and the approved iteration strategy.

Do not silently hide Excel's 1-based indexing without explicit documentation.

## `Worksheet`

Implement:

- `name` and `set_name`;
- `range("A1:B2")`;
- `cell(row, column)` or equivalent explicit coordinate API;
- `used_range` only with documented caveats;
- raw dispatch access.

## `Range`

Implement:

- `address`;
- row and column counts;
- scalar `value2` and setter;
- scalar `formula` and setter;
- `formula2` where supported;
- `clear_contents`;
- `offset`;
- `resize`;
- raw dispatch access.

For multi-cell ranges, scalar getters must return a structured array-required error or use the approved provisional value API.

## Options types

Prefer builder-like `OpenOptions`, `SaveAsOptions`, and `CloseOptions`. Map to named COM arguments where supported. Do not expose Excel enums as arbitrary integers in the normal API.

## Testing

Use deterministic mock-dispatch tests and opt-in real-Excel tests. Add a real-Excel example that creates a workbook, writes and reads `A1`, saves, closes, and quits. Ensure cleanup after intermediate failure.

## Acceptance workflow

A workflow equivalent to this must work:

```rust
let excel = ExcelApp::create()?;
let book = excel.workbooks().add()?;
let sheet = book.worksheets().item("Sheet1")?;
sheet.range("A1")?.set_value2("Hello")?;
let value = sheet.range("A1")?.value2()?;
book.save_as(path)?;
book.close(false)?;
excel.quit()?;
```

Commit, push, and open a draft PR.

Do not merge. Stop after reporting the draft PR.
