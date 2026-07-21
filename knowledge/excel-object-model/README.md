# Excel object-model knowledge base

This directory is a small, normalized evidence base for later Excel Automation
research. It is not a mirror of Microsoft Learn, a type library, or a future
Rust API. The checked-in records preserve concise implementation-relevant facts
and a link to each selected VBA-Docs source page; remarks, tutorials, and VBA
examples are intentionally not copied.

## Contents

- `SOURCE_MANIFEST.toml` is the authoritative pinned input revision and
  selection policy.
- `schema/` defines the version-1 JSON Schema contracts for each record type.
- `data/*.jsonl` contains stable, UTF-8, LF-terminated canonical objects,
  members, relationships, and enumerations. `data/source.json` records the
  deterministic ingestion coverage and review queue.
- `generated/` contains small reproducible indexes; never edit it by hand.
- `typelib/` is a separate, installed-Excel type-library evidence layer. It
  does not alter `data/` or claim runtime behavior; its source manifest is
  portable and deliberately records no raw user path.
- `generated/typelib/` contains deterministic reports derived from the
  type-library evidence layer; never edit it by hand.
- `runtime/` is a separate evidence layer for explicit owned-Excel experiments.
  It records version-specific observations or explicit host blockers; it never
  changes the documentation or typelib layers.
- `generated/runtime/` contains deterministic runtime summaries; never edit it
  by hand.

Every record carries its source repository, exact commit, repository-relative
path, extraction method, verification dimensions, implementation status, and
a field-level provenance map. `documentation: true` means only that an
official documentation page was selected. `typelib` and `runtime` remain
false throughout this initial data set.

## Stable identifiers

IDs derive from the selected page's `api_name` metadata, normalized by segment
without using file traversal order, headings, or page-title capitalization.
Examples include `Excel.Application`, `Excel.Workbook`,
`Excel.Range.Value2`, and `Excel.Workbooks.Open`. A page title must agree
case-insensitively with the metadata; a semantic disagreement is skipped and
reported rather than guessed. When the source assigns one API name to two
different documented member kinds, the records use a deterministic `#kind`
suffix (for example, `Excel.Chart.Activate#event`) and appear in the unresolved
classification report. Same-kind collisions are ingestion errors.

The collection flag is source-title evidence. An item type is recorded only
when a concise source statement supports it. It is not a claim about a default
member, enumeration support, indexing, or a public collection wrapper.

## Rebuild and verify

Normal validation is offline and never contacts Microsoft or Excel:

```powershell
cargo run --offline --manifest-path tools/excel-com-kb/Cargo.toml -- validate `
    --root knowledge/excel-object-model
cargo run --offline --manifest-path tools/excel-com-kb/Cargo.toml -- check `
    --root knowledge/excel-object-model
```

The type-library audit is a separate Windows-only, read-only operation. It
uses `LoadTypeLibEx` for an explicitly supplied local file or `LoadRegTypeLib`
for the registered Excel 1.9 library. It does not activate Excel or modify COM
registration:

```powershell
cargo run --offline --manifest-path tools/excel-com-typelib-audit/Cargo.toml -- audit `
    --root knowledge/excel-object-model `
    --typelib <path-to-EXCEL.EXE> `
    --windows-version <recorded-windows-version> `
    --excel-file-version <recorded-excel-file-version> `
    --office-bitness <recorded-office-bitness>
cargo run --offline --manifest-path tools/excel-com-typelib-audit/Cargo.toml -- check `
    --root knowledge/excel-object-model `
    --typelib <path-to-EXCEL.EXE> `
    --windows-version <recorded-windows-version> `
    --excel-file-version <recorded-excel-file-version> `
    --office-bitness <recorded-office-bitness>
```

To regenerate from the pinned source, supply a separate clean checkout. It is
deliberately neither vendored nor cached in this repository:

```powershell
cargo run --offline --manifest-path tools/excel-com-kb/Cargo.toml -- ingest `
    --source <path-to-vba-docs-checkout> `
    --manifest knowledge/excel-object-model/SOURCE_MANIFEST.toml `
    --output knowledge/excel-object-model/data
cargo run --offline --manifest-path tools/excel-com-kb/Cargo.toml -- generate `
    --root knowledge/excel-object-model
cargo run --offline --manifest-path tools/excel-com-kb/Cargo.toml -- analyze `
    --root knowledge/excel-object-model
git diff --check
```

The `ingest` command verifies the supplied checkout's Git `HEAD` against the
manifest when it is a Git checkout. Curated test fixtures intentionally omit a
Git directory and use a fixture manifest.

## Updating upstream

An update is a reviewed maintenance operation: obtain a separate upstream
checkout, inspect its `LICENSE` and `LICENSE-CODE`, change the manifest commit,
retrieval date, selection, and licence strings if needed, then ingest,
validate, generate, and review the complete diff and coverage report. Do not
use a branch tip as ordinary input or commit an upstream clone/archive. Record
new parser ambiguities instead of resolving them through undocumented runtime
assumptions.

For design rationale, limitations, and how later prompts should use the data,
see [the Prompt 02 research record](../../docs/research/excel-com/02-object-model-knowledge-base.md).
For the documentation-only object-model analysis generated from the same
canonical records, see [the Prompt 03 research record](../../docs/research/excel-com/03-excel-object-model-analysis.md).
For the installed-typelib audit and its runtime boundaries, see [the Prompt 04
research record](../../docs/research/excel-com/04-core-excel-typelib-audit.md).
For the explicit Range runtime probe and its current host-blocked result, see
[the Prompt 05 research record](../../docs/research/excel-com/05-range-variant-safearray-runtime.md).
