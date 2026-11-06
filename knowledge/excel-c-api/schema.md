# C API symbol inventory record format

The canonical inventory is UTF-8 **JSON Lines** (`symbols.jsonl`): one independent object per record, one object per line, no comment lines. Current records belong to the pinned **Microsoft Excel 2013 XLL SDK, header 15.0** baseline unless a future reviewed source version is explicitly entered. This is architecture/data, **not generated Rust or a Cargo source file**.

## Required shape (schema version `1`)

| Field | Type | Rule |
| --- | --- | --- |
| `id` | string | Stable `CAPI-[0-9]{6}` identifier; never reuse, even after removal |
| `symbol` | string | Exact case-sensitive C declaration name as printed in the header |
| `kind` | string | One of `struct`, `type-flag`, `special-function-id`, `worksheet-function-id`, `macro-command-id`, `function-prototype`; expand enum in a reviewed schema change |
| `description` | string | One-sentence human explanation of the entry; distinguish header intent from verified runtime behavior |
| `source` | object | Immutable `repository`, `revision`, `path`, `blob_sha`, `line_start`, `line_end`; a single declaration anchor, 1-based inclusive |
| `edna_ids` | string[] | Zero or more unique `EDNA-01` through `EDNA-25` references; **related** capability, not implementation |
| `rust_status` | string | `unassessed`, `absent`, `partial`, `implemented`, `not-applicable` (based on direct current Rust evidence) |
| `rust_evidence` | string[] | Committed file+revision+symbol links where `rust_status` is assessed; empty for `unassessed` |
| `excel_host_status` | string | `not-run`, `pass`, `fail`, `blocked`; a host result requires an external linked run record |
| `excel_host_evidence` | string[] | Dated host/version/bitness/operation and outcome references; empty for `not-run` |
| `behavior_refs` | string[] | URLs or citation identifiers for documented/historical/tested behavior, each qualified in its owning evidence note; no unqualified behavior claims |

The catalogue intentionally has **no verified overall completeness metric**. The header version, exact provenance and verification history live with the records. A future multi-SDK catalogue may extend `source` to `sources` while preserving legacy IDs; do not insert a second record with the same canonical identity merely because an SDK revision changes.

## Version 2 source details

Version 2 retains the version-1 required fields and identities; each record explicitly sets `schema_version: 2`. The original 15 IDs and source anchors are unchanged. In addition to the six version-1 kinds, version 2 recognizes `scalar-typedef`, `pointer-typedef`, `struct-tag`, `function-pointer-typedef`, `ownership-flag`, `error-code`, `flow-code`, `return-code`, `event-id`, `function-modifier` and `other-constant`. The [working census report](../../docs/evidence/excel-c-api-source-census.md) explains the provisional counting convention and category limits.

| Field | Meaning when applicable |
| --- | --- |
| `expression`, `value` | Original object-like macro expression and safely evaluated integer; no implied C signedness. |
| `underlying` | Underlying type name of a scalar typedef. |
| `tag` | Struct tag for an aggregate typedef. |
| `target` | Corresponding aggregate typedef for a pointer typedef or struct tag. |
| `signature` | Whitespace-normalized prototype or callback typedef declaration. |
| `linkage` | Conditional C++ linkage context for prototypes. |
| `conditional_fields` | Guarded member spellings of XLOPER; nested members remain part of the aggregate. |
| `guard` | Active supported preprocessor guards on a macro candidate, if any. |
| `alias_of` | Explicit single-name macro alias target, if any; equal numbers alone do not imply an alias. |

For a v2 record, source details must match the candidate derived from its pinned source blob. These fields do not establish full aggregate field semantics, platform type widths, ABI layout or actual Excel behavior. There is no retirement or supersession representation yet; never remove or reuse an ID. Version 1 remains readable under its original, narrower checks.

## Invariants

1. **Identity:** case-sensitive `source.path + kind + symbol` is unique per canonical declaration across catalogue records. One ID can have multiple version variants once the schema explicitly models them; avoid source-version-based duplicate identities.
2. **Evidence:** no unpinned `master` or `main` source in records; every `source` has an immutable Git SHA and blob SHA, resolvable path and valid inclusive line span. Exact source text at that span must contain the declared name.
3. **Coverage:** `rust_status=implemented` requires direct Rust source or test evidence; `excel_host_status=pass` requires a real Excel run record and its host/bitness. **Never infer status from a constant existing in the C header**.
4. **One-to-many:** `edna_ids` is a *reference* list and can be empty, including for raw symbols not currently surfaced by Excel-DNA. All EDNA IDs must exist in the [canonical EDNA register](../../docs/architecture/excel-dna-subsystems-and-crates.md); IDs are not automatically crate names.
5. **Evolution:** record explicit source aliases and conditional guards where represented; two aliases with distinct printed names get distinct identities linked as aliases. Never discard historical evidence without explicit supersession.
6. **Source types:** the *core Excel C API* inventory excludes helper-only SDK framework headers, Excel-DNA native glue, Windows SDK, Office COM and .NET CLR host files. Track those in [header provenance #114](https://github.com/evnekdev/excel-api/issues/114), not as invented Excel C API symbols.
7. **Licensing:** only facts, identifiers, brief definitions and provenance pointers belong in this public catalogue. The copyrighted Dalton PDF belongs in private evidence storage, not the repository.

## Audit workflow

1. Consult the header provenance inventory in [#114](https://github.com/evnekdev/excel-api/issues/114) and the unresolved rights review in [#120](https://github.com/evnekdev/excel-api/issues/120).
2. Independently enumerate the SDK declarations by category (including macro vs worksheet numbers, preprocessor aliases, conditional symbols, duplicates and disabled/commented-out declarations).
3. Compare every candidate with `XLCALL.H` source, preserve lines and revision, assign next unused ID, then run JSONL, duplicate-ID/identity, source-span and EDNA-link checks.
4. Separate raw ABI fidelity tests from safe wrapper legality and native Excel acceptance; open scoped leaf issues if there are real implementation gaps.
5. Record the reviewed per-category denominators and coverage separately **only after** the exhaustive extraction reconciles with the source.

## Sample record

This historical version-1 example shows the seed identity before its version-2 source-detail enrichment; Rust/Excel runtime status remains unassessed:

```json
{"id":"CAPI-000005","symbol":"xltypeNum","kind":"type-flag","description":"Excel value type bit used to identify a numeric XLOPER value.","source":{"repository":"evnekdev/excel-api","revision":"6c8354d6536e18b58559f9cc3134f25a9e58d478","path":"tools/Excel2013XLLSDK/INCLUDE/XLCALL.H","blob_sha":"4eb50e8fd3aea40d283181b7dcc74871ba88d854","line_start":243,"line_end":243},"edna_ids":["EDNA-02","EDNA-03"],"rust_status":"unassessed","rust_evidence":[],"excel_host_status":"not-run","excel_host_evidence":[],"behavior_refs":[]}
```

See [architecture](../../docs/architecture/c-api-inventory-layer.md) for the separate semantic and Excel-DNA inventories.

## Local validation

The [validator](../../tools/excel-c-api-inventory/README.md) accepts v1 and v2. It checks v2 source details and supported macro values against the pinned source, but does not establish general C semantics, ABI fidelity, full preprocessor coverage or actual host runs. It rejects host `pass` claims pending a verifiable evidence format. All commands check identities and removals against reachable catalogue history from the current branch and `origin/master`; missing required history is reported as unverified, not proof of stable identity.
