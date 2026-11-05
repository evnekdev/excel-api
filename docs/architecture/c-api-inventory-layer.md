# Canonical Excel C API inventory layer

**Status:** architecture established; raw symbol catalogue is a deliberately incomplete, source-verified seed. Do **not** use it as a percentage-of-SDK or implemented-feature score.

Tracking: [architecture #116](https://github.com/evnekdev/excel-api/issues/116), [programme #78](https://github.com/evnekdev/excel-api/issues/78), [EDNA-02 #90](https://github.com/evnekdev/excel-api/issues/90). Header-evidence completion is separately tracked by [#114](https://github.com/evnekdev/excel-api/issues/114); the Dalton behavior-book evidence index by [#115](https://github.com/evnekdev/excel-api/issues/115).

## Why this sits below Excel-DNA

**The Microsoft Excel C API is the contract; Excel-DNA is an ecosystem built on top of it.** We therefore inventory the raw C API independently of which definitions Excel-DNA chose to expose. Excel-DNA's managed declarations and wrappers provide *implementation and behavioral-comparison evidence*, not primary ABI authority. Rust implementation coverage is a third, independently assessed dimension.

The layers and links are:

```text
Pinned Microsoft SDK headers and official documented contracts
    => knowledge/excel-c-api/symbols.jsonl  (CAPI-xxxxxx, canonical raw declarations)
    => behavior evidence (Microsoft docs, versioned book sections, native Excel probes)
    => EDNA-01..25 subsystem capabilities (many-to-many references; richer framework features)
    => excel-api-sys / excel-api / excel-api-macros (per-symbol mapping and safe API assessment)
    => installed PyroApp acceptance (separate, consumer-specific proof)
```

This is an **evidence/data dependency**, not a crate dependency, a GitHub parent/child chain for every symbol, or a claim that all EDNA capabilities arise from the C API. Office Ribbon, RTD and Office-specific COM protocols have independent interfaces; managed CLR hosting is not an Excel C API capability.

## Authoritative evidence order and limits

| Evidence | Purpose | Where / status |
| --- | --- | --- |
| **Microsoft Excel XLL SDK** | Exact C declarations, constants, structures, callback prototypes and conditional compilation as published in that SDK | Existing `tools/Excel2013XLLSDK/INCLUDE/XLCALL.H` (version 15.0); **raw primary anchor** |
| **Official Microsoft C API documentation** | Supported operation, availability, restrictions and version-specific behavior | Link individually from behavior records; do not assume current docs and 2013 header are identical |
| **Excel-DNA** | Actual higher-level wrappers, registration, lifetime and support-code decisions | Pinned source `Excel-DNA/ExcelDna` `f484749260d5cc43c35f7aea5f7a0b0f8f74a78f`; supplemental |
| **Steve Dalton (Wiley, 2nd ed., 2007)** | Historical practical cases, pitfalls, crash patterns, test scenarios | Copyrighted private project Library copy; bibliographic/evidence register pending [#115](https://github.com/evnekdev/excel-api/issues/115) |
| **Rust source and native host** | Implementation existence, ABI correctness and observed runtime behavior | Current source/revision and real-Excel evidence must be recorded separately for each assessed entry |

The checked-in Microsoft SDK also includes sample-only headers, callback-bridge source, import libraries, documentation and an EULA. The **full header provenance register remains a separate outstanding task #114**. Do not import additional proprietary headers into the public repository as an incidental inventory step. The book is not an authoritative definition of modern Microsoft 365 behavior.

## Canonical ownership and identifiers

- **Canonical raw catalogue:** [`knowledge/excel-c-api/symbols.jsonl`](../../knowledge/excel-c-api/symbols.jsonl), format in [`schema.md`](../../knowledge/excel-c-api/schema.md). Stable `CAPI-000001`-style IDs are permanent; never renumber or reuse one. Later audited declarations append IDs even if source order differs.
- **Identity:** `(normalized canonical header path, symbol kind, exact case-sensitive symbol name)`; store version/guard variants under the *same* identity with distinct sources when necessary. Preserve alias and supersession links rather than silently combining distinct meanings.
- **Reference relationship:** `CAPI-xxxxxx` => zero or more `EDNA-xx` subsystem references. `EDNA-xx` => zero or more CAPI entries *plus* additional non-C-API capabilities. Many-to-many links do **not** imply that a symbol or an EDNA feature is implemented.
- **GitHub issue ownership:** the [EDNA-02 issue #90](https://github.com/evnekdev/excel-api/issues/90) owns raw ABI and declaration completeness; functional wrappers and behavioral acceptance belong under their appropriate EDNA owners. Share symbol IDs and use related/blocked-by references; avoid two primary parents or duplicate implementation.
- **Revision rules:** retain immutable source revision + blob SHA + physical line range for the source baseline; if the source revision changes, retain the historical anchor and add a reviewed replacement/version source. Do not automatically overwrite an older source anchor or mark removed identifiers as never having existed.

## What the inventory will eventually cover

1. **C types and aggregate ABI:** scalar aliases, `XLOPER/XLOPER12`, `FP/FP12`, references, strings, handles and memory representation.
2. **Constants and identifiers:** type/error/return/ownership/permission flags, `xlf*` function identifiers, `xlc*` macro commands, `xl*` auxiliary callbacks, event IDs, registration codes, availability and numeric aliases.
3. **Native prototypes/exports:** `Excel4/4v/12/12v`, `XLCallVer`, callbacks and XLL lifecycle export contracts (document separately when not literal declarations in `XLCALL.H`).
4. **Behavioral contracts:** legal caller context, allowed argument/result types, memory ownership, thread safety, macro-type permissions, cancellations, scheduling, recalculation, version guards and error/abort behavior.
5. **Support or non-Excel layers:** Windows SDK, optional sample-framework headers, Office/COM interfaces, CLR host headers and build tools tracked in source provenance but **never counted as core C API declarations**.

The raw inventory is based on the declared SDK surface, **not** Excel-DNA's coverage, and does not require a high-level safe wrapper for every numeric constant. Coverage denominators are valid only after a pinned-source exhaustive extraction and a reviewed completeness reconciliation.

## Three independent progress measures

| Dimension | Meaning | Do not confuse it with |
| --- | --- | --- |
| **Source inventory** | Declaration accounted for in a pinned source baseline, with a reviewed identity | An imported constant being callable |
| **Rust implementation** | Matching raw declaration / safe abstraction implemented, documented, and tested where applicable | Permission to call it from any callback |
| **Excel host verification** | Specific operation exercised and passed on an identified Excel release and x86/x64 architecture | A mocked test, SDK example or Excel-DNA implementation |

The schema uses **`unassessed`** by default for Rust and host coverage. For behavior claims, distinguish **official/documented**, **historical/book**, **upstream-observed**, **host-tested** and **open hypothesis**; add a version or host environment for each. If the book raises a conflict with Microsoft documentation or observed host behavior, create an investigation issue, do not silently choose one.

## Work sequence

- **[#114](https://github.com/evnekdev/excel-api/issues/114):** finalize exact primary/supporting C/C++ header list, hashes, include dependencies, license disposition.
- **[EDNA-02 exhaustive audit #117](https://github.com/evnekdev/excel-api/issues/117):** enumerate all actual declarations in the pinned `XLCALL.H`, including aliases, guards, duplicate spellings and collisions; compare an independently reviewed count against the source. Mark any parser limitations. Do not silently infer symbols from names or sample code.
- **EDNA behavior work:** annotate the inventory using official references and page-precise citations from Dalton; open investigations for unknown legality or surprising scenarios.
- **EDNA coverage work:** map independently assessed C API entries and higher-level Excel-DNA capabilities to exact Rust implementation paths and host tests, then create meaningful leaf issues.
- **Product acceptance:** prove PyroApp parity separately; no reduction in its supported modes, exports, UI or bitness.

**Current acceptance:** architecture and a limited source-verified seed only. No complete C API coverage, Rust implementation completeness, x86 acceptance or real-Excel validation is asserted here.
