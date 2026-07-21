# Codex Prompt 01 — Research COM Automation Foundations

## Objective

Establish an authoritative, implementation-oriented description of the Windows COM Automation model required by an Excel client library.

This prompt is documentation and research only. Do not create the `excel-com` crate and do not modify production Rust code.

## Repository

Work in `evnekdev/excel-api`.

## Branch

Create exactly one branch:

```text
docs/excel-com-01-com-foundations
```

## Required preparation

Before editing:

1. Update from the latest `origin/master`.
2. Confirm that the worktree is clean.
3. Read:
   - `README.md`
   - `ARCHITECTURE.md`
   - `ARCHITECTURE_INDEX.md`
   - `COM_ARCHITECTURE.md`
   - `OPTIONAL_INTEGRATIONS_ROADMAP.md`
   - `IMPLEMENTATION_ROADMAP.md`
   - `docs/adr/0033-core-1-0-excludes-optional-integrations.md`
   - existing RTD research under `docs/research/`
4. Inspect current `windows` and `windows-core` usage in `examples/minimal-rtd-server`.

## Authority rules

Use authoritative Microsoft documentation as the primary source:

- Microsoft Learn Win32 COM documentation;
- Microsoft Learn Automation documentation;
- official documentation for `IUnknown`, `IDispatch`, `VARIANT`, `BSTR`, `SAFEARRAY`, `DISPPARAMS`, `EXCEPINFO`, and COM apartments.

Secondary sources may be used only to locate terminology or examples. Do not base architecture conclusions on blogs, Stack Overflow, generated AI content, or undocumented behavior.

Record source URLs directly in the research document. Summarize contracts rather than copying long passages.

## Required deliverable

Create:

```text
docs/research/excel-com/01-com-automation-foundations.md
```

## Required analysis

### COM identity and ownership

Explain:

- `IUnknown`;
- `QueryInterface`;
- `AddRef` and `Release`;
- interface identity;
- object identity versus interface pointer identity;
- ownership transfer conventions;
- how the `windows` crate manages COM interface reference counts;
- which operations remain unsafe even with generated bindings.

State the invariants a Rust wrapper must preserve.

### Automation interfaces

Explain:

- `IDispatch`;
- `GetTypeInfoCount` and `GetTypeInfo`;
- `GetIDsOfNames`;
- `Invoke`;
- DISPIDs;
- default members;
- method invocation;
- property get;
- property put;
- property put-reference;
- `DISPID_PROPERTYPUT`;
- positional and named arguments;
- reversed argument order in `DISPPARAMS`;
- optional arguments represented by missing variants;
- locale identifiers;
- `argErr`;
- deferred exception information where relevant.

Include pseudocode for a correct generic invocation sequence.

### Automation values

Describe:

- `VARIANT` and `VARTYPE`;
- scalar value types;
- `VT_EMPTY`, `VT_NULL`, `VT_ERROR`, `VT_DISPATCH`, `VT_UNKNOWN`, `VT_ARRAY`, and `VT_BYREF`;
- `BSTR`;
- OLE Automation dates;
- Currency;
- `VARIANT_BOOL`;
- ownership and cleanup;
- when `VariantClear` is required;
- how `windows::Win32::System::Variant::VARIANT` behaves.

Identify which values should become distinct Rust enum variants.

### Arrays

Describe:

- `SAFEARRAY`;
- dimensions and bounds;
- element type;
- descriptor ownership;
- data ownership;
- locking;
- element access;
- destruction;
- arrays of variants and arrays of primitives;
- non-zero lower bounds;
- multidimensional indexing conventions.

Do not make Excel-specific claims unless the official COM documentation establishes them.

### Errors

Explain:

- `HRESULT`;
- success and failure ranges;
- `DISP_E_*` errors;
- `EXCEPINFO`;
- member-not-found errors;
- type and argument mismatch;
- server exceptions;
- preserving HRESULT and structured exception fields in a Rust error type.

### Proposed internal boundary

End with a provisional, non-binding internal model:

```rust
ComApartment
DispatchObject
AutomationValue
AutomationArray
InvokeKind
InvokeArgs
AutomationError
```

For each type, state its responsibility and which raw COM details it hides. Do not finalize the public API.

## Required source audit

Create a table:

```text
Topic | Official source | Contract extracted | Open question
```

Every architecture-relevant conclusion must point to an official source.

## Open questions

End with a numbered list of questions requiring Excel-specific documentation, type-library inspection, or a real Excel experiment. Do not guess.

## Validation

Run:

```text
cargo fmt --all --check
cargo clippy --workspace --all-targets --all-features -- -D warnings
cargo test --workspace --all-features
cargo test --workspace --doc
```

Check all added Markdown links.

## Acceptance criteria

The prompt is complete when:

- the COM Automation substrate is described accurately enough to implement a generic dynamic Automation client;
- ownership and cleanup obligations are explicit;
- `IDispatch::Invoke` argument construction is unambiguous;
- no Excel-specific public API has been prematurely frozen;
- unresolved issues are clearly identified;
- all claims link to official Microsoft documentation.

## Completion

Commit the changes, push the branch, and open a draft pull request. The PR description must include scope, official sources, important findings, unresolved questions, and exact validation outcomes.

Do not merge. Stop after reporting the draft PR.
