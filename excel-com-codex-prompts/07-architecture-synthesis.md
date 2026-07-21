# Codex Prompt 07 — Synthesize the `excel-com` Architecture

## Objective

Convert reviewed research into stable architecture decisions, ADRs, crate boundaries, milestones, and acceptance criteria.

Do not implement production COM calls in this prompt.

## Repository and branch

Repository: `evnekdev/excel-api`

Branch:

```text
architecture/excel-com-07-synthesis
```

## Required reading

Read:

- all files under `docs/research/excel-com/`;
- `ARCHITECTURE.md`;
- `ARCHITECTURE_INDEX.md`;
- `COM_ARCHITECTURE.md`;
- `EXCELDNA_CAPABILITY_MAP.md`;
- `OPTIONAL_INTEGRATIONS_ROADMAP.md`;
- `IMPLEMENTATION_ROADMAP.md`;
- ADR-0033;
- current workspace manifests;
- core release documents.

Research documents are evidence. Architecture documents must contain explicit decisions.

## Required files

Create:

```text
EXCEL_COM_ARCHITECTURE.md
EXCEL_COM_CAPABILITY_MAP.md
EXCEL_COM_ROADMAP.md
docs/adr/0034-excel-com-automation-client-boundary.md
docs/adr/0035-excel-com-dynamic-dispatch-strategy.md
docs/adr/0036-excel-com-apartment-and-lifecycle-model.md
```

Update:

```text
ARCHITECTURE_INDEX.md
OPTIONAL_INTEGRATIONS_ROADMAP.md
IMPLEMENTATION_ROADMAP.md
README.md
```

Prefer documentation-only scope. Do not add production crate code.

## Mandatory decisions

### Product boundary

State that `excel-com` is:

- an Excel COM Automation client;
- Windows desktop focused;
- independently versioned;
- independent of `excel-api`, `excel-api-sys`, and `excel-api-macros`;
- not RTD, Ribbon, or a custom task pane;
- not a replacement for file-format crates;
- not supported for unattended service automation.

### Dependency graph

Define:

```text
windows / windows-core
        ↓
excel-com
        ↓
user automation program
```

No core XLL crate may depend on `excel-com`. Any future XLL integration must use a separate opt-in bridge crate.

### Internal layers

Freeze responsibilities for:

- COM apartment/lifecycle;
- Automation value conversion;
- dynamic dispatch;
- Excel-specific typed wrappers;
- reliability and retry policy;
- optional integrations.

### Threading

Decide supported apartment model, `!Send`/`!Sync`, message-pump expectations, thread-crossing rules, future dedicated STA worker, and marshaling policy.

### Lifecycle

Decide created versus attached instances, explicit `Quit`, drop behavior, child lifetimes, leak diagnostics, setting guards, and busy-call retry policy.

### Values

Freeze the conceptual distinction between:

```rust
AutomationValue
ExcelValue
RangeValues
```

Define scalar and matrix semantics without freezing every convenience conversion.

### Dispatch strategy

Decide dynamic `IDispatch` first, permanent raw escape hatch, DISPID caching, generated metadata deferred, and fully generated interfaces deferred unless evidence justifies them.

### Public release scope

Define `0.1` support for:

- start and attach;
- visibility and application settings;
- workbook add/open/save/close;
- worksheet lookup/add;
- range lookup;
- scalar and rectangular values;
- formulas;
- macro execution;
- raw dispatch;
- structured errors;
- deterministic cleanup.

Explicitly defer events, complete generated interfaces, broad charts/pivots, macOS, remote DCOM, server automation, and automatic XLL integration.

## Roadmap

Create milestones:

```text
EC0 Research and architecture
EC1 Crate skeleton and COM apartment
EC2 Dynamic dispatch kernel
EC3 Excel activation and lifecycle
EC4 Workbook/Worksheet/Range vertical slice
EC5 Rectangular value transport
EC6 Reliability and guards
EC7 Broader object model
EC8 Events
EC9 Type-library tooling
EC10 Optional excel-api bridge
```

For each milestone include scope, prerequisites, deliverables, tests, real-Excel validation, non-goals, and acceptance gate.

## Acceptance criteria

Implementation prompts must not need to invent ownership, threading, lifecycle, dispatch semantics, error boundaries, value representation, dependency direction, or support claims.

## Completion

Run documentation and workspace checks. Commit, push, and open a draft PR.

Do not merge. Stop after reporting the draft PR.
