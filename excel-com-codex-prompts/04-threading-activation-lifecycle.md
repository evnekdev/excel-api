# Codex Prompt 04 — Research Apartments, Activation, and Excel Lifecycle

## Objective

Define the threading, apartment, activation, process-ownership, and shutdown model for `excel-com`.

This prompt is research and architecture evidence only.

## Repository and branch

Repository: `evnekdev/excel-api`

Branch:

```text
research/excel-com-04-threading-lifecycle
```

## Required sources

Use official Microsoft documentation for:

- `CoInitializeEx` and `CoUninitialize`;
- STA and MTA rules;
- COM marshaling and the Global Interface Table where relevant;
- Running Object Table;
- `GetActiveObject`;
- `CoCreateInstance`;
- COM message loops;
- `IMessageFilter` and rejected calls;
- Microsoft Office threading guidance;
- Microsoft guidance on server-side Office Automation;
- Excel `Application.Quit`;
- workbook close and save prompts.

## Required deliverable

Create:

```text
docs/research/excel-com/04-threading-activation-and-lifecycle.md
```

## Required analysis

### Apartment model

Determine:

- the initial supported apartment model;
- whether caller-owned STA is supported;
- whether crate-owned STA is supported;
- consequences of `CoInitializeEx` returning `S_OK`, `S_FALSE`, or `RPC_E_CHANGED_MODE`;
- balancing `CoUninitialize`;
- message-loop requirements;
- whether COM objects should be `!Send` and `!Sync`;
- what may safely cross thread boundaries;
- how marshaling would work if later required.

Provide explicit Rust invariants.

### Activation modes

Analyze:

```rust
ExcelApp::create()
ExcelApp::attach_active()
```

For each mode determine the underlying COM operation, likely instance identity, ownership expectations, drop/quit behavior, multiple-instance limitations, and future attach-by-window/process options.

### Process lifecycle

Investigate:

- whether `CoCreateInstance` creates a new process or may reuse one;
- hidden Excel instances;
- open workbooks and unsaved prompts;
- `DisplayAlerts`;
- child COM references keeping Excel alive;
- release order for `Range`, `Worksheet`, `Workbook`, and `Application`;
- orphan `EXCEL.EXE` processes;
- explicit `Quit`;
- preserving attached instances;
- Excel shutdown while references remain.

Define a provisional created-versus-attached ownership model.

### Rejected and busy calls

Document:

- `RPC_E_CALL_REJECTED`;
- `RPC_E_SERVERCALL_RETRYLATER`;
- `IMessageFilter`;
- `RetryRejectedCall`;
- message pumping;
- retry delays;
- bounded timeout;
- cancellation;
- modal dialogs;
- why arbitrary sleeps are insufficient.

Propose, but do not implement, a `RetryPolicy`.

### Reentrancy

Analyze nested calls, events, calls back into Excel from handlers, locks held across COM calls, deadlock hazards, and interaction with the existing project rule against holding locks while calling Excel.

### Unsupported environments

State the support boundary for Windows desktop Excel, interactive sessions, Windows services, unattended server automation, Excel Online, macOS, Wine, and remote DCOM.

Do not make unsupported claims.

## Experimental evidence

Where possible, create narrow probes for creating Excel, attaching, repeated create/quit cycles, dropping without `Quit`, retaining child objects after parents, closing unsaved workbooks, and detecting remaining Excel processes.

Record exact environment and results.

## Required recommendations

Conclude with recommendations for:

- default apartment;
- `!Send`/`!Sync`;
- caller-owned versus crate-owned apartment;
- explicit versus automatic `Quit`;
- created versus attached policy;
- retry-policy location;
- safe drop behavior;
- future worker-thread abstraction.

## Validation and completion

Run relevant checks. Commit, push, and open a draft PR.

Do not merge. Stop after reporting the draft PR.
