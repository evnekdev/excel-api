# Codex Prompt 09 — Implement Excel Activation and `ExcelApp`

## Objective

Implement creation and attachment of Microsoft Excel, together with explicit process-ownership and shutdown semantics.

## Repository and branch

Repository: `evnekdev/excel-api`

Branch:

```text
feature/excel-com-09-application
```

## Required scope

Implement:

```rust
ExcelApp::create()
ExcelApp::attach_active()
```

according to the approved architecture.

### Creation

Use the approved COM activation path for `Excel.Application`.

Handle ProgID/CLSID lookup, `CoCreateInstance`, interface conversion, activation errors, Excel not installed, and bitness/registration diagnostics where detectable.

### Attachment

Use the approved active-object mechanism. Document limitations when no active instance exists, several instances exist, the ROT exposes only one representative object, or process-specific attachment is unsupported.

### Ownership mode

Represent created and attached instances explicitly. A created instance may use convenience cleanup policies. An attached instance must never be silently terminated.

Do not call `Application.Quit` unconditionally from `Drop`.

### Initial application API

Implement:

- `visible` and `set_visible`;
- `version`;
- `display_alerts` and setter;
- `screen_updating` and setter;
- `enable_events` and setter;
- calculation-mode access only if authoritative enum mapping exists;
- `quit`;
- raw `DispatchObject` access.

### Drop behavior

Dropping `ExcelApp` must release the COM reference, avoid panicking, avoid indefinite blocking, avoid unexpected user-visible actions, preserve attached instances, and surface cleanup failures only through explicit methods.

### Quickstart example

Implement an example that creates Excel, prints the version, makes Excel visible, and explicitly calls `quit`.

## Real-Excel validation

Add an opt-in test or script recording Excel version/build, creation, visibility, quit, remaining process state, and attachment behavior.

Do not make licensed Excel a normal CI requirement. Do not use process killing as normal cleanup.

## Acceptance

- creation works on supported Windows desktop Excel;
- attachment works with an active instance;
- attached instances are preserved;
- explicit quit works for a clean created instance;
- errors retain HRESULT and operation context;
- no broad object-model wrappers are added yet.

Commit, push, and open a draft PR.

Do not merge. Stop after reporting the draft PR.
