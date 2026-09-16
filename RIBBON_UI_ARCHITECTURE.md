# Ribbon UI Architecture

Status: **roadmapped optional post-1.0 integration; not implemented**.

Ribbon customization is not required for ordinary native Excel worksheet
functions, commands, lifecycle behavior, asynchronous UDFs, or cooperative
dispatch. It has a distinct Office UI callback, XML/metadata, deployment,
trust, and compatibility contract from the Excel 12 C API core.

The concrete implementation milestone is H5 in
`docs/architecture/native-xll-hosting-roadmap.md`. This document records the
Ribbon-specific architectural rules that H5 must preserve.

## 1. Crate and dependency boundary

Ribbon support should live in an optional Windows/Office integration crate,
working name `excel-api-ribbon`.

The normal `excel-api`, `excel-api-sys`, `excel-api-macros`, and minimal XLL
must not acquire an unconditional Office COM dependency, COM registration side
effect, or Ribbon packaging requirement.

The Ribbon crate may depend on the reviewed generic COM foundation from
`docs/architecture/com-automation-roadmap.md` and/or `windows-core` where the
latter materially reduces unsafe vtable code. It must not create a second
permanent manual implementation of `IUnknown`/`IDispatch` ownership if the
generic COM layer already provides the required semantics.

## 2. Office contract must be pinned before implementation

Before production code is written, audit the installed/current Office type
library and authoritative Microsoft documentation for the exact interfaces and
registration model required by the selected topology.

The evidence set must pin at least:

- the interface that supplies Ribbon XML/custom UI;
- the Office add-in connection/disconnection contract required by Excel;
- callback `IDispatch` method discovery/invocation rules;
- `IRibbonUI` lifetime and invalidation methods;
- callback parameter/return Automation types used by supported controls;
- registration keys/values, bitness and per-user/per-machine behavior;
- load, unload and Excel-shutdown ordering.

Common Office COM interface names may guide research, but the public Rust API
must not be frozen from memory or examples alone.

## 3. Ribbon callback capability is not Excel C API capability

A Ribbon callback can arrive on Excel's UI thread without being an Excel 12 C
API callback. Thread identity therefore does not create `WorksheetContext`,
`MacroContext`, or any other `excel-api` capability.

Hard rule:

> Ribbon callbacks may perform only operations justified by the Ribbon/COM
> callback contract. Work requiring Excel C API macro capability must be queued
> through the H3 wake adapter and executed only after Excel enters a registered
> XLL helper command that yields a genuine `MacroContext`.

This preserves the existing project rule that callback capability is stronger
than "runs on the main thread".

## 4. Artifact topology is an evidence-driven decision

Do not assume Ribbon code belongs in the same binary as the XLL merely because
both run in Excel.

H5 must compare at least:

### 4.1 Combined binary

One native module exports XLL lifecycle/function symbols and also participates
in the Office COM add-in/class-server lifecycle.

Advantages:

- one application module and potentially direct shared Rust state;
- fewer deployed binaries.

Risks to prove:

- XLL unload and COM reference-count unload protocols must reach one consistent
  quiescent state;
- Office may retain COM references after `xlAutoClose` begins;
- background/queued callbacks must not target code after Excel unloads XLL
  exports;
- class-factory/server lock state must agree with XLL runtime generations.

### 4.2 Companion Ribbon DLL

A separate Rust COM add-in DLL owns Ribbon lifetime and communicates with the
XLL through registered commands and/or a deliberately narrow process-local
coordination channel.

Advantages:

- XLL and COM unload protocols stay separate;
- failure and registration boundaries are easier to audit initially.

Costs:

- application state cannot be shared through unrelated Rust statics;
- a narrow coordination contract is required for live status/query callbacks;
- packaging contains another binary.

The first production topology should minimize lifetime ambiguity, not artifact
count. Combining binaries later is acceptable only after reload/shutdown tests
prove the shared lifecycle.

## 5. Callback surface

The public consumer API should be typed and metadata-driven rather than a raw
`IDispatch::Invoke` switch written by every application.

Conceptually, application code should be able to define callbacks for common
Ribbon shapes such as:

- `onLoad` / UI handle capture;
- action/click callbacks;
- boolean enabled/visible/pressed queries;
- string label/text/status queries;
- selection/index callbacks for supported controls;
- explicit invalidation requests.

Exact supported signatures are determined by the audited Office contract.
Unsupported callback shapes fail at build time or registration time rather than
silently coercing arbitrary Automation values.

Every COM/vtable/callback boundary contains panics and maps failures to a
controlled HRESULT/fallback appropriate for that Office callback.

## 6. `IRibbonUI` ownership and invalidation

The `onLoad` callback supplies the UI object used for later invalidation. The
Rust wrapper must:

- own exactly the required COM reference;
- remain apartment-correct;
- never become `Send`/`Sync` merely for convenience;
- invalidate controls only while the Ribbon generation is active;
- become safely unusable after disconnect/shutdown;
- release its reference in the correct lifecycle phase.

Application code should request whole-Ribbon or named-control invalidation
through typed methods rather than exposing a raw COM pointer.

## 7. Generated Ribbon contract

Large add-ins should not maintain Ribbon XML, callback names and Rust callback
registration as three independent sources of truth.

A preferred consumer pattern is a deterministic application contract that can
generate:

- Ribbon XML;
- callback method metadata;
- compile-time callback signature checks;
- control IDs and duplicate detection;
- installer registration metadata where appropriate;
- acceptance-test inventory.

PyroApp's existing package-contract generator is the first concrete consumer:
its future Rust path should emit Rust callback metadata rather than generated
C# while preserving the same package-contract validation principle.

## 8. Relationship to current-host Application access

Ribbon code often needs the current Excel object model. It should use the exact
host bridge designed in H2 rather than generic ROT attachment or creation of a
new Excel process.

Where possible, long-lived Ribbon state should remain owned Rust data.
Apartment-bound Excel COM wrappers should be reacquired or retained only under
a documented apartment/lifetime rule; they must not be sent to background
threads.

## 9. Relationship to the macro wake adapter

Ribbon actions that need delayed/continued macro work use H3:

```text
Ribbon callback
    |
    | enqueue owned application task
    v
H3 wake queue
    |
    v
Excel invokes registered XLL helper command
    |
    v
MacroContext
    |
    v
application macro handler
```

This is especially important for workflows that must yield to Excel before
calculation or dependency processing can occur.

The Ribbon layer must not create a second independent main-thread scheduling
system.

## 10. Registration, trust and installer behavior

Ribbon COM registration is an installer concern separate from XLL
`xlfRegister` function registration.

The production design must record and test:

- ProgID/CLSID or other required Office add-in identity;
- exact server path and bitness;
- per-user versus per-machine registration scope;
- Office add-in load behavior;
- signing/trust state;
- deterministic uninstall/rollback;
- conflict detection rather than silently overwriting another registration.

Testing must never add global Trusted Locations, weaken macro/XLM security, or
modify organization-wide Office policy.

## 11. Validation matrix

Unit/native fixture tests should cover:

- metadata generation and duplicate callback/control rejection;
- `IUnknown`/`IDispatch` identity and reference counts;
- callback argument coercion;
- panic containment;
- UI-generation invalidation and stale-handle rejection;
- shutdown state transitions;
- installer metadata generation.

Real Excel tests should cover:

- clean load with the chosen registration scope;
- Ribbon XML render;
- action callbacks;
- query callbacks;
- `IRibbonUI` invalidation;
- callback -> H3 macro handoff;
- repeated load/unload and Excel restart;
- uninstall/reinstall;
- no callback after unload;
- no leaked Excel process or COM server reference;
- coexistence with ordinary XLL functions and native async UDFs.

UI Automation may be used as an acceptance harness to click actual controls,
but passing UI Automation is not a substitute for COM lifetime tests.

## 12. First acceptance consumer

PyroApp is the intended first substantial acceptance consumer because its
current Ribbon controls application-level connection state, UDF execution mode,
derivative calculation, documentation, migration, diagnostics and About UI.

The framework should support those control/callback shapes without importing
PyroApp-specific settings, worker, ChemApp or derivative logic into
`excel-api-ribbon`.

H5 is complete only when the generic Ribbon layer has real Excel evidence and a
consumer such as PyroApp can reproduce its required control behavior without
Excel-DNA.