# Native XLL hosting and pure-Rust add-in roadmap

Status: architecture decision and staged implementation roadmap.

Baselines reviewed:

- `excel-api` `origin/master` at `8d4def8a9117f1cc56965935f6644640bc5d1616`
  (`docs: define com-automation extraction roadmap`);
- `pyroapprs` `origin/master` at `b9eb9d2d4d75181421a4059f4e42b7f2bd1278b1`
  (`test: make installer scope acceptance deterministic`).

This document turns the current PyroApp/Excel-DNA replacement analysis into a
framework roadmap. PyroApp is the first demanding acceptance consumer, but the
new abstractions must remain useful to unrelated native Rust XLLs.

This roadmap is **optional post-core-1.0 work**. It does not reopen the core
Excel 12 ABI boundary, does not make Ribbon or general COM a core 1.0
requirement, and does not weaken the existing callback-capability or ownership
rules.

Related architecture:

- `MAIN_THREAD_DISPATCH_ARCHITECTURE.md`;
- `ASYNC_ARCHITECTURE.md`;
- `COM_ARCHITECTURE.md`;
- `RIBBON_UI_ARCHITECTURE.md`;
- `docs/architecture/com-automation-roadmap.md`;
- `docs/architecture/excel-com-migration-support.md`.

## Executive decision

A pure-Rust replacement for a substantial Excel-DNA application such as
PyroApp is feasible without reproducing all of Excel-DNA. The existing
`excel-api` core already covers the difficult worksheet ABI pieces: XLL
lifecycle, registration, generated thunks, callback-borrowed inputs, owned
returns, arrays, references, typed callback capabilities, commands, native
asynchronous UDF handles, and the cooperative main-thread dispatcher.

The remaining framework gap is **host integration**, not worksheet-function
marshaling. A production pure-Rust add-in additionally needs:

1. small registration/location conveniences used by large add-ins;
2. access to the `Excel.Application` object belonging to the **current hosting
   Excel process**, without starting or arbitrarily attaching to another Excel;
3. an autonomous but capability-correct way to ask Excel to execute queued
   work as a genuine registered macro callback;
4. optional Office Ribbon COM integration;
5. a separately validated decision for applications that currently rely on
   Excel-DNA's RTD-style `ExcelAsyncUtil.Run` semantics.

The recommended dependency direction is:

```text
application XLL (for example pyroapprs-xll)
        |
        +---------------------> excel-api
        |                       Excel 12 ABI, values, registration,
        |                       lifecycle, async handles, dispatcher
        |
        +---------------------> excel-api-desktop      [new, optional]
        |                       current host Application bridge,
        |                       autonomous macro wake adapter
        |                         |              |
        |                         |              +--> excel-api
        |                         +------------------> excel-com
        |
        +---------------------> excel-api-ribbon       [new, optional]
        |                       Office Ribbon COM adapter
        |                         |
        |                         +--> com-automation / windows-core
        |
        +---------------------> application-specific Rust
                                worker pool, IPC, protobuf, settings,
                                DAT parsing/editing, gRPC, dialogs

excel-com --> com-automation --> Windows COM/OLE
```

Crate names are preferred working names, not frozen public API. The important
boundary is that `excel-api` core must not grow a dependency on the Excel COM
object model merely because desktop applications need one.

## 1. PyroApp as the acceptance consumer

The current PyroApp C# layer exercises several distinct facilities. They must
not be treated as one migration problem.

### 1.1 Already covered by `excel-api`

The following Excel-DNA roles already have a native Rust framework equivalent:

- worksheet-function and command registration;
- scalar, string, array and general Excel values;
- optional/missing input handling through the supported closed type mapping;
- Excel references where a preserving `U` argument is actually required;
- DLL-owned return allocation and `xlAutoFree12` cleanup;
- callback panic containment;
- lifecycle open/close and registration cleanup;
- typed worksheet/thread-safe/macro/lifecycle callback capabilities;
- native Excel async `>`/`X` registration and `xlAsyncReturn` completion;
- a bounded cooperative dispatcher with generation and shutdown control.

Therefore ordinary PyroApp functions such as runtime information, list/get/set
families, data transforms and optimization helpers do not justify a new
framework subsystem. They are primarily an application-language port.

### 1.2 Small framework gaps

Large add-ins need several facilities that should be added to the normal XLL
surface rather than reimplemented by every consumer:

- `help_topic` in generated/static function registration metadata;
- a process-lifetime owned XLL path and XLL directory;
- a typed callback helper for the verified active-workbook directory query
  currently performed through `xlfGetDocument` selector 2;
- an ergonomic generated or explicit large `AddInDescriptor` construction path
  so dozens of annotated functions do not require hand-maintained parallel
  metadata lists.

These additions are independent of COM/Ribbon work and should remain usable by
headless/minimal XLLs.

### 1.3 Host COM requirements

PyroApp's derivative-matrix command is the most useful acceptance scenario
because it needs the actual host Excel object model rather than a second Excel
instance. Its current contract needs, among other things:

- current `Application` and active worksheet;
- `Range` lookup, `Value2`, row/column/cell counts, resize and clear;
- `Application.Calculation` and `CalculationState`;
- worksheet calculation;
- `Application.Intersect`;
- `Application.StatusBar`;
- `Application.ScreenRefresh`;
- a way to yield to Excel and later resume in a genuine macro callback.

Most Range/calculation operations already exist in `excel-com`. The main new
framework abstraction is current-host acquisition; a few metadata-only Excel
members need typed wrappers.

### 1.4 Application-specific responsibilities that stay out of `excel-api`

The following current PyroApp responsibilities remain application code:

- isolated ChemApp worker process pool and restart policy;
- named-pipe framing and request correlation;
- protobuf schemas and generated messages;
- local/remote calculation transport selection;
- future gRPC client/server behavior;
- ChemApp runtime configuration and worker scratch directories;
- DAT parsing/editing and parser policy;
- PyroApp settings and execution-generation semantics;
- derivative numerical algorithm and B1:B6 workbook contract;
- progress/About/error dialogs;
- installer product policy.

Once the XLL itself is Rust, PyroApp should call its Rust DAT/parser crates
directly. The existing C ABI DLL may remain for non-Rust consumers, but it
should not be retained merely to preserve a C# P/Invoke boundary that no longer
exists.

## 2. Small core additions

### 2.1 Registration help topic

`FunctionRegistration` and `#[excel_function]` should gain an optional help
URL/topic that is passed to `xlfRegister` in the documented position. This is
part of the registration contract, not a Ribbon feature.

Conceptual syntax:

```rust
#[excel_function(
    name = "XLL_CA_VERSION",
    thunk = "xll_ca_version",
    category = "PYROAPP",
    description = "Returns the ChemApp version.",
    help_topic = "https://example.invalid/functions/ca-version/"
)]
fn ca_version() -> ExcelValue { /* ... */ }
```

Compile-time validation should reject embedded NULs and whatever length limit
the existing counted-string registration path already enforces. Empty strings
should normalize to absence.

### 2.2 XLL location

A successful runtime open already has enough information to identify the
loaded XLL. The runtime should retain an owned canonical-or-best-effort path
with no callback borrow and expose read-only process-lifetime access such as:

```rust
runtime.xll_path()
runtime.xll_directory()
```

The exact naming is not frozen. The API must distinguish "path unavailable"
from an empty path and must not depend on the process current directory.

### 2.3 Active workbook directory

Add a typed semantic helper around the verified `xlfGetDocument` directory
query. Do not expose an unrestricted numeric-selector API merely because the
underlying Excel function accepts one.

The query remains callback-capability-gated. For asynchronous work, callers
must resolve and own the path while still inside the original legal Excel
callback, then send only the owned path to the worker.

### 2.4 Large add-in descriptor generation

For large applications, registration metadata should be generated from the
same annotated functions/commands or from an application-owned contract file.
The framework should prefer deterministic explicit generated tables over
linker-section discovery unless a separate design proves linker discovery is
portable and auditable.

PyroApp's existing package-contract generator is a good consumer pattern: it
can eventually emit Rust descriptor tables and Ribbon callback metadata from
one source of truth.

## 3. Current-host `Application` bridge

### 3.1 Why `OwnedApplication` and generic ROT attachment are wrong

An in-process XLL needs the `Application` object for the Excel process that is
currently hosting that XLL.

It must **not** call `OwnedApplication::new`, because that intentionally creates
a fresh private Excel process. It must also not choose a generic
`GetActiveObject("Excel.Application")` result as the stable path, because
multiple Excel processes can exist and ROT selection does not prove host
identity.

### 3.2 New ownership mode

The preferred abstraction is conceptually:

```rust
HostedApplication
```

with these hard rules:

- it represents the exact hosting Excel instance;
- Excel owns process lifetime;
- it never grants `Quit` authority;
- dropping it releases only this add-in's COM reference;
- it remains apartment/thread-bound and is neither `Send` nor `Sync`;
- acquiring it never silently starts another Excel process;
- acquiring it never silently falls back to an unrelated active Excel;
- its process identity is verified during live tests;
- ordinary `excel-com::Application` behavior is reused rather than duplicated.

The bridge should live above both `excel-api` and `excel-com`, preferably in a
Windows-only optional `excel-api-desktop` crate, to avoid a dependency cycle.

### 3.3 Acquisition research gate

The exact current-host acquisition mechanism is deliberately not frozen in
this architecture note. Candidate Windows/Office mechanisms include deriving
the native object model from the current Excel window and other host-provided
bridges used by mature XLL frameworks. Before implementation, the selected path
must be pinned against authoritative documentation, Excel-DNA/other-client
evidence where useful, and real multi-process Excel tests.

Acceptance is behavioral, not merely "a COM call returned an Application":

1. launch two isolated Excel processes;
2. load the test XLL in only one;
3. acquire `HostedApplication` from an Excel-issued callback;
4. prove its window/process identity matches the hosting process;
5. prove the other Excel is untouched;
6. close/reload the XLL repeatedly without leaking a process or reference;
7. prove `HostedApplication` has no shutdown ownership.

### 3.4 Missing Excel object-model helpers

The PyroApp acceptance path justifies implementing and live-testing at least:

- `Application.StatusBar` get/set, preferably with a restoring guard where
  useful;
- `Application.Intersect` with explicit optional-empty semantics;
- `Application.ScreenRefresh` as a best-effort presentation operation.

A raw/general `Application.Run` public wrapper is **not** required merely for
PyroApp. The macro-wake adapter below may use a narrowly internal typed call so
the framework does not encourage arbitrary dynamic macro execution.

## 4. Autonomous macro wake adapter

### 4.1 Preserve the M17 dispatcher boundary

M17 intentionally separates queuing from notification: enqueueing does not wake
Excel, and queued work can run only in a later genuine compatible Excel-issued
callback. That ownership/capability model remains correct and should not be
replaced.

The missing piece is a **notification adapter** that causes Excel to invoke a
registered helper macro when queued macro-capable work exists.

Conceptually:

```text
producer thread / COM callback / timer-independent application code
        |
        | enqueue owned work
        v
M17-compatible bounded queue
        |
        | coalesced wake request
        v
UI-thread hidden/message-only window
        |
        | Application.Run(unique registered helper macro)
        v
Excel invokes XLL helper command
        |
        v
genuine MacroContext
        |
        v
drain compatible work
```

### 4.2 A Windows message does not create Excel capability

The hidden-window callback is useful only as a rendezvous on Excel's UI thread.
Being on that OS thread does **not** manufacture `MacroContext` and does not make
arbitrary `Excel12v` calls legal.

The window procedure may perform only the narrowly audited host operation needed
to ask Excel to invoke the registered helper macro. Excel itself must then
enter the XLL through the registered macro export, where normal thunk/context
rules create the genuine capability.

### 4.3 Busy/retry behavior

Excel can reject `Application.Run` while editing, modal, busy, calculating, or
otherwise unable to enter macro context. The adapter therefore needs bounded,
coalesced retry behavior similar in principle to mature XLL frameworks:

- one pending wake flag rather than one Windows message per task;
- a UI-thread timer only to retry the **wake attempt**, never to call Excel C
  API directly;
- explicit busy/rejected/transient/fatal classification;
- no spin loop and negligible idle CPU;
- shutdown cancels the timer and destroys the window before code can unload;
- stale runtime generations cannot wake a newer runtime;
- repeated messages after shutdown are safe no-ops or are prevented by teardown.

The retry interval and backoff policy are implementation details to validate in
real Excel, not fixed by this document.

### 4.4 Registered helper macro lifecycle

The wake subsystem should register one hidden/uniquely named helper command per
runtime generation, retain its registration ID, and unregister it during close
only after new work is disabled and pending/running work is retired or drained
according to the dispatcher contract.

The helper name must avoid collisions between XLL paths/processes/reloads and
must never be user-controlled formula text.

### 4.5 Application work extension model

M17 currently has a deliberately sealed operation catalogue and does not accept
an arbitrary `Box<dyn FnOnce(&MacroContext)>`. Do not remove that safety choice
accidentally just to support one application.

Before stable implementation, choose and review one extension model. The
preferred direction is an **owned typed payload plus statically registered
handler**. For example, an application may enqueue a small `Send + 'static`
message such as `ContinueDerivative(run_id)`; the registered handler receives
that owned message plus callback-scoped `MacroContext`/host access and looks up
application-owned state on the Excel thread.

An arbitrary macro-closure API remains possible only as a separate explicit
capability design. If adopted, callback borrows must be impossible to retain and
shutdown must still prove that no closure can execute after unlink.

For PyroApp specifically, COM `Range`/`Worksheet` wrappers should not cross
threads with queued work. Store owned IDs/addresses/state and reacquire host
objects inside the real macro callback.

### 4.6 Relationship to `xlcOnTime`

A verified hidden-window plus `Application.Run` adapter can solve the actual
wake requirement without depending on undocumented/insufficiently documented
`xlcOnTime` behavior or an XLM-policy exception. The existing issue #30
research remains useful evidence until the new adapter passes its own live
matrix; it should not be deleted merely because this roadmap prefers another
path.

## 5. Ribbon architecture

### 5.1 Optional crate boundary

Ribbon support should remain an optional Windows/Office integration, preferably
in an `excel-api-ribbon` crate. It must not add Office COM dependencies to the
normal worksheet-only XLL path.

The implementation requires a reviewed Office COM contract for the interfaces
needed by the chosen add-in topology, including Ribbon XML delivery, callback
`IDispatch`, `IRibbonUI` lifetime/invalidation, connection/disconnection, and
installer registration.

### 5.2 Ribbon callbacks are not `MacroContext`

A Ribbon callback arriving on Excel's UI thread does not by itself prove Excel
C API callback capability. Ribbon code may perform audited COM object-model work
through its COM callback context, but operations that require Excel C API macro
capability must enqueue a macro task through the wake adapter and run later in
the registered XLL macro.

This prevents "main thread" from becoming an unsafe substitute for the existing
capability model.

### 5.3 Artifact topology is a research decision

Do not assume the production Ribbon provider must share the XLL binary.
Evaluate at least:

1. a combined XLL + COM-class-server binary with one coordinated unload state;
2. a companion Rust Ribbon COM DLL with independent COM lifetime that invokes
   XLL commands/uses a narrow process-local coordination channel.

The existing COM architecture deliberately keeps RTD COM lifetime separate
from the XLL. Ribbon may have different constraints, but combining hosts must
be justified by unload/reload evidence rather than artifact-count preference.
The first prototype may use the topology that produces the smallest auditable
lifetime surface.

### 5.4 Generated callback metadata

Large applications should be able to generate Ribbon XML/callback metadata from
a single application contract. The framework should validate callback names,
argument/return shapes and duplicate controls at build time where practical.

PyroApp's current package contract can serve as the first consumer: instead of
generating C#, it can emit deterministic Rust metadata plus installer inputs.

## 6. Async execution is a separate compatibility decision

### 6.1 Native async is not Excel-DNA `ExcelAsyncUtil.Run`

The current `excel-api` async implementation uses Excel's native asynchronous
UDF contract (`>` return, `X` handle, later `xlAsyncReturn`). Excel-DNA's
`ExcelAsyncUtil.Run`, as used by current PyroApp, is an RTD/topic-oriented
behavior with its own identity, recalculation and disconnect semantics.

These mechanisms must not be treated as interchangeable without live evidence.

### 6.2 Pure-Rust cutover does not require RTD initially

PyroApp's release default is Blocking. Therefore the first pure-Rust cutover can
use ordinary synchronous XLL functions while retaining the existing isolated
worker processes. "Blocking" here means the worksheet call waits for its IPC
response; ChemApp remains outside Excel.

This sharply reduces the first migration risk and keeps the async decision
separate from the language/runtime cutover.

### 6.3 Native-async parity experiment

After the blocking pure-Rust path is stable, test whether one native async
registration can reproduce the product behavior required by PyroApp's optional
Async mode, including any immediate/blocking preference semantics.

The matrix must cover:

- completion before and after the entry thunk returns;
- dependency ordering and recalculation;
- dynamic-array/spill results;
- cancellation and calculation-ended events;
- input deep-copy and workbook-relative path capture;
- execution/connection generation changes;
- switching a running request between user execution preferences;
- derivative-matrix forced-Blocking scope;
- close/unload and late completion suppression.

If native async cannot preserve the required user-visible contract, design a
separate optional RTD/topic subsystem. Do not change the core async contract to
imitate RTD.

## 7. IntelliSense, dialogs and 32-bit support

### 7.1 IntelliSense

Function/argument descriptions and help topics are part of normal registration
and should be preserved. Excel-DNA's enhanced inline IntelliSense overlay is a
separate UI Automation feature and is not a blocker for a pure-Rust XLL.

A future `excel-api-intellisense` experiment may reuse the same registration
metadata, but should not be coupled to the first host/Ribbon milestones.

### 7.2 Application dialogs

Progress, About and diagnostic dialogs are application UI. The framework may
provide host-window handles or generic parent-window helpers if broadly useful,
but it should not acquire PyroApp-specific dialog abstractions.

### 7.3 32-bit Excel

Current `excel-api` deliberately targets 64-bit Excel while PyroApp still builds
x86 sidecars. Full historical deployment parity therefore requires a separate
32-bit XLL ABI/packaging effort. It is not a prerequisite for proving the
pure-Rust architecture on the supported x64 target.

No code should claim x86 support until ABI layout, calling convention, export,
packaging and real 32-bit Excel tests pass.

## 8. Safety and compatibility rules

The host-integration work must preserve these rules throughout all milestones:

1. `excel-api` callback capability is stronger than thread identity; no helper
   invents a context merely because it runs on Excel's UI thread.
2. Background threads may not call arbitrary Excel C API functions.
3. COM wrappers remain apartment-bound and are not made `Send`/`Sync` to make a
   scheduler convenient.
4. Cross-thread queued state is owned Rust data. Excel COM objects are reacquired
   in the owning callback/apartment unless an explicitly reviewed COM marshaling
   API is used.
5. `HostedApplication` never starts or quits Excel and never attaches to an
   unverified unrelated Excel process.
6. Worker/process/transport policy stays outside the framework.
7. Ribbon COM and XLL lifecycle each have explicit shutdown states; no callback
   may target unloaded code.
8. Panics are contained at every C/COM/XLL callback boundary.
9. Security settings, Trusted Locations, macro/XLM policy and organization-wide
   Office policy are never weakened to make a test pass.
10. Existing core 1.0 public behavior remains stable unless a separately
    reviewed semver change is intentional.
11. The existing `excel-com` live migration suite remains a regression gate for
    COM refactoring used by host integration.
12. A pure-Rust PyroApp migration preserves its public worksheet function names,
    normal not-found/error conventions and array shapes unless PyroApp itself
    documents a deliberate compatibility change.

## 9. Validation strategy

### 9.1 Deterministic unit/native tests

Use mocks/native fixtures for:

- help-topic registration position and lifetime;
- XLL path ownership and reload cleanup;
- typed document-directory call arguments/results;
- wake queue bounds, coalescing and generation isolation;
- timer/retry state machine;
- helper-macro registration/unregistration ordering;
- shutdown with queued/selected/running work;
- panic containment;
- Ribbon callback metadata and COM refcount/QueryInterface behavior once added.

### 9.2 Real Excel host tests

Real Excel is authoritative for:

- exact current-host `Application` identity with two Excel processes;
- host acquisition during open/command/Ribbon paths that are actually supported;
- StatusBar/Intersect/ScreenRefresh wrappers;
- hidden-window wake while idle, calculating, editing, modal and shutting down;
- `Application.Run` busy/retry behavior;
- no callback after unload/reload;
- automatic derivative-style yield/resume sequencing;
- Ribbon load, callbacks, invalidation, uninstall/reload;
- native async parity if/when that milestone is attempted.

Every harness owns and terminates only the Excel process it creates.

### 9.3 PyroApp cross-repository acceptance

PyroApp is the integration proof, not the source of generic abstractions.
Acceptance should compare the Rust XLL against the maintained package contract
and current release tests:

- exact registered public XLL function inventory;
- argument names/help/category/help URLs;
- representative scalar and spill results;
- workbook-relative file resolution;
- local worker startup, pool concurrency and shutdown;
- DAT list/get/set behavior through direct Rust calls;
- manual derivative-matrix transaction;
- automatic derivative-matrix continuation;
- Ribbon settings/status/actions;
- clean Excel shutdown with no worker or COM/Ribbon leak.

The current C# XLL should remain the behavioral oracle until each corresponding
Rust acceptance gate passes.

## 10. Staged roadmap

The milestone prefix `H` means **host integration** and intentionally does not
renumber historical core `M`, optional `E`, or COM-extraction `C` milestones.

### H0 - Architecture baseline

Status: **complete by this document**.

Deliverables:

- classify current PyroApp/Excel-DNA responsibilities;
- define framework/application boundaries;
- define host-Application and macro-wake invariants;
- make async/Ribbon/IntelliSense/x86 independent decisions;
- connect the work to the generic `com-automation` extraction roadmap.

Gate: architecture review only; no runtime behavior changes.

### H1 - XLL metadata and location ergonomics

Implement:

- optional `help_topic` in registration metadata and macro syntax;
- process-lifetime XLL path/directory access;
- typed active-workbook-directory query;
- deterministic large add-in descriptor generation or an approved application
  contract integration pattern.

Tests:

- exact `xlfRegister` argument construction;
- compile-pass/compile-fail macro coverage;
- lifecycle reload ownership;
- real Excel smoke for help/link metadata and workbook-relative path capture.

Gate: no COM/Ribbon dependency is introduced into core crates.

### H2 - Current-host Application bridge

Implement/prove:

- optional Windows `excel-api-desktop` bridge;
- exact host-Excel acquisition with no ROT ambiguity;
- `HostedApplication` ownership mode;
- missing `excel-com` StatusBar, Intersect and ScreenRefresh wrappers;
- process/apartment/thread invariants and diagnostics.

Tests:

- two simultaneous Excel processes;
- only the hosting process is observed/mutated;
- repeated XLL load/unload;
- no Quit authority;
- existing `excel-com` live suites remain green.

Gate: host identity and lifetime must pass real Excel before H3 depends on it.

### H3 - Autonomous macro wake adapter

Implement:

- UI-thread hidden/message-only synchronization window;
- coalesced posted wake;
- unique registered helper macro;
- narrowly audited `Application.Run` bridge;
- busy/edit/modal retry timer;
- runtime-generation and unload safety;
- reviewed application task extension model compatible with M17.

Tests:

- wake from a producer thread;
- batching/coalescing;
- editing and modal retry;
- calculation/busy conditions;
- close while queued/retrying/running;
- reload with stale producer handles;
- idle CPU/latency sanity.

Gate: callback work runs only after Excel enters the helper XLL macro and yields
a genuine `MacroContext`.

### H4 - Derivative-style host acceptance

This is primarily a consumer-proof milestone rather than a new generic API.
Use a compact acceptance fixture, then PyroApp itself, to prove:

- active-sheet range resolution;
- Manual-mode worksheet calculation and `CalculationState` wait;
- Automatic-mode write -> yield -> wake -> resume ordering;
- status-bar restoration and presentation refresh;
- cancellation/restoration and error aggregation;
- no COM wrapper crosses a producer thread.

Gate: the current PyroApp B1:B6 Manual and Automatic acceptance semantics can be
implemented without Excel-DNA `QueueAsMacro`.

### H5 - Ribbon COM integration

Before code, pin the required Office type-library/interface contract and select
combined-vs-companion binary topology.

Implement:

- optional `excel-api-ribbon` crate;
- COM activation/connection lifecycle required by the selected topology;
- Ribbon XML provider;
- callback `IDispatch` surface;
- owned `IRibbonUI` wrapper and invalidation;
- generated callback metadata from application contracts;
- installer registration and reversible cleanup.

Where practical, reuse the generic COM server/IDispatch machinery from
`docs/architecture/com-automation-roadmap.md` rather than growing a second
manual COM-vtable implementation. If the generic server milestone is not ready,
any `windows-core` spike is provisional and must not freeze duplicate ownership
abstractions.

Gate: real Excel load/click/query/invalidate/uninstall/reload tests pass without
loosening Office security policy.

### H6 - Pure-Rust x64 blocking consumer cutover

Cross-repository PyroApp gate:

- build the shipping XLL in Rust;
- port worksheet functions and commands;
- call Rust DAT/parser code directly;
- port worker pool/named-pipe/protobuf client code to Rust application crates;
- use H2/H3 for derivative behavior;
- use H5 for the product Ribbon;
- retain Blocking as the supported UDF execution mode for the first cutover;
- remove the shipping dependency on .NET/Excel-DNA only after full package
  contract and installer acceptance passes.

This milestone is the earliest point at which PyroApp's C# layer can be removed
without requiring RTD-style async parity, enhanced IntelliSense or 32-bit XLL
support.

### H7 - Async behavior parity decision

Run the native-async matrix in section 6. If native async satisfies the product
contract, implement the smallest policy layer above the existing M16 engine.
If it does not, write a separate RTD/topic architecture before adding code.

Gate: no claim of parity based only on successful `xlAsyncReturn`; behavior must
match the selected user-facing execution contract through recalculation,
cancellation, preference changes and unload.

### H8 - Hardening and optional parity extensions

Possible independent workstreams:

- enhanced IntelliSense UI;
- 32-bit Excel ABI and packaging;
- richer Ribbon controls/task panes if needed;
- signing/installer hardening;
- longer reload/soak/busy-state matrices;
- public API/semver stabilization for `excel-api-desktop` and
  `excel-api-ribbon`.

None of these should block H6 unless the product release explicitly requires
that parity dimension.

## 11. Dependency on the COM Automation extraction roadmap

The two roadmaps are related but should not be serialized unnecessarily.

- H1 is independent of generic COM extraction.
- H2 benefits from C1/C2 generic apartment/interface/dispatch ownership, but may
  use the current `excel-com` implementation while that extraction proceeds as
  long as ownership is not duplicated.
- H3 primarily needs the hosted Excel bridge and normal XLL registration; it
  does not require BYREF/events/marshaling.
- H5 should preferentially consume the generic COM server/IDispatch work planned
  around C5 rather than create a permanent second server kernel.
- H7 RTD, if ultimately required, has its own COM callback/marshaling concerns
  and should reuse the generic foundation where the contracts match.

The practical order can therefore be:

```text
H1
 |
 +----> H2 ----> H3 ----> H4
 |                         |
 |                         +----> H6 (with H5)
 |
COM C1/C2 -----------------+
       \
        \--> COM server groundwork ----> H5

H7 after H6 stability
H8 independently as product requirements demand
```

## 12. Completion criteria

The host-integration roadmap is complete when a substantial application can be
built and shipped without a managed Excel integration layer while preserving
the core safety model.

For the first acceptance application, that means:

- the XLL is Rust;
- worksheet and command registration is native;
- current-host Excel object-model access is exact and ownership-safe;
- queued UI/macro work can wake Excel autonomously without inventing callback
  capability;
- Ribbon integration is Rust and unload-safe;
- application workers/transports remain application-owned;
- the user-visible worksheet API and derivative command pass compatibility
  tests;
- Excel closes/reloads without stale callbacks, COM references or worker
  processes;
- any remaining Async/IntelliSense/x86 differences are explicit product scope,
  not hidden framework limitations.

Until those gates pass, the existing managed PyroApp layer remains the
behavioral reference rather than being deleted early.