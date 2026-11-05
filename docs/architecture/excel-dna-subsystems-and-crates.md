# Excel-DNA subsystems and the Rust crate boundary

Status: **architecture inventory and proposed ownership**, not implementation/release approval.
Parent programme: [#78](https://github.com/evnekdev/excel-api/issues/78).
Scoped architecture issue: [#79](https://github.com/evnekdev/excel-api/issues/79).
Related long-term capability plan: [Excel-DNA roadmap](excel-dna-capability-roadmap.md), [#76](https://github.com/evnekdev/excel-api/issues/76).

Evidence inspected 2026-11-06:
- [Excel-DNA/ExcelDna](https://github.com/Excel-DNA/ExcelDna) `master` @ `f484749260d5cc43c35f7aea5f7a0b0f8f74a78f`, including `Source/ExcelDna.Host.Shared`, `Host`, `Host.NativeAOT`, `Loader`, `Integration`, `ManagedHost`, `SourceGenerator.NativeAOT`, `AddIn.Tasks` and `PackedResources`;
- [evnekdev/excel-api](https://github.com/evnekdev/excel-api) `master` @ `6dcaf03812e5b61cca62bdd4f8c4766506317107` (issue policy commit), `Cargo.toml`, `ARCHITECTURE.md`, `ARCHITECTURE_INDEX.md`, `SUPPORT_MATRIX.md`, crate manifests/modules and RTD/Ribbon design;
- [evnekdev/pyroapprs](https://github.com/evnekdev/pyroapprs) `master` @ `7dad2f0e57beb393d9f3d5a3edea5779c0e908f6`, established product contracts.

Source inspection is **not** proof of complete source-line coverage or real-Excel acceptance. `excel-api` statuses are scoped to the exact inspected sources. Where the source does not establish full behavior, mark partial/unverified.

## 1. Architectural decision

**Retain the three existing publishable core crates:**

1. **`excel-api-sys`** — raw Excel 12 C ABI: `XLOPER12`, `FP12`, Excel errors/type/return codes, SDK function/command numbers, x64/x86 layouts, raw entry points and callback signatures. No ownership or safe call policy.
2. **`excel-api`** — safe Excel C API semantics: owned and callback-borrowed values, memory, ranges/references, coercions, context-restricted calls, registration model, command/lifecycle runtime, async, cooperative dispatcher, diagnostics. Add modules here, not independent crates for each header enum or operation family.
3. **`excel-api-macros`** — proc-macro/code generation for exports, signature checking, registration descriptors, metadata, and future conversion/authoring ergonomics. Runtime state and Excel calls belong in `excel-api`.

The repo also contains an **`excel-com` workspace member**, which is a distinct, general Automation investigation and **must not become a dependency** of this C-API-first programme. The experimental `examples/minimal-rtd-server` is an unpublished RTD research specimen, **not** a fourth published core crate.

**Plan optional, separate crates where the external host protocol and lifecycle truly differ:** `excel-api-ribbon` and, if native-async parity is insufficient or streaming is required, `excel-api-rtd`. A narrow scheduling/wake adapter may eventually deserve `excel-api-host` or similar, **only after a legal live-Excel mechanism is demonstrated**. Do not implement new crates now. Other potential separate crates should wait for concrete consumers.

No amount of Excel C API wrapper coverage can turn Office Ribbon, COM RTD, custom task panes or a UI Automation IntelliSense overlay into Excel12v calls. Their native Rust implementations may use narrowly scoped Office/Windows interop, not a general Excel object-model clone.

## 2. Canonical Excel-DNA subsystem register

Every subsystem has a **stable ID, readable title and plain-English description**. Use all three in GitHub issue titles/bodies and architecture decisions: for example, **`EDNA-02 — Excel C API types, constants and enumerations`**. This tells a reader which low-level definitions allow Rust to communicate with Excel, without requiring the reader to memorize `EDNA-02`.

**ID rules:** `EDNA-NN` IDs are permanent identifiers for *capabilities*, not GitHub issue numbers, crate names, implementation states, or urgency. Assign IDs sequentially; **never renumber or reuse** a published ID. New subsystems receive the next unused ID. If a subsystem later splits, retain the original as an umbrella or mark it superseded with an explicit mapping, and give genuinely new capabilities new IDs. Merge/retire decisions preserve redirect links and historical traceability. Avoid an ID for every source symbol; leaf issues can use names like `EDNA-02 / XLOPER12 error enum` while remaining under their broader subsystem.

**Coverage vocabulary:** *core* = meaningful implementation exists; *partial* = material gaps; *preview* = implementation needs real Excel validation; *experimental* = research prototype; *deferred* = not a production feature; *replacement* = managed/.NET technique whose outcome has a native Rust equivalent but no literal port. These describe inspected source, **not** verified native PyroApp acceptance.

| Stable ID | Understandable subsystem name | What it does | Excel-DNA source evidence | Current Rust mapping and recommendation |
| --- | --- | --- | --- | --- |
| **EDNA-01** | **Native XLL hosting and add-in lifecycle** | Loads a Rust add-in into Excel, registers its entry points, and safely starts, stops and unloads it. | `Host.Shared`, `Loader/XlAddIn.cs` | `excel-api-sys`; `excel-api/runtime.rs`; `minimal-xll`. x64 core; x86 unsupported; extend core, PyroApp P0. |
| **EDNA-02** | **Excel C API types, constants and enumerations** | Defines the exact numbers and binary structures Rust uses to communicate with Excel, such as `XLOPER12`, error codes and function IDs. | `Integration/XlCall.cs`; `Loader/XlTypes.cs` | `excel-api-sys/types.rs`, `constants.rs`, `functions.rs`. Raw ABI partly covered; extend `-sys`, PyroApp P0. |
| **EDNA-03** | **Excel values and type conversions** | Turns Excel numbers, text, booleans, errors and arrays into usable Rust values and back. | `Loader/XlMarshalContext.cs`, `XlDirectMarshal*.cs` | `excel-api/borrowed.rs`, `value.rs`, `convert.rs`. Substantial core; add missing conversions in core, P0. |
| **EDNA-04** | **Memory ownership and returning results** | Ensures Excel and the XLL release the right memory at the right time, without leaks, dangling pointers or double frees. | `Loader/XlMarshalContext.cs`; `Host.Shared/addin.cpp` | `excel-api/excel_owned.rs`, `return_plan.rs`, `return_alloc.rs`. Core; extend audited ownership rules only, P0. |
| **EDNA-05** | **Cells, ranges and Excel references** | Represents individual cells or rectangular/multi-area ranges, and supports legally reading and changing worksheet values. | `Integration/ExcelReference.cs` | `excel-api/borrowed.rs`, `value.rs`, `excel_call.rs`. Partial; typed `xlCoerce`/`xlSet`/addresses in core, P0. |
| **EDNA-06** | **Calling Excel functions and commands** | Allows an add-in to invoke supported Excel worksheet functions, XLM commands and C API services from a legal callback. | `Integration/XlCall.cs`; `Loader/XlCallImpl.cs` | `excel-api/excel_call.rs`; `-sys/constants.rs`. Narrow safe catalogue; expand typed, context-safe descriptors, P0. |
| **EDNA-07** | **Worksheet function registration and help metadata** | Exposes Rust functions as Excel formulas, with names, argument signatures, categories, descriptions and help links. | `Loader/XlRegistration.cs`; `Integration/ExcelAttributes.cs` | `excel-api/registration.rs`, `metadata.rs`; `excel-api-macros`. Core basics; add HelpTopic and parity metadata, P0. |
| **EDNA-08** | **Excel commands, macros and shortcuts** | Exposes callable commands rather than cell formulas, and optionally connects them to shortcuts or classic menus. | `Loader/XlRegistration.cs`; `Integration/MenuManager.cs` | `excel-api/registration.rs`; `excel-api-macros`. Basic commands; extend core registration, P0/P2. |
| **EDNA-09** | **Advanced registration and conversion pipelines** | Customizes how UDF arguments, return values and execution wrappers are generated, including optional or variable arguments. | `Integration/Registration/*`; `ExtendedRegistration/*` | `excel-api/convert.rs`, `registration.rs`; macros. Partial; Rust traits/proc macros, no new crate yet, P2. |
| **EDNA-10** | **Thread-safe and parallel worksheet calculation** | Controls which functions Excel may calculate on multiple threads or in clusters and which C API calls they may safely make. | `Integration/ExcelAttributes.cs`; `Loader/XlRegistration.cs` | `excel-api/context.rs`, `registration.rs`, `excel_call.rs`. Core with host verification gaps, P0/P2. |
| **EDNA-11** | **Native asynchronous worksheet functions** | Runs long computations without keeping a calculation thread blocked and safely sends the result back to Excel. | `Integration/ExcelAsyncHandle.cs`; `NativeAsyncTaskUtil.cs` | `excel-api/async_udf.rs`. Preview; not equivalent to RTD-backed `Run`, PyroApp P0. |
| **EDNA-12** | **Real-time data, subscriptions and observables** | Streams changing values into formulas using topic subscriptions, notifications and automatic recalculation. | `Integration/ExcelRtdServer.cs`, `ExcelRtdObserver.cs`, `ExcelAsyncUtil.cs` | Experimental `minimal-rtd-server` example. Conditional `excel-api-rtd` crate after live Excel proof, P0 feasibility. |
| **EDNA-13** | **QueueAsMacro and main-thread work scheduling** | Transfers deferred work into an Excel-approved macro callback so it can safely interact with the workbook. | `Integration/ExcelSynchronizationContext.cs` | `excel-api/dispatcher.rs`. Cooperative preview; autonomous wake unresolved (#30), P0 feasibility. |
| **EDNA-14** | **Calculation events and cancellation notifications** | Responds to Excel calculation-ended or cancelled signals and coordinates pending operations and shutdown. | `Integration/ExcelIntegration.cs`, `ExcelAsyncUtil.cs` | `excel-api/async_udf.rs`, `context.rs`, `runtime.rs`. Partial/preview; keep callback lifecycle in core, P0. |
| **EDNA-15** | **Modern Office Ribbon controls** | Adds a custom tab, buttons, menus and callbacks to Excel's modern Ribbon interface. | `Integration/CustomUI/ExcelRibbon.cs`, `RibbonUI.cs` | Architecture only; no production crate. Strong optional `excel-api-ribbon` candidate; Office protocol, P0. |
| **EDNA-16** | **Custom task panes and classic Excel UI** | Creates dockable panels or older-style menu/toolbar controls that go beyond ordinary worksheet functions. | `Integration/CustomUI/ExcelCustomTaskPane.cs`, `ExcelCommandBars.cs` | No production integration. Optional, consumer-driven separate Office UI module/crate, P3. |
| **EDNA-17** | **Office add-in COM protocol support** | Handles the Office-specific interfaces, activation, registration and object lifetime needed by Ribbon or task-pane add-ins. | `Integration/ComInterop/*`; `CustomUI/ExcelRibbon.cs` | Separate research; no accepted native protocol API. Narrow optional protocol foundations; never general `excel-com`, P0 for Ribbon. |
| **EDNA-18** | **Formula IntelliSense and interactive help** | Shows improved formula suggestions, argument hints and help while users type formulas in Excel. | `Loader/XlRegistration.cs`; separate `Excel-DNA/IntelliSense` | Core registration metadata; no native overlay. Optional enhanced UI component, parity decision P0, later P3. |
| **EDNA-19** | **Persistent object handles for worksheet formulas** | Lets formulas refer to long-lived application objects by stable identifiers instead of copying entire objects into cells. | `Integration/ObjectHandles/*` | No general registry. Start with an optional core module if demanded; native Rust-owned objects, P3. |
| **EDNA-20** | **Add-in manifests and external library loading** | Describes which components and configuration an add-in loads and how those components are discovered. | `Integration/DnaLibrary.cs`, `ExternalLibrary.cs`; `ManagedHost/*` | Cargo/native XLL build; no `.dna` compatibility. Prefer Rust-native metadata; managed host semantics not copied, P2/P3. |
| **EDNA-21** | **Embedded resources and XLL asset packing** | Bundles images, configuration and supporting files into an add-in and retrieves them at runtime. | `PackedResources/*`; `Host.Shared/ExcelDna.Host.rc` | XLL packaging scripts and Rust resources. Build-time tools, not core runtime; add only for consumers, P1/P2. |
| **EDNA-22** | **Developer build, signing and debugging tools** | Provides repeatable add-in builds, x86/x64 packaging, signing, debugging, templates and deployment checks. | `AddIn.Tasks/*`; `ExcelDnaPack`; `SourceGenerator.NativeAOT/*` | Scripts, examples, packaging architecture. Keep scripts; extract `excel-api-tools` only for shared need, P0/P2. |
| **EDNA-23** | **Diagnostics, errors and troubleshooting** | Records startup, registration and runtime problems and provides actionable error reporting without crashing Excel. | `Integration/Logging.cs`, `LogDisplay.cs`; `Loader/LoaderLogging.cs` | `excel-api/diagnostics.rs`, `error.rs`, `runtime.rs`. Core diagnostics; optional Windows UI only, P1/P2. |
| **EDNA-24** | **Managed CLR and .NET assembly hosting** | Loads .NET code and assemblies and isolates managed plug-ins; this is needed by Excel-DNA but not a native Rust XLL. | `Host/*`; `ManagedHost/*`; `Integration/AssemblyLoader.cs` | Not applicable as literal Rust functionality. Replace outcomes with native build/link choices; no CLR port, P3 optional. |
| **EDNA-25** | **NativeAOT integration and generated interop** | Lets .NET add-ins be compiled ahead of time and generates native-facing glue; Rust already compiles native binaries. | `Host.NativeAOT/*`; `SourceGenerator.NativeAOT/*` | `excel-api-macros`; native Cargo builds. Reuse Rust proc-macro/export techniques; no literal AOT host, P2/P3. |

These **25 distinct, broad capability subsystems** refine the previous 18 provisional table rows: the older row numbers were **not stable IDs**. The inventory explicitly separates callbacks, memory ownership, Office COM protocols, resources and managed-only delivery mechanisms without implying one Rust crate per ID.

### Cross-cutting implementation rule

An `EDNA-xx` subsystem is an **issue-hierarchy branch**, not automatically a Cargo package. For example, `EDNA-02` can have child issues for verified C API constants, enumerations and structures, all implemented in `excel-api-sys`. `EDNA-15` and `EDNA-12` represent separate Office protocols and may justify optional crates only after their feasibility criteria pass. **No new crates are created by numbering.**

Open decisions belong in the investigation parent [#80](https://github.com/evnekdev/excel-api/issues/80) and should cite their relevant stable IDs by readable name. The full long-term programme is [#78](https://github.com/evnekdev/excel-api/issues/78); this numbering decision is [#81](https://github.com/evnekdev/excel-api/issues/81). Parent issues remain open until every required child and the parent's own acceptance criteria are satisfied.

## 3. Crate decisions

| Proposed crate or module | Decision **now** | Extraction threshold | Risk if prematurely extracted |
| --- | --- | --- | --- |
| `excel-api-sys`, `excel-api`, `excel-api-macros` | **Retain** as three publishing foundations | Already distinct by raw ABI, safe runtime and proc-macro compiler boundary | Fragmented safety/ownership model if split arbitrarily |
| `excel-api/src/reference`, `calls/*`, `registration/*`, `values/*`, `handles` | **Modules within `excel-api` as needed**, not new crates | Cohesive code organization; split file/module only when actual source complexity/ownership requires it | Cyclic cross-crate dependencies, duplicated callback capability types |
| `excel-api-ribbon` | **Plan separate optional crate; not implemented** | A verified Office Ribbon host ABI, topology, callbacks, metadata, install/unload model and PyroApp consumer | Implicit COM dependency or unverifiable startup model in core |
| `excel-api-rtd` | **Plan conditional separate optional crate; no production API yet** | Clean-Excel RTD activation, topics/unsubscribe and async parity evidence; reusable server contract | Prematurely claiming Excel formula or async support from COM unit tests |
| `excel-api-host` / wake adapter | **Research-only naming**; no crate yet | A verified, legal mechanism for autonomously invoking a genuine XLL macro callback; independent OS resources and host lifecycle | Treating a background/UI thread as legal for Excel12v; circular event loops |
| `excel-api-intellisense` | **Deferred and consumer-driven** | Independent native add-in/overlay required and metadata interoperability validated | Recreating a complex Windows overlay when ordinary metadata is adequate |
| `excel-api-taskpane` | **Deferred and consumer-driven** | An actual pane/UI application requires Office task-pane interfaces and separate lifetime | A general Excel COM automation facade hidden in the core |
| `excel-api-tools` / `xtask` | **Optional build-time package, not runtime dependency** | Multiple consumers share reproducible packaging, signing, verification and templates | New maintenance cost for scripts that already work |
| `excel-api-legacy` | **Deferred and only if requested** | Real consumer requiring pre-Excel-12 `XLOPER`/Excel4 ABI | Weakening modern Excel12 assumptions without demand |
| `excel-com` | **Not part of this programme** | None under selected constraints | Violates explicit C-API-first/no-general-Automation requirement |

**New crates to create immediately: zero.** The new *boundaries to design* are Ribbon (strong), RTD (conditional, likely relevant to full PyroApp async semantics), and macro-wake (conditional research). Implement verified C API gaps in the three existing crates.

### Dependency sketch

```text
                                  [application / consumer XLL]
                                              |
                        +---------------------+-------------------+
                        |                                         |
                excel-api-macros                         optional native
                 (compile-time)                        Office integrations
                        |                                      |
                        v                                      +--> excel-api-ribbon
                   excel-api                                  +--> excel-api-rtd*
              safe callbacks + runtime                       +--> excel-api-host*
                        |
                        v
                  excel-api-sys
                   raw Excel ABI

* Optional only after a validated independent protocol and consumer.
  Office/RTD callbacks do not directly gain Excel C API callback privileges.
  'excel-com' remains an unrelated workspace member; no dependency edge.
```

This diagram describes conceptual use: proc macros emit references to core runtime APIs, not necessarily normal Cargo dependencies into `excel-api`.

## 4. PyroApp is the first end-to-end acceptance boundary

Current product constraints [public XLL API](https://github.com/evnekdev/pyroapprs/blob/master/architecture/xll-public-api.md), [legacy compatibility](https://github.com/evnekdev/pyroapprs/blob/master/architecture/legacy-pyroapp-compatibility.md), [UDF/derivative contract](https://github.com/evnekdev/pyroapprs/blob/master/architecture/udf-execution-and-derivative-matrix.md), and [required product baseline](https://github.com/evnekdev/pyroapprs/blob/master/release/PYROAPP-REQUIRED-PRODUCT.json) stay authoritative.

| Need | Required framework owners | Priority/unknown |
| --- | --- | --- |
| 56 currently registered public UDFs; 50 marked macro type; documented help topics | `excel-api-sys` + `excel-api` + `excel-api-macros` | **First wave**: metadata/registration, stable shape/types, workbook and caller data, error conversions |
| Workbook-relative data-file paths, typed address/range reads/writes | `excel-api` typed C API calls/reference helpers | **First wave**: `xlfGetDocument`, `xlCoerce`, `xlSet` and other verified operations, callback legalities |
| Blocking and Async selectable and mode-switch behavior | `excel-api` native async + possible `excel-api-rtd` | **Early feasibility gate**: 50 macro-type exports cannot be naively re-registered as native async UDFs and assumed equivalent |
| Modern ribbon controls, callbacks, status/mode/UI | Future `excel-api-ribbon` | **Required** for PyroApp parity; Office protocol is distinct from C API |
| Derivative-matrix B1:B6, exact-trigger recalculation, progress, restore | Typed core C API + legal scheduler + PyroApp transaction logic | **Early feasibility gate**: no established Excel C API equivalent of `Application.CalculationState` and no demonstrated autonomous queue wake |
| x64 and x86 installed-product builds | `-sys` ABI + core + generated exports + packaging tools | **Required**; x64 supported in current core, x86 unsupported. Respect PyroApp x64-first build and independent real-Excel acceptance requirements |
| Formula IntelliSense, diagnostic dialogs, error/UI parity | Core registration metadata, later optional UI tools | **Explicit parity decision** before release; ordinary Excel help metadata != Excel-DNA IntelliSense overlay |

The 56 UDF count is an inspected product surface, not a claim of 56 successful native conversions. **No shipping cutover** occurs until all mandatory installed-product scenarios pass; limited native samples or x64 blocking proofs are development checkpoints only.

## 5. Hierarchical issue plan

Adopt one **single programme parent**, not a flat list of unconnected tickets:

```text
#78  Programme — Rust Excel-DNA functional parity [OPEN]
  #79  Architecture — subsystem and crate allocation [this document]
  [future] Native XLL ABI + x86/x64 hosting
     [future] Excel 12 header enum / constant / union (leaf, if independently testable)
     [future] x86 ABI/exports and real Excel verification
  [future] Safe Excel C API surface
     [future] typed xlfGetDocument
     [future] typed xlSet, memory ownership and legal macro context
  [future] UDF authoring, registration and conversion
     [future] HelpTopic metadata with compile-time conformance
  [future] Async / RTD / dispatcher and calculation lifecycle
     [future] native-vs-RTD async parity experiment
     [future] autonomous legal macro wake feasibility
  [future] Office UI and Ribbon
     [future] Ribbon host ABI + callback lifecycle
  [future] Packaging, diagnostics, test harness and tools
  [future] PyroApp full installed-product cutover
     [future] linked implementation/acceptance issues in pyroapprs
```

Each **leaf** must have an SDK or upstream evidence pointer, exact interface and owning path, effect on ABI/callback permission/memory, explicit test case, acceptance condition and status. An individual enum or `XLOPER12` variant deserves a leaf issue **when meaningful work and verification are independent**; don't create one per numeric constant by default.

A parent issue remains open until **every required child recursively resolves and parent-level acceptance passes**. Depend-on links and incidental issue mentions are not sub-issue containment. Use GitHub native sub-issues where available; current connector lacks the native write operation, so [#78](https://github.com/evnekdev/excel-api/issues/78) and [#79](https://github.com/evnekdev/excel-api/issues/79) presently use explicit reciprocal textual fallback. When cross-repository changes are actually scoped, create/identify a corresponding issue in every edited repository *before* changes.

Do not create a large set of empty implementation issues before respective source-level requirements and evidence are identified. Decompose one capability at a time and track it under its parent.

## 6. Guardrails and deferred decisions

1. **No unconditional Office/COM deps in core**; using a narrow Office `IRibbonExtensibility` or `IRtdServer` implementation is not the same as importing `excel-com` general Automation.
2. **No fabricated callback capability.** All Excel12/Excel12v calls obey the actual registered callback and legal timing. An RTD, Ribbon or window callback never automatically carries `MacroContext`.
3. **No API breadth by constant count.** An SDK constant, type tag or enum can be correct while its wrapper is illegal or the host version cannot execute it.
4. **No automatic crate per subsystem.** Split by ABI/compiler/plugin or separately deployed protocol, not upstream C# namespace/file size.
5. **No premature general Excel object model.** Derivative matrix and Office interaction problems need focused experiments; unsupported capabilities stay flagged.
6. **No changes to `excel-api` core 1.0 release scope or PyroApp hotfix/release contracts** as a side effect of this architecture map.
7. **Verification status:** inspected repository sources and committed architecture only; Rust build, runtime, RTD activation, Ribbon callbacks, x86 Excel and product acceptance **NOT RUN** in this task.

## 7. Primary source anchors

- [Excel-DNA native host](https://github.com/Excel-DNA/ExcelDna/tree/master/Source/ExcelDna.Host.Shared)
- [Excel-DNA integration](https://github.com/Excel-DNA/ExcelDna/tree/master/Source/ExcelDna.Integration)
- [Excel-DNA loader/marshalling](https://github.com/Excel-DNA/ExcelDna/tree/master/Source/ExcelDna.Loader)
- [Excel-DNA function registration](https://github.com/Excel-DNA/ExcelDna/tree/master/Source/ExcelDna.Integration/Registration)
- [Excel-DNA extended registration](https://github.com/Excel-DNA/ExcelDna/tree/master/Source/ExcelDna.Integration/ExtendedRegistration)
- [Excel-DNA Office UI](https://github.com/Excel-DNA/ExcelDna/tree/master/Source/ExcelDna.Integration/CustomUI)
- [Excel-DNA build tasks](https://github.com/Excel-DNA/ExcelDna/tree/master/Source/ExcelDna.AddIn.Tasks)
- [Rust core source](https://github.com/evnekdev/excel-api/tree/master/crates), [workspace manifest](https://github.com/evnekdev/excel-api/blob/master/Cargo.toml), [core support matrix](https://github.com/evnekdev/excel-api/blob/master/SUPPORT_MATRIX.md)
- [RTD research](https://github.com/evnekdev/excel-api/blob/master/RTD_STREAMING_ARCHITECTURE.md), [Ribbon research](https://github.com/evnekdev/excel-api/blob/master/RIBBON_UI_ARCHITECTURE.md)
