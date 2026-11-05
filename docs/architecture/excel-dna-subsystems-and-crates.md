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

## 2. Upstream-to-Rust subsystem map

Status: **core** = meaningful implementation in current core; **partial** = material functionality/gaps remain; **preview** = implemented but required host validation pending; **experimental** = nonproduction prototype; **absent/deferred** = not an available production API; **replacement** = .NET-specific technique not needed literally.

| # | Excel-DNA broad subsystem, representative source | Current Rust owner / evidence | Coverage and proposed location |
| --- | --- | --- | --- |
| 1 | **Native XLL host, loader, exports and lifecycle** — [`Host.Shared`](https://github.com/Excel-DNA/ExcelDna/tree/master/Source/ExcelDna.Host.Shared), `Loader/XlAddIn.cs` | `excel-api-sys` raw callbacks; `excel-api/runtime.rs`; `examples/minimal-xll` | **Core x64, x86 absent**. Keep in core, add x86 target-specific ABI/exports/tests; do **not** copy the CLR host. |
| 2 | **Excel 12 ABI structures, constants and function codes** — `Integration/XlCall.cs`, `Loader/XlTypes.cs` | `excel-api-sys/src/types.rs`, `constants.rs`, `functions.rs` | **Partial catalogue**. Expand in `excel-api-sys`; each verified enum/structure can be a fine-grained issue, not a new crate. Raw numeric IDs do not assert safe runtime support. |
| 3 | **Value/string/error/array marshalling and conversions** — `Loader/XlMarshalContext.cs`, `XlDirectMarshal*.cs` | `excel-api/borrowed.rs`, `value.rs`, `convert.rs`, `error.rs`, `return_plan.rs` | **Substantial core**. Expand inside `excel-api`; preserve Excel-owned vs XLL-owned memory, callback lifetimes, scalar/multi shapes and bounds. |
| 4 | **References and worksheet read/write** — `Integration/ExcelReference.cs`; `XlCall.cs` | `excel-api/borrowed.rs`, `value.rs`, `excel_call.rs` | **Partial**. Add typed reference manipulation, `xlCoerce`/`xlSet`, address/text reference functions within safe core; legality is per callback, not just per thread. |
| 5 | **C API dispatch: worksheet functions, XLM commands and special calls** — `Integration/XlCall.cs`, `Loader/XlCallImpl.cs` | `excel-api/excel_call.rs` typed descriptor/capability backend | **Narrow core catalogue**. Grow by legal context and purpose in `excel-api`; raw IDs in `-sys`. Do not promote all upstream `XlCall` constants to validated safe operations. |
| 6 | **UDF/command registration, metadata and shortcuts** — `Loader/XlRegistration.cs`; `Integration/ExcelAttributes.cs` | `excel-api/registration.rs`, `metadata.rs`; `excel-api-macros` | **Core basic, partial breadth**. HelpTopic, optional/default params, richer type text, commands and registration info belong in existing crates. |
| 7 | **Extended authoring: function processors, registration transforms, parameter/return converters, params, lambdas** — `Integration/Registration/*`, `ExtendedRegistration/*` | `excel-api-macros` and `excel-api/convert.rs`, `registration.rs` | **Partial**. Prefer compile-time Rust traits/proc macros over C# reflection/expression-tree translation; no new crate until a reusable third-party registration pipeline justifies it. |
| 8 | **Thread-safe UDF, MTR/cluster flags, call legality** — `Integration/ExcelAttributes.cs`, `Loader/XlRegistration.cs` | `excel-api/context.rs`, `registration.rs`, `excel_call.rs` | **Core with host verification gaps**. Keep context and threading invariants within `excel-api`; cluster support is not implied by a flag alone. |
| 9 | **Native async and cancellation** — `Integration/ExcelAsyncHandle.cs`, `NativeAsyncTaskUtil.cs` | `excel-api/async_udf.rs` | **Preview**. Keep safe native `xlAsyncReturn` lifecycle in core; cannot assume parity with RTD-backed Excel-DNA `ExcelAsyncUtil.Run`. |
| 10 | **RTD, observable topics, streaming and async-as-recalc** — `Integration/ExcelRtd*.cs` and `ExcelAsyncUtil.cs` | `examples/minimal-rtd-server`; `RTD_STREAMING_ARCHITECTURE.md` | **Experimental only**. Independent COM `IRtdServer` lifecycle and packaging. Strong **conditional new crate `excel-api-rtd`** after clean-host activation and PyroApp async feasibility evidence. Not core Excel C API. |
| 11 | **Main-thread coordination, QueueAsMacro, timer/message retry and cancellation events** — `Integration/ExcelSynchronizationContext.cs` | `excel-api/dispatcher.rs`; calculation event handlers in `async_udf.rs` | **Cooperative dispatcher preview; autonomous wake absent**. Keep queue/capability in core; investigate a **narrow optional host/wake adapter** after evidence. Excel-DNA uses Windows messages + `Application.Run`; a window callback alone has no Excel C API capability. |
| 12 | **Modern Ribbon extensibility and commands** — `Integration/CustomUI/ExcelRibbon.cs`, `RibbonUI.cs`, `ComInterop/ExcelComAddIn.cs` | `RIBBON_UI_ARCHITECTURE.md`; no production crate | **Deferred, first strong new-crate candidate: `excel-api-ribbon`**. Office Ribbon protocol, callbacks, XML, invalidation, COM lifetime/registration. A Ribbon callback is not an Excel C API macro context. |
| 13 | **Custom task panes, old CommandBars and broader COM add-in integration** — `Integration/CustomUI/ExcelCustomTaskPane.cs`, `ExcelCommandBars.cs`; `ComInterop/*` | Existing optional research, separate `excel-com` not used by this programme | **Deferred**. Future focused UI/Office crate(s) only after a consumer requires them. General Excel object-model Automation is **outside** this C-API-first framework; don't add as a mandatory foundation. |
| 14 | **Formula help, registration metadata and IntelliSense** — `Loader/XlRegistration.cs`; independent [Excel-DNA/IntelliSense](https://github.com/Excel-DNA/IntelliSense) | `excel-api/registration.rs`, `metadata.rs`; no overlay | **Basic metadata partial; enhanced overlay absent**. Extend core function metadata first; optional separate IntelliSense add-in/overlay later. UI Automation overlay is not part of Excel12v. |
| 15 | **Managed object handles and callable object lifecycle** — `Integration/ObjectHandles/*` | No general object registry in current core | **Deferred**. Start with an optional `excel-api/handles` module only if a native consumer needs tokenized Rust objects; generation-safe IDs and unload cleanup. Do not copy CLR object semantics literally. |
| 16 | **Add-in discovery/configuration, resource pack, native vs managed host loading** — `Integration/DnaLibrary.cs`, `ExternalLibrary.cs`, `Host`, `ManagedHost`, `PackedResources` | Native `Cargo`/XLL build and `PACKAGING_ARCHITECTURE.md` | **Native replacement**, not literal `.dna`/AppDomain host. Resource embedding and optional metadata declarations in tools/build layer; add `.dna` compatibility only for a demonstrated migration demand. |
| 17 | **Build, pack, signing, debug support and NativeAOT toolchain** — `AddIn.Tasks/*`, `ExcelDnaPack`, `SourceGenerator.NativeAOT` | Existing build/package scripts, `PACKAGING_ARCHITECTURE.md`, examples | **Partial tool equivalence**. Build on Rust native compilation and reproducible scripts; future `excel-api-tools`/**xtask** only if reusable workflows outgrow scripts. Rust does not need .NET NativeAOT hosting. |
| 18 | **Diagnostics, exceptions, UI logging and startup/shutdown** — `Integration/Logging.cs`, `ExcelIntegration.cs`, `LogDisplay.cs`, `Loader/LoaderLogging.cs` | `excel-api/diagnostics.rs`, `error.rs`, `runtime.rs`, thunk panic containment | **Core diagnostics partial**. Expand structured tracing and error policies in core; standalone logging windows are optional UI integrations, not mandatory core dependencies. |

`Excel-DNA` is assembled from multiple native, managed, Office/Windows and build-time components. The row map is **functional equivalence**, not package-for-package replication. Excel-DNA itself does not imply one crate per subsystem.

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
