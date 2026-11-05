# Excel-DNA-equivalent Rust infrastructure: C-API-first roadmap

Status: **proposed architecture roadmap**, not an implementation claim.
Tracking: [excel-api #76](https://github.com/evnekdev/excel-api/issues/76).
First acceptance consumer: [evnekdev/pyroapprs](https://github.com/evnekdev/pyroapprs).
Inspected 2026-11-06:
- `Excel-DNA/ExcelDna/master` at `f484749260d5cc43c35f7aea5f7a0b0f8f74a78f`;
- `evnekdev/excel-api/master` at `dcaa8e0ada6afaf2cc5303c2d5d764caaa411d6c`;
- `evnekdev/pyroapprs/master` at `7dad2f0e57beb393d9f3d5a3edea5779c0e908f6`.

## 1. Outcome and design rule

Make the `excel-api` family a reusable, native Rust framework covering the broad functional capability families of Excel-DNA over time. The **first practical release milestone** is replacement of PyroApp's entire managed Excel-DNA integration *without reducing the current product's required functionality*. Later stages provide broader Excel-DNA analogue coverage for unrelated native Rust add-ins.

This is **not** a literal C# or .NET port. The native Excel **C API is the primary host integration model**. Do not import or depend on the separate `excel-com` general Automation/object-model crate for this programme. For features the Excel C API does not expose, explicitly classify unavoidable, narrowly scoped native Office/Windows protocols rather than disguising them as `Excel12v` functions. No generic Excel object-model clone is a prerequisite.

The previous [H0–H8 native XLL roadmap](native-xll-hosting-roadmap.md) is historical, partially useful research. Its dependencies on current-host Excel object-model Automation and `excel-com` are **not the selected architecture** for this programme. Reconcile those milestones through scoped follow-up issues; do not silently overwrite historical findings.

Maintain the ownership, callback legality, memory/free, generation isolation, panic containment, and native test invariants already implemented in the three core crates. No speculative large-scale reorganization; introduce optional crates only for a demonstrated, separate protocol/lifecycle boundary.

## 2. Capability equivalence inventory

Status vocabulary: **implemented** = source exists, **preview** = functionality exists but material live Excel acceptance is incomplete, **experimental** = research/prototype, **missing** = no production capability identified, **out-of-C-API** = a separate native protocol is required, **not-applicable** = managed technique is unnecessary in Rust. These statuses describe the inspected revision, not the current build or test execution.

| Excel-DNA family | Rust framework target | Current gap/decision | Priority |
| --- | --- | --- | --- |
| Native XLL ABI, `xlAuto*` lifecycle, exports | `excel-api-sys`, `excel-api` | x64 implemented; complete x86 XLL ABI and real host validation missing | P0 |
| Function/command registration, metadata, help, categories | `excel-api`, `excel-api-macros` | basic static registration implemented; help topics, ergonomic large descriptors, broader dynamic registration absent/limited | P0/P2 |
| `XLOPER12`, UTF-16, arrays, errors, missing/nil | core | substantial implemented; audit PyroApp-specific coercion, shape and special-sentinel parity | P0 |
| Reference/sheet/range semantics | core | borrowed references implemented; reference construction, legal read/write, address conversion and multi-area operations need broadening | P0/P1 |
| Excel function and XLM command invocation | `excel-api-sys` identifiers + typed `excel-api` catalogue | currently intentionally narrow safe catalogue; comprehensive verified coverage is long-term work | P0/P2 |
| Workbook, caller and document metadata | typed Excel C API functions | missing relevant `xlfGetDocument` and other information functions for path/context parity | P0 |
| Worksheet updates and recalculation | typed special/XLM command wrappers, legal macro contexts | `xlSet`, `xlcCalculateNow`, range write/clear, selected calculation commands need verified contracts and tests | P0 |
| Thread-safe calculation | core | implemented registration/contexts; host matrix and call legality must be verified | P0/P1 |
| Native async `xlAsyncReturn` | `excel-api` | preview; not equivalent to Excel-DNA RTD-backed `Run`; async/macro-sheet registration conflict needs early resolution | P0 |
| RTD/observables/streaming | optional native Rust protocol crate | experimental prototype only; Excel RTD is a COM protocol, **not** the Excel C API | P0 feasibility, P2 full API |
| `QueueAsMacro` and reliable autonomous macro wake | optional host scheduling adapter + core dispatcher | core queue exists but does not wake Excel; Excel-DNA uses hidden window, retry and `Application.Run`; a C-API-only alternative has **not** been demonstrated | P0 feasibility |
| XLL command macros, menu commands | `excel-api` | basic command registration implemented; larger command/XLM catalogue and lifecycle adapters missing | P0/P2 |
| Ribbon control, callbacks and invalidation | optional narrowly scoped native Office Ribbon protocol | modern Ribbon requires Office interfaces, **not** Excel12v; keep separate from `excel-com` and core | P0 product parity |
| Native error dialogs/progress/logging | optional Rust/Windows UI helpers, application UI | no framework-wide UI required in core; first consumer owns product-specific presentation | P0/P2 |
| Excel function IntelliSense | separate metadata discovery and optional native UI helper | Function Wizard metadata != Excel-DNA IntelliSense overlay; external IntelliSense uses UI Automation | P0 parity decision, P3 implementation |
| Managed registration pipeline and converters | Rust traits, explicit codegen, proc macros | do not copy reflection; provide needed typed conversion hooks, handlers and optional registration transforms | P1/P2 |
| Native object handles and persistent state | optional Rust-owned handle registry | safe generation-aware token/ownership scheme; managed CLR object representation not needed | P3 |
| Add-in manifests/resource embedding | Rust build/metadata and packaging tools | native equivalents needed for selected consumer features; `.dna` compatibility optional, not automatically required | P1/P3 |
| Packaging, signing, testing, debug/CLI tools | tool/xtask-style utilities | basic XLL packaging exists; cross-bitness, resource and integration workflows need expansion | P0/P2 |
| General Office COM, custom task panes, CommandBars | separate optional protocol-specific crates | **out-of-C-API**; do not imply generic Excel Automation; prioritize only with real consumers | P3 |
| Managed CLR, reflection, assemblies, AppDomains, NativeAOT host | no literal Rust implementation | replace *outcomes* with native linking, generated registration, Rust plugin loading only when needed | not-applicable |
| Excel-DNA independent IntelliSense add-in | optional compatible metadata source / native overlay | independently deployed and non-core in upstream; inspect its own repository/version separately | P3 |

Sources: [upstream Excel-DNA](https://github.com/Excel-DNA/ExcelDna), its [XlCall](https://github.com/Excel-DNA/ExcelDna/blob/master/Source/ExcelDna.Integration/XlCall.cs), [registration](https://github.com/Excel-DNA/ExcelDna/tree/master/Source/ExcelDna.Integration/Registration), [async](https://github.com/Excel-DNA/ExcelDna/blob/master/Source/ExcelDna.Integration/ExcelAsyncUtil.cs), [synchronization](https://github.com/Excel-DNA/ExcelDna/blob/master/Source/ExcelDna.Integration/ExcelSynchronizationContext.cs), [RTD](https://github.com/Excel-DNA/ExcelDna/blob/master/Source/ExcelDna.Integration/ExcelRtdServer.cs), and [IntelliSense](https://github.com/Excel-DNA/IntelliSense).

The 1000+ constant declarations in upstream `XlCall.cs` are **not** evidence that 1000 modern Excel operations are universally safe or supported. For each new `Excel12v` safe wrapper, capture verified ID, arguments, result ownership, callback permissions, MTR legality, target Excel versions, and observed host behaviour. Never enable a generic unsafe escape hatch merely to reach a target count.

## 3. Milestone A — PyroApp fully native, with feature parity

Normative authority stays in `pyroapprs`:
- [public XLL API](https://github.com/evnekdev/pyroapprs/blob/master/architecture/xll-public-api.md)
- [legacy workbook compatibility](https://github.com/evnekdev/pyroapprs/blob/master/architecture/legacy-pyroapp-compatibility.md)
- [Blocking/Async and derivative matrix](https://github.com/evnekdev/pyroapprs/blob/master/architecture/udf-execution-and-derivative-matrix.md)
- [required product](https://github.com/evnekdev/pyroapprs/blob/master/release/PYROAPP-REQUIRED-PRODUCT.json)
- [package contract](https://github.com/evnekdev/pyroapprs/blob/master/release/PYROAPP-PACKAGE-CONTRACT.json)
- [acceptance matrix](https://github.com/evnekdev/pyroapprs/blob/master/acceptance/PYROAPP-ACCEPTANCE-MATRIX.json)

Inspected PyroApp C# exports: **56 worksheet functions**, of which **50** declare `IsMacroType = true`; all 56 declare `HelpTopic`. Preserve names, parameter order, types and optional/missing/error handling, scalar/spill-array shape, string punctuation, indexing/sentinels, current Blocking/Async selectable modes, connection/error UI, derivative matrix, Ribbon, diagnostics and installed-product lifecycle. A reduced x64/Blocking-only prototype is **not** a feature-equivalent PyroApp release.

Implementation slices, each requiring a scoped issue in the owning repository before code:

1. **Compatibility fixtures** — generate Rust-oriented assertions from current installed-product behaviour, scientific fixtures and required public API; freeze representative input/output oracle sets. Never change the product baseline to match an incomplete implementation.
2. **Excel C API gaps** — help topics, `xlfGetDocument`/workbook-relative paths, reference/value coercion, safe callback-scoped worksheet read/write/clear, `xlSet`, legitimate recalculation and trigger functions, XLL provenance/path and x64/x86 ABI.
3. **Async feasibility gate, early** — PyroApp's `ExcelAsyncUtil.Run` is RTD/topic-driven and 50 exports are macro-sheet-equivalent. The native `xlAsyncReturn` registration model prohibits that combination. Compare native async with RTD-backed Rust topic mode (or another evidenced design) in a slow-formula, mode-switch, recalculation, cancellation, late-completion and unload test matrix. Keep fallback as a separate optional protocol implementation; do not silently change semantics.
4. **Macro scheduling feasibility gate, early** — the core M17 dispatcher is cooperative and does **not** autonomously wake Excel. Determine whether safe, supported C-API-only scheduling can meet PyroApp's requirements. The legacy Excel-DNA implementation uses Windows message delivery and Office `Application.Run`; do not treat an arbitrary main-thread window message as a valid Excel C API context. Record an explicit architecture decision if a non-C-API mechanism is unavoidable. No `excel-com` dependency.
5. **Native Ribbon feasibility gate** — required current Ribbon is a modern Office protocol, outside the Excel C API. Implement reusable native Rust Ribbon callbacks using narrowly scoped Office interfaces only if approved; isolate their lifecycle and never permit Excel12v calls from ordinary Ribbon callbacks without a genuine legal XLL context. A pure Excel-C-API-only product *cannot* claim Ribbon-equivalence until an alternative accepted UI requirement is negotiated.
6. **Rust PyroApp XLL** — implement 56 UDF adapters and command callbacks, reuse existing Rust DAT parsing, typed protocol/IPC and isolated ChemApp workers. Keep framework generic; application owns scientific semantics, settings and transactions. Do not port unnecessary C# layers into the framework.
7. **Derivative-matrix transaction** — prove B1:B6 preflight, range read/write, trigger exactly once, Automatic/Manual calculation semantics, async continuations, run lock, progress/cancel and restoration. A formal `Application.CalculationState` equivalent has **not** been identified in the Excel C API. Do not invent one or use an undocumented assumption; use observed, tested alternatives or explicitly record impossibility.
8. **Native UI, install and lifecycle** — provide the required dialogs/status/Ribbon, native package contract output and x64/x86 installers, without Excel-DNA or .NET runtime payload. Preserve scoped original PyroApp release/hotfix rules.
9. **Acceptance** — first prove x64 components, then x86 under PyroApp's release order; real installed desktop Excel, correct provenance, live Ribbon clicks, all mandatory scenarios, freeze installer bytes and link hashes. Compilation, mocked unit tests and direct XLL smoke do not imply release acceptance.

**Do not remove the managed reference implementation until the accepted native candidate demonstrates the full product contract.** An experimental Rust prototype may run side-by-side in controlled tests, with no ambiguity about which XLL served a formula.

## 4. Milestone B — Broader Excel-DNA functional equivalents

After the PyroApp acceptance candidate (and in parallel only when isolated and low risk), grow the framework in separate cohesive waves:

- **B1: C API catalogue breadth** — organize verified Excel 12 function/command descriptors by family (information, references, values, workbook/sheet, calculation, macros, registration, internationalization); broader safe typed wrappers and explicit return ownership.
- **B2: Registration and conversions** — typed argument and return conversion traits, dynamic registration where needed, variable arguments, function processors/interceptors, compile-time metadata quality, diagnostic and compatibility tooling.
- **B3: Async, RTD and events** — fully validated native async, optional independently packaged RTD server/topic and observable primitives, cancellation/shutdown, event/recalc observers, safe macro scheduling; independent protocol boundaries.
- **B4: Authoring and UI extensions** — Office Ribbon/custom controls, optional classic menu integration, UI metadata/resources and documentation, enhanced IntelliSense with optional native overlay, diagnostics/help.
- **B5: Data/state ergonomics** — safe persistent Rust object handles, reference construction and disjoint/multi-area operations, typed array marshaling, optional external data providers.
- **B6: Tooling and extension ecosystem** — reproducible x64/x86 template/build/package/signing/debug/testing tools, generator plugins and explicit opt-in native dependency/resource packaging.
- **B7: Optional broader Office protocols** — custom task panes, COM add-ins, legacy CommandBars, general office host integration *only where a separate consumer proves need*. This is not an implicit mandate to clone the Excel object model.

Each feature requires one owner, an issue, source evidence, minimal example, unit tests, targeted real-host tests and an explicit support level. Prioritize compatibility and usable API surfaces over numerical breadth.

## 5. Crate and dependency design

Current foundation (retain):
```text
excel-api-sys        // ABI definitions, verified raw constants, FFI signatures
   -> excel-api      // owned values, callback capabilities, safe calls, lifecycle
   -> excel-api-macros // Rust proc macros for typed UDF/command exports
   -> consumer XLL   // PyroApp first; any Rust XLL later
```
This is conceptual ownership, not a directive to rearrange existing dependency directions.

Potential opt-in modules or crates *only if justified by a concrete consumer*:
- `excel-api-rtd`: minimal native RTD protocol and topic lifecycle (not general Excel Automation).
- `excel-api-ribbon`: minimal Office Ribbon/COM ABI and callbacks; independent lifetime from XLL core.
- `excel-api-host` or similar: host message/wake scheduling bridge **after** a verified legal invocation path exists.
- `excel-api-tools`/examples: build, registration manifest, signing/packaging and diagnostic helpers.
- Optional IntelliSense transport/overlay implementation in a separate crate/repository if separately justified.

No unconditional Windows COM dependency enters ordinary worksheet-function `excel-api`. Do not conflate `Excel C API` with all Office extensibility protocols. No Excel COM Automation object-model wrapper is a target of this roadmap.

## 6. Execution sequence and early decision points

### Wave 0 — Audit and focused feasibility (before broad implementation)
- Compare upstream Excel-DNA's individual capability modules with core Rust implementation and enumerate gaps by **API, behaviour, real-Excel evidence** separately.
- Test the three architecture risks: (A) native async vs macro-type/RTD parity, (B) legal autonomous scheduling and calculation completion without Excel object-model Automation, (C) required modern Ribbon provisioning without C API support.
- Test/plan real 32-bit Excel ABI and host compatibility.
- Do not assume these risks have affirmative solutions; each needs a concrete measured outcome and a stop/rethink condition.

### Wave 1 — Minimum reusable foundation
- Extend verified C API descriptors and x86 ABI, preserve memory/callback invariants.
- Implement needed registration metadata, reference helpers, workbook path and recalc APIs.
- Stabilize unit and in-Excel regression tests of each added capability.

### Wave 2 — Native PyroApp adapters
- Port functions in independent families, using golden comparisons and current managed add-in as reference.
- Port callback-safe IPC worker integration and optional UI modules only after feasibility decisions.
- Keep protected installer, package and release code unchanged until explicitly scoped release work.

### Wave 3 — Full product integration
- Resolve Async/Blocking runtime, macro scheduling, derivative command and Ribbon.
- Demonstrate both bitness targets and installed-product acceptance. Only then retire Excel-DNA from shipping.

### Wave 4 onward — General Excel-DNA coverage
- Implement B1–B7 as demand-based, independently testable modules with explicit support statuses.

## 7. Acceptance and evidence

Maintain three independent axes for each capability: **documented/specification coverage**, **source implementation coverage**, and **verification in real desktop Excel**. Mark PASS / FAIL / NOT RUN / BLOCKED with host Excel version, bitness, sample workbook, SHA/revision and exact scenario. A mock success is not a real-Excel pass.

For PyroApp, authoritative numerical semantics and compatibility stay in its own repo; framework tests validate general XLL behaviour. Respect cross-repository issue-first rules: link one scoped issue in each affected repository *before* any code changes. Avoid cloning dozens of placeholder issues until a cohesive implementation slice is ready.

## 8. Pending explicit decisions

1. Can full required PyroApp Automatic derivative sequencing be implemented without an Excel Automation object model? Microsoft Excel 12 C API coverage and behaviour must be proven with live experiments; no assumed `CalculationState` analogue.
2. Can the required dynamic modern Ribbon be provided by a minimal native Rust Office interface bridge while keeping this programme C-API-first and not importing `excel-com`? If the requirement is **C API exclusively, no other Office protocols**, this is an explicit incompatibility and must be resolved before claiming full parity.
3. Does native async meet the current RTD/topic-based execution preference semantics? If not, RTD remains an explicit protocol exception, separate from the XLL core.
4. Will enhanced IntelliSense require a Rust-only overlay or an approved temporary standalone deployment? Do not silently claim ordinary `xlfRegister` help metadata provides identical IntelliSense.
5. What is the minimum supported 32-bit ABI, OS/Excel version and native installation test matrix for first release?

Answers should be recorded as accepted decisions linked to issue #76 or subsequent bounded investigation issues, with negative evidence preserved.
