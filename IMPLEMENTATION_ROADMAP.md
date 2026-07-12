# Rust-Native Excel API Implementation Roadmap

## Mission

Build a mature Rust ecosystem for native Microsoft Excel add-ins that:

- calls the Excel C API directly;
- produces native `.xll` files without requiring .NET;
- exposes safe, idiomatic Rust abstractions;
- supports incremental adoption, from raw FFI to high-level procedural macros;
- remains explicit about Excel threading, lifetime, and memory rules.

This roadmap is informed by the Excel-DNA feature set, but it deliberately prioritizes the capabilities that are natural and valuable in a Rust-native implementation.

See [EXCELDNA_CAPABILITY_MAP.md](EXCELDNA_CAPABILITY_MAP.md) for the full feature-by-feature assessment.

## Proposed workspace

```text
excel-api/
├── crates/
│   ├── excel-api-sys/       # Raw Excel C API ABI, constants and structures
│   ├── excel-api/           # Safe values, calls, registration and lifecycle API
│   ├── excel-api-macros/    # #[excel_function], #[excel_command], add-in metadata
│   ├── excel-api-runtime/   # Registration registry, dispatch, allocations, logging
│   ├── excel-api-com/       # Optional COM, RTD and Ribbon support
│   └── excel-api-pack/      # Optional packaging CLI/tooling
├── examples/
│   ├── basic-xll/
│   ├── arrays/
│   ├── commands/
│   ├── async/
│   └── ribbon-rtd/
├── EXCELDNA_CAPABILITY_MAP.md
├── IMPLEMENTATION_ROADMAP.md
└── README.md
```

The crate boundaries should remain provisional until the first working XLL clarifies which abstractions genuinely need separation.

## Design principles

1. **Raw and safe layers remain separate.** `excel-api-sys` mirrors the ABI; higher crates enforce ownership and context rules.
2. **No unwinding across FFI.** Every exported entry point catches panics and converts failures into an Excel error or lifecycle failure.
3. **Excel contexts are typed.** Calculation-safe, macro/main-thread, thread-safe, and asynchronous contexts should expose different operations.
4. **Memory ownership is visible.** Returned strings and arrays must have one documented allocation/free strategy centered on `xlAutoFree12`.
5. **Macros generate glue, not policy.** Conversion traits and registration descriptors remain usable without procedural macros.
6. **Windows-only features are optional.** Core Excel C API support should not pull in COM, WebView2, Tokio, or GUI frameworks.
7. **One working vertical slice precedes broad coverage.** A reliable scalar UDF is more valuable than many unfinished interfaces.

## Milestone overview

| Milestone | Outcome | Main capability groups | Suggested release |
|---|---|---|---|
| M0 | Verified ABI research and project skeleton | headers, build targets, test strategy | pre-release/internal |
| M1 | Excel loads a Rust XLL and calls scalar UDFs | lifecycle, registration, primitive conversion | `0.1.0` |
| M2 | Safe strings, errors, missing values and arrays | ownership, `XLOPER12`, dynamic arrays | `0.2.0` |
| M3 | Ergonomic proc-macro registration | metadata, arguments, generated thunks | `0.3.0` |
| M4 | Commands, Excel calls and execution contexts | macro API, caller/reference wrappers, MTR safety | `0.4.0` |
| M5 | Robust production runtime | logging, diagnostics, tests, packaging | `0.5.0` |
| M6 | Native async UDFs and main-thread dispatch | async handles, executor integration | `0.6.0` |
| M7 | Optional COM, RTD and Ribbon layer | streaming, Ribbon, automation | separate/experimental crates |
| M8 | Tooling ecosystem | packing, docs generation, templates, IntelliSense research | `1.0` candidate |

## M0 — ABI research and repository foundation

### Deliverables

- Cargo workspace and crate skeletons.
- Windows x64 build target first; x86 policy documented but deferred unless required.
- Verified Rust definitions for:
  - `XLOPER12` and nested unions/structures;
  - `FP12` or equivalent array structures;
  - Excel error codes and `xltype*` flags;
  - XLL lifecycle entry-point signatures;
  - `Excel12v` function pointer signature.
- A small ABI verification program comparing C and Rust `sizeof`, `alignof`, and field offsets.
- Licensing and provenance notes for constants derived from Microsoft SDK/header material.
- CI that at least runs formatting, linting, unit tests, and Windows compilation.

### Key decisions

- Whether raw bindings are handwritten, generated with `bindgen`, or generated once and curated.
- How to obtain `Excel12v`: import library, delayed linking, or runtime `GetProcAddress`.
- Whether the first supported toolchain is MSVC only.
- Exact `panic` strategy at exported boundaries.

### Exit criteria

A minimal `.xll` can be built, Excel recognizes its lifecycle exports, and ABI tests pass on x64 Windows.

## M1 — First complete XLL vertical slice

### Deliverables

- `xlAutoOpen`, `xlAutoClose`, `xlAutoFree12`, and `xlAddInManagerInfo12`.
- Safe wrapper for `Excel12v` with a slice of argument pointers.
- Manual registration builder for one or more UDFs via `xlfRegister`.
- Exported static thunks for primitive signatures.
- Initial conversions:
  - `f64`;
  - `bool`;
  - integer types with checked conversion;
  - `Result<T, ExcelError>`.
- Example `RUST.ADD` function.
- Panic containment and a minimal diagnostic log file.

### Exit criteria

The example XLL loads in 64-bit desktop Excel, appears in the Function Wizard, calculates correctly, unloads safely, and survives deliberate Rust panics without crashing Excel.

## M2 — Safe Excel values and owned return memory

### Deliverables

- Raw `XLOPER12` wrapper and high-level `ExcelValue` enum.
- Correct handling of:
  - UTF-16 length-prefixed strings;
  - Excel errors;
  - empty and missing values;
  - rectangular `xltypeMulti` arrays;
  - references and simple references;
  - Excel-allocated versus add-in-allocated values.
- `FromExcel` and `IntoExcel` traits.
- Single owned-return allocation model released by `xlAutoFree12`.
- Dynamic array examples returning vectors/matrices.
- Property and boundary tests for conversion and ownership code.

### Recommended memory strategy

Use a self-contained heap allocation for each returned complex value. The exposed `XLOPER12` should point into backing storage owned by the same allocation. Mark it for add-in freeing and recover the allocation in `xlAutoFree12` through a stable header/container layout. Avoid a global pointer registry unless Excel ABI constraints make the self-describing allocation impractical.

### Exit criteria

Strings and arrays can be returned repeatedly under recalculation and workbook close without leaks, double frees, use-after-free, or corruption in stress tests.

## M3 — Procedural macros and ergonomic registration

### Deliverables

- `#[excel_function]` macro supporting:
  - Excel-visible name;
  - category;
  - function and argument descriptions;
  - help topic;
  - volatile/thread-safe/macro-type flags;
  - aliases.
- `#[excel_command]` macro.
- Exact native thunk generation for the annotated Rust signature.
- Compile-time rejection of unsupported exported types.
- A registration inventory or explicit `excel_addin!` aggregation mechanism.
- Optional parameters and configurable missing/empty semantics.
- Generated metadata tests and compile-fail UI tests.

### Open architectural choice

Automatic distributed registration is convenient but can make link behavior opaque. Compare:

1. `inventory`/link-section discovery;
2. an explicit `excel_addin!(functions...)` root macro;
3. proc-macro generation of a known registry module through a build step.

Start with the most deterministic option, even if slightly more verbose.

### Exit criteria

A user can create a normal Rust function, annotate it, build an XLL, and receive correct Excel registration and conversions without writing ABI glue.

## M4 — Commands, Excel API calls and execution contexts

### Deliverables

- Typed wrappers for commonly needed Excel calls:
  - caller and add-in name;
  - coercion;
  - sheet/workbook identity;
  - status bar and calculation commands;
  - registration/unregistration.
- `ExcelReference` and caller information abstractions.
- Registered commands and optional keyboard shortcuts.
- Typed execution contexts, for example:

```rust
pub struct WorksheetContext<'a> { /* restricted */ }
pub struct ThreadSafeContext<'a> { /* no Excel calls */ }
pub struct MacroContext<'a> { /* main-thread command API */ }
```

- Runtime checks that reject invalid Excel calls from multi-threaded calculation.
- Main-thread dispatcher prototype using a hidden/message-only Win32 window or queued macro mechanism.

### Exit criteria

The API makes common legal operations easy and common illegal operations difficult. Thread-safe UDF examples run under Excel multi-threaded recalculation without invoking forbidden callbacks.

## M5 — Production hardening, diagnostics and packaging baseline

### Deliverables

- Structured logging with opt-in file and debugger sinks.
- Clear conversion, registration and lifecycle errors.
- Build metadata and diagnostic report function.
- Excel-version and architecture detection.
- Stress workbooks and automated smoke-test harness where practical.
- Example projects for scalars, arrays, commands and errors.
- Reproducible release artifacts for x64.
- Basic packaging command that copies/renames the built DLL to `.xll` and includes optional configuration/resources.
- Security review of unsafe blocks, exported symbols, panic behavior and allocation paths.

### Exit criteria

The library is suitable for real internal add-ins and failures can be diagnosed without attaching a native debugger.

## M6 — Native asynchronous functions

### Deliverables

- Excel native async registration and handle types.
- `xlAsyncReturn` wrapper.
- Proc-macro support for `async fn` or a builder-level async adapter.
- Pluggable executor abstraction:
  - default lightweight worker pool;
  - optional Tokio adapter.
- Cancellation registry tied to workbook/add-in shutdown.
- Main-thread completion/dispatch guarantees.
- Examples for delayed calculation and I/O-style workloads.

### Exit criteria

Async functions return control immediately, update the correct cell safely, cancel on shutdown, and do not block Excel's UI thread.

## M7 — Optional COM, RTD and Ribbon ecosystem

These features should live outside the core crates because they are Windows-specific and substantially increase binary and conceptual complexity.

### `excel-api-com`

- Access to the Excel COM object model through the `windows` crate.
- COM initialization and apartment-state helpers.
- Rust traits wrapping RTD server topics and update callbacks.
- Stream/channel adapters for RTD.
- Ribbon extensibility interfaces and XML loading.
- Ribbon callback dispatch.
- Optional task-pane experiments using WebView2 or native Win32 UI.

### Risks

- COM registration and deployment complexity.
- STA/main-thread requirements.
- Excel version differences.
- Callback lifetime and reference-count correctness.
- High testing cost compared with C API UDFs.

### Exit criteria

Keep these crates experimental until at least one production-quality RTD server and one Ribbon example have been exercised across supported Excel versions.

## M8 — Tooling and parity extensions

### Candidate tools

- `cargo excel-api new` project template.
- Manifest-driven add-in metadata.
- One-file or few-file packing, potentially through PE resources or an extraction bootstrapper.
- Function-reference Markdown/HTML generation from registration metadata.
- Workbook test harness and headless/automated Excel smoke tests.
- Code signing integration documentation.
- IntelliSense helper research.
- Native dependency packaging and load-path management.

### IntelliSense position

Excel-DNA's IntelliSense experience is valuable but is not part of the supported Excel C API. Treat it as a separate companion project after core registration metadata is stable. A possible implementation would expose metadata from the XLL and use a separate Windows UI Automation helper process or COM add-in to render function and argument assistance.

## Capability priority table

| Capability group | Core project? | Earliest milestone | Relative effort | Main risk |
|---|---:|---:|---:|---|
| XLL lifecycle | Yes | M1 | Low-medium | ABI/export correctness |
| Excel C API calls | Yes | M1/M4 | Medium | context/thread restrictions |
| Primitive UDFs | Yes | M1 | Medium | thunk ABI |
| Strings and arrays | Yes | M2 | High | ownership and freeing |
| Registration macros | Yes | M3 | High | generated exports and diagnostics |
| Commands/macros | Yes | M4 | Medium | main-thread execution |
| Multi-threaded UDFs | Yes | M4 | Medium | accidental forbidden calls |
| Logging/diagnostics | Yes | M1/M5 | Medium | avoiding reentrancy |
| Async UDFs | Yes, later | M6 | High | lifetime, cancellation, dispatch |
| Object handles | Optional core extension | M5+ | Medium | stale handles and workbook lifecycle |
| RTD | Optional crate | M7 | Very high | COM correctness |
| Ribbon | Optional crate | M7 | Very high | COM/UI deployment |
| Custom task panes | Optional/experimental | M7+ | Very high | UI hosting complexity |
| Packing | Tooling | M5/M8 | High | native extraction and security |
| IntelliSense overlay | Companion project | M8+ | Very high | unsupported UI integration |
| .NET loading/attributes | No | Never | N/A | contrary to Rust-native mission |

## Testing strategy

### Unit tests

- All conversion traits.
- UTF-16 boundary cases.
- Error and missing/empty distinctions.
- Registration type-string generation.
- Allocation ownership transitions.

### ABI tests

- Structure size, alignment and offsets against a C helper built from authoritative headers.
- Calling-convention smoke tests.
- Export-name inspection with `dumpbin` or `llvm-objdump`.

### Excel integration tests

- Load/unload cycles.
- Function Wizard metadata.
- Scalar and array calculations.
- Recalculation loops.
- Workbook close while functions/tasks are active.
- Multi-threaded recalculation.
- Deliberate panic and error paths.

### Long-running stress tests

- Repeated string and multi-array returns.
- Multiple workbooks and add-in reloads.
- Async cancellation and shutdown.
- Handle-table churn if object handles are implemented.

## Initial issues to create

1. Define supported Excel and Windows versions.
2. Establish raw `XLOPER12` bindings and ABI tests.
3. Export minimal XLL lifecycle functions.
4. Resolve and wrap `Excel12v`.
5. Register a manually declared `RUST.ADD` UDF.
6. Design owned return memory and `xlAutoFree12` behavior.
7. Implement primitive `FromExcel`/`IntoExcel` conversions.
8. Implement UTF-16 strings.
9. Implement rectangular arrays and dynamic-array example.
10. Prototype `#[excel_function]` thunk generation.
11. Define legal API operations by execution context.
12. Add Windows CI and artifact build.

## Definition of a credible `1.0`

A `1.0` release does not need full Excel-DNA feature parity. It should guarantee:

- stable core value and registration APIs;
- safe scalar, string, error, missing and array conversions;
- reliable x64 XLL lifecycle and unloading;
- procedural-macro UDF and command registration;
- documented thread-safety and Excel-call restrictions;
- robust diagnostics and reproducible packaging;
- tested support for specified Excel versions;
- a clear compatibility policy for future Excel and Rust releases.

RTD, Ribbon, task panes, packing, and IntelliSense can evolve independently without blocking a stable Rust-native C API foundation.

## Reference sources

- Excel-DNA project: <https://github.com/Excel-DNA/ExcelDna>
- Excel-DNA documentation: <https://excel-dna.net/docs/>
- Microsoft Excel XLL SDK overview: <https://learn.microsoft.com/office/client-developer/excel/welcome-to-the-excel-software-development-kit>
- Microsoft Excel C API reference: <https://learn.microsoft.com/office/client-developer/excel/excel-c-api>
