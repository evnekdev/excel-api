# Excel-DNA Capability Map for a Rust-Native Excel API

## Purpose

This document gives a bird's-eye view of the functionality provided by the Excel-DNA ecosystem and evaluates how each capability could be reproduced in Rust.

The goal is **not** to port Excel-DNA line by line. Excel-DNA is primarily a bridge between Excel and the .NET runtime, while `excel-api` should be a Rust-native toolkit built directly on the Excel C API and selected Windows APIs.

The feasibility labels used below are:

| Label | Meaning |
|---|---|
| **Direct** | Naturally implementable with the Excel C API and Rust FFI. |
| **Practical** | Feasible in Rust, but requires a substantial runtime, tooling, or Windows integration layer. |
| **Possible, specialized** | Technically possible, but expensive, fragile, or outside the core mission. |
| **Not applicable** | Exists mainly to support .NET and should not be reproduced as such. |

## Executive summary

The highest-value Excel-DNA capabilities are reproducible in Rust:

- native XLL loading and lifecycle;
- worksheet function and command registration;
- conversion between Excel values and Rust types;
- arrays and dynamic-array results;
- thread-safe and macro-type functions;
- asynchronous and streaming calculations;
- Excel API calls from Rust;
- ribbon, COM, RTD, and task-pane integration;
- packaging, diagnostics, and documentation generation.

The hardest areas are not the basic Excel C API. They are:

1. safe ownership of `XLOPER12` values and Excel-allocated memory;
2. automatic generation of ABI-compatible exported thunks for arbitrary Rust signatures;
3. Excel main-thread dispatch and recalculation semantics;
4. COM, Ribbon XML, custom task panes, and Windows UI integration;
5. IntelliSense overlays, which rely on Windows UI Automation rather than a supported Excel UDF API;
6. producing a polished one-file add-in packaging experience.

## Capability matrix

### 1. Native XLL host and add-in lifecycle

| Excel-DNA capability | What it provides | Rust feasibility | Proposed Rust strategy | Priority |
|---|---|---:|---|---:|
| Native `.xll` add-in | Excel loads a native DLL with the `.xll` extension | **Direct** | Build a `cdylib`; export required XLL entry points with `#[unsafe(no_mangle)] extern "system"`; provide x86 and x64 targets | P0 |
| `xlAutoOpen` | Add-in initialization and function registration | **Direct** | Runtime-owned registry emitted by procedural macros or explicit builder calls | P0 |
| `xlAutoClose` | Add-in shutdown | **Direct** | Unregister functions, stop worker threads, release COM/UI resources | P0 |
| `xlAutoAdd` / `xlAutoRemove` | Hooks for Add-in Manager actions | **Direct** | Optional lifecycle trait and generated exports | P1 |
| `xlAddInManagerInfo12` | Display add-in name/description in Excel | **Direct** | Static metadata generated at compile time | P0 |
| `xlAutoFree12` | Return-memory cleanup callback | **Direct but safety-critical** | Central allocation registry or self-describing owned return blocks | P0 |
| Excel version and architecture handling | Compatibility across Excel versions and bitness | **Practical** | Runtime capability detection; separate x86/x64 artifacts; feature gates for newer APIs | P1 |
| Crash isolation and guarded boundaries | Prevent managed exceptions from escaping into Excel | **Direct conceptually** | Wrap every export with `catch_unwind`; convert panics to Excel errors and diagnostic logs; document `panic = "unwind"` requirement | P0 |

### 2. Raw Excel C API access

| Excel-DNA capability | What it provides | Rust feasibility | Proposed Rust strategy | Priority |
|---|---|---:|---|---:|
| `Excel12` / `Excel12v` calls | Invoke Excel worksheet, macro-sheet, and command functions | **Direct** | `excel-api-sys` bindings plus a safe variadic-vector wrapper around `Excel12v` | P0 |
| `XLOPER12` representation | Universal Excel value structure | **Direct but safety-critical** | `#[repr(C)]` raw bindings and a separate safe `ExcelValue` abstraction | P0 |
| Function IDs and constants | `xlGetName`, `xlfRegister`, `xlCoerce`, etc. | **Direct** | Generate constants from Microsoft headers or maintain a verified definitions module | P0 |
| Caller information | Cell, sheet, workbook, command-bar context | **Direct** | Typed wrappers around `xlfCaller`, `xlSheetNm`, `xlGetName`, and reference values | P1 |
| Excel error values | `#VALUE!`, `#N/A`, `#REF!`, etc. | **Direct** | `ExcelError` enum with exact C API discriminants | P0 |
| Missing and empty values | Distinguish omitted, empty, and nil arguments | **Direct** | Dedicated `Missing`, `Empty`, and `ExcelValue` variants; `Option<T>` policy configurable | P0 |
| References and ranges | Sheet references without immediately reading values | **Direct** | Safe wrappers for `xltypeRef` and `xltypeSRef`; explicit lifetime/copy semantics | P1 |
| Coercion | Convert Excel references/values using Excel rules | **Direct** | `ExcelValue::coerce(...)` and typed conversion traits using `xlCoerce` | P1 |
| Sheet/workbook operations | Call Excel commands and worksheet functions | **Direct with restrictions** | Separate calculation-safe and macro-context APIs; prevent invalid calls from thread-safe UDFs | P1 |

### 3. Function and command registration

| Excel-DNA capability | What it provides | Rust feasibility | Proposed Rust strategy | Priority |
|---|---|---:|---|---:|
| Automatic UDF discovery | Finds attributed .NET methods | **Practical** | Link-time distributed registration using `inventory`, generated registry modules, or explicit `excel_addin!` macro | P0 |
| `ExcelFunction` metadata | Name, description, category, help topic, volatility, thread-safety, macro-type flags | **Direct** | `#[excel_function(...)]` procedural macro generating registration descriptors | P0 |
| `ExcelArgument` metadata | Argument names and descriptions | **Direct** | Parameter-level helper syntax in the proc macro or companion attributes | P0 |
| Excel commands/macros | Register commands callable by menu, shortcut, or Excel | **Direct** | `#[excel_command(...)]`; register with command type text and generated ABI thunk | P1 |
| Function Wizard integration | Names, descriptions, categories, help links | **Direct** | Populate `xlfRegister` fields exactly; generate optional help topics | P0 |
| Explicit registration pipeline | Transform registrations before publishing | **Practical** | A `Registration` builder and middleware pipeline operating on descriptors before `xlfRegister` | P1 |
| Registration transformations | Logging, caching, timing, suppression, wrapper generation | **Practical** | Attribute-driven proc-macro wrappers plus runtime middleware traits | P2 |
| Optional parameters | Defaults and omitted arguments | **Direct** | `Option<T>`, explicit default expressions in attributes, and conversion of `xltypeMissing` | P1 |
| Variadic/`params` arguments | Variable number of arguments | **Possible, specialized** | Fixed maximum arity at Excel ABI boundary or accept a trailing multi-cell range/array; true arbitrary ABI arity is not available | P3 |
| Nullable values | Managed nullable argument support | **Direct equivalent** | `Option<T>` and explicit `Nullable<T>` policy where empty and missing must differ | P1 |
| Suppress from Function Wizard | Hide helper functions | **Direct** | Registration flags/name conventions and optional no-register exports | P2 |
| Function aliases | Multiple Excel names for one implementation | **Direct** | Generate multiple descriptors pointing to one internal Rust function | P2 |

### 4. ABI thunk generation and Rust signature support

| Capability | Rust feasibility | Proposed strategy | Priority |
|---|---:|---|---:|
| Exported function per UDF | **Direct for known signatures** | Procedural macro emits `extern "system"` thunk with an unmangled export name | P0 |
| Arbitrary argument counts | **Practical with generated code** | Proc macro generates exact native signature for each function; impose Excel's argument-count limits | P0 |
| Automatic type conversion | **Practical** | `FromExcel` and `IntoExcel` traits invoked by generated thunk | P0 |
| Borrowed strings/arrays | **Practical but lifetime-sensitive** | Borrow only for duration of the call; use wrapper types such as `ExcelStr<'a>` and `ExcelArrayView<'a>` | P1 |
| Owned return values | **Direct but safety-critical** | Allocate result backing storage in one owned block released by `xlAutoFree12` | P0 |
| Generic Rust functions | **Not directly exportable** | Require monomorphized non-generic wrapper functions | P1 |
| Methods and closures | **Practical internally** | Register static generated thunks that dispatch to stored `fn` pointers or synchronized trait objects | P2 |
| User-defined structs | **Practical via conversion traits** | Users implement `FromExcel`/`IntoExcel`, derive macro later | P2 |
| Panic and error conversion | **Direct** | Support `Result<T, ExcelError>` and configurable mapping of application errors to `#VALUE!` or strings | P0 |

### 5. Excel-to-Rust type conversion

| Excel value/functionality | Rust representation | Feasibility | Notes | Priority |
|---|---|---:|---|---:|
| Number | `f64` | **Direct** | Native Excel numeric type | P0 |
| Integer | `i16`, `i32`, `i64`, `usize` | **Direct with validation** | Excel stores most values as `f64`; reject overflow/fractional input | P0 |
| Boolean | `bool` | **Direct** | Native Boolean XLOPER type | P0 |
| Text | `String`, `&str`, `ExcelStr` | **Direct** | Excel 12 strings are length-prefixed UTF-16 | P0 |
| Error | `ExcelError` | **Direct** | Preserve exact error values | P0 |
| Empty/missing | marker types / `Option<T>` | **Direct** | Policy must distinguish omitted and blank cells | P0 |
| Date/time | wrapper types or `chrono` feature | **Practical** | Excel serial dates plus workbook date-system concerns | P2 |
| Cell range values | `ExcelArray`, slices, matrix views | **Direct** | Rectangular `xltypeMulti`; row-major facade over Excel layout | P0 |
| Cell references | `ExcelReference` | **Direct** | Avoid accidental dereference/coercion | P1 |
| Dynamic arrays | owned rectangular result | **Direct** | Return `xltypeMulti`; modern Excel spills automatically | P0 |
| Jagged arrays | no native equivalent | **Possible by normalization** | Reject or pad into a rectangle | P2 |
| Arbitrary objects | handles/IDs | **Practical** | Return opaque strings/numbers and store Rust objects in a synchronized handle table | P2 |
| Big integers/decimals | strings or loss-checked numbers | **Practical** | Excel numeric precision remains the limiting factor | P3 |

### 6. Calculation behavior

| Excel-DNA capability | Rust feasibility | Proposed strategy | Priority |
|---|---:|---|---:|
| Volatile UDFs | **Direct** | Set registration type modifier and expose metadata option | P1 |
| Thread-safe UDFs | **Direct with API restrictions** | Set thread-safe registration flag; expose a restricted context that cannot call unsafe Excel APIs | P1 |
| Macro-type UDFs | **Direct** | Set macro-type modifier; mark API as main-thread-only | P1 |
| Cluster-safe UDFs | **Possible, specialized** | Registration flag only initially; real HPC compatibility requires strict function semantics and testing | P3 |
| Recalculation triggers | **Direct** | Wrappers for `xlcCalculateNow`, `xlcCalculateDocument`, and async completion APIs where appropriate | P2 |
| Caller-sensitive functions | **Direct** | Typed caller queries, with restrictions documented for MTR | P2 |
| Caching/memoization | **Practical** | Optional middleware using hashable normalized arguments and bounded concurrent caches | P2 |
| Cancellation | **Practical** | Cooperative cancellation token for async work; integrate Excel abort checks where supported | P2 |

### 7. Asynchronous and streaming functions

| Excel-DNA capability | What it provides | Rust feasibility | Proposed Rust strategy | Priority |
|---|---|---:|---|---:|
| Native Excel async UDFs | Function returns immediately and later supplies a result | **Practical** | Register async handle argument; spawn work on a runtime; complete via `xlAsyncReturn` | P2 |
| `Task<T>` functions | .NET task integration | **Not applicable directly** | Rust equivalent accepts `async fn` and generates runtime dispatch | P2 |
| Observable/streaming functions | Repeated updates through RTD | **Practical** | `Stream<Item = T>` adapter backed by RTD/COM or a native recalculation/update mechanism | P3 |
| Async runtime selection | Uses .NET scheduler/task infrastructure | **Practical** | Pluggable executor; default small thread pool; optional Tokio feature kept out of core | P2 |
| Main-thread marshaling | Schedule work that must run on Excel's UI thread | **Practical and important** | Hidden message-only window, COM marshaling, or Excel macro queue; expose `ExcelDispatcher` | P1 |
| Cancellation and workbook shutdown | Stop outstanding tasks | **Practical** | Runtime task registry keyed by async handles; cancellation on close/unregister | P2 |

### 8. RTD and real-time data

| Excel-DNA capability | Rust feasibility | Proposed strategy | Priority |
|---|---:|---|---:|
| RTD server implementation | **Practical, Windows-specific** | Implement Excel RTD COM interfaces using the `windows` crate; provide a safe topic server trait | P3 |
| Topic subscription lifecycle | **Practical** | Concurrent topic map and callback channel to Excel | P3 |
| Observable-to-RTD adapter | **Practical** | Adapt Rust `Stream` or broadcast channels to RTD topics | P3 |
| Throttled updates | **Direct once RTD exists** | Coalesce topic updates and respect Excel's refresh behavior | P3 |
| Out-of-process RTD | **Possible, specialized** | Separate COM local server executable; useful for crash isolation but much more deployment complexity | P4 |

### 9. Ribbon, menus, keyboard shortcuts, and UI

| Excel-DNA capability | Rust feasibility | Proposed strategy | Priority |
|---|---:|---|---:|
| Legacy menu/command integration | **Direct** | Excel C API command registration and menu APIs | P2 |
| Ribbon XML | **Practical, Windows-specific** | COM add-in interfaces via `windows`; embed or load custom UI XML | P3 |
| Ribbon callbacks | **Practical** | COM-visible callback object dispatching into Rust functions | P3 |
| Images/icons | **Practical** | Embedded resources and COM-compatible image return values | P3 |
| Custom task panes | **Possible, specialized** | COM custom task pane interfaces plus native WebView2 or Win32 child UI | P4 |
| Modeless windows/dialogs | **Practical** | Win32, `windows`, or a lightweight Rust GUI framework; marshal actions to Excel main thread | P3 |
| Keyboard shortcuts | **Direct for commands** | Supply shortcut metadata during command registration | P2 |
| Status bar and alerts | **Direct** | Safe wrappers around relevant Excel API functions | P2 |

### 10. COM integration and automation

| Excel-DNA capability | Rust feasibility | Proposed strategy | Priority |
|---|---:|---|---:|
| Access Excel COM object model | **Practical** | `windows` crate generated bindings; obtain application object from Excel context | P3 |
| Expose COM classes | **Practical but deployment-sensitive** | Implement IUnknown/IDispatch classes in Rust; registration-free COM where possible | P4 |
| COM add-in lifecycle | **Practical** | Implement `IDTExtensibility2` and related Office interfaces | P4 |
| Automation add-ins | **Possible, specialized** | COM-visible function object; lower priority than XLL UDFs | P4 |
| VBA interoperability | **Practical** | Export commands/UDFs and optionally COM automation objects | P3 |
| Registration-free deployment | **Possible** | Side-by-side manifests or runtime registration; requires extensive Windows-version testing | P4 |

### 11. IntelliSense and function help

| Excel-DNA capability | Rust feasibility | Proposed strategy | Priority |
|---|---:|---|---:|
| Function Wizard descriptions | **Direct** | Supply names/descriptions/help topic through `xlfRegister` | P0 |
| Native formula autocomplete metadata for UDFs | **Not exposed by Excel** | No supported C API path is known | — |
| In-sheet IntelliSense overlay | **Possible, specialized and fragile** | Separate Windows UI Automation component that tracks Excel UI and overlays function/argument help | P4 |
| Metadata source | **Direct** | Reuse proc-macro registration descriptors; serialize optional JSON/XML index | P2 |
| VBA function help import | **Possible** | Read workbook metadata or external files; not core | P4 |
| Rich argument controls | **Possible, specialized** | Custom overlay UI; date pickers, enum lists, links | P5 |

Excel-DNA IntelliSense is a separate project and uses Windows UI Automation because Excel does not expose a supported UDF IntelliSense API. A Rust implementation should therefore be treated as an optional companion add-in, not part of the foundational C API crate.

### 12. Packaging and deployment

| Excel-DNA capability | Rust feasibility | Proposed strategy | Priority |
|---|---:|---|---:|
| One-file packed XLL | **Practical** | Embed resources into the XLL or append a signed resource container; extract/load dependencies securely | P3 |
| Configuration file (`.dna`) | **Not necessary in the same form** | Prefer compile-time Rust metadata; optionally support TOML for packaging and runtime configuration | P1 |
| Dependency packing | **Practical** | Static linking where licenses permit; otherwise embed DLLs and extract with hashes/versioned directories | P3 |
| 32-bit and 64-bit packages | **Direct** | CI matrix for `i686-pc-windows-msvc` and `x86_64-pc-windows-msvc` | P1 |
| Code signing | **Practical** | Document `signtool`; CI integration with protected certificate service | P2 |
| NuGet integration | **Not applicable** | Publish Rust crates on crates.io and release XLL templates/binaries on GitHub | P1 |
| Project templates | **Practical** | `cargo-generate` template and later `cargo excel-api new/build/package` subcommands | P1 |
| Add-in installation | **Practical** | Manual `.xll`, trusted-location guidance, optional installer/MSIX later | P2 |
| Self-update | **Possible, specialized** | Separate updater process; avoid in core due to enterprise security concerns | P5 |

### 13. Diagnostics, logging, and debugging

| Excel-DNA capability | Rust feasibility | Proposed strategy | Priority |
|---|---:|---|---:|
| Registration diagnostics | **Direct** | Structured errors containing function name, type text, and Excel return code | P0 |
| Log file | **Direct** | `tracing` facade with file/debug-output subscribers configured by environment or TOML | P1 |
| Unhandled exception reporting | **Direct equivalent** | Panic hook plus `catch_unwind`; optional minidump integration | P1 |
| Debugger launch/development mode | **Practical** | Build profile helpers; optional wait-for-debugger flag; launch Excel from cargo task | P2 |
| Excel-visible error details | **Practical** | Configurable error strings, diagnostics worksheet command, or handle-based error lookup | P2 |
| Telemetry | **Possible but opt-in only** | Explicit feature and application-owned sink; none in core by default | P5 |

### 14. Documentation generation

| Excel-DNA capability | Rust feasibility | Proposed strategy | Priority |
|---|---:|---|---:|
| Generate function reference from annotations | **Direct** | Proc macro emits metadata; build tool outputs Markdown/HTML/JSON | P2 |
| CHM help generation | **Possible but legacy** | Generate HTML and optionally invoke external CHM tooling; not a default target | P4 |
| Help-topic links from Function Wizard | **Direct** | Register local or web URLs generated from metadata | P2 |
| Examples and return documentation | **Direct** | Extend `#[excel_function]` metadata and/or extract Rust doc comments | P2 |
| API docs for Rust developers | **Direct** | rustdoc plus mdBook or MkDocs documentation | P1 |

### 15. Add-in extensibility and plugin loading

| Excel-DNA capability | Rust feasibility | Proposed strategy | Priority |
|---|---:|---|---:|
| Load multiple managed assemblies | **Not applicable directly** | Rust code is normally linked into the XLL | — |
| Dynamically load native plugins | **Possible but unsafe/complex** | Stable C ABI plugin interface using `abi_stable`-style patterns or generated C headers | P4 |
| Load add-in code from bytes | **Possible, specialized** | Manual PE loading is risky and unnecessary; prefer normal Windows loader and signed files | P5 |
| Runtime function discovery | **Practical within linked binary** | Distributed registration or generated registry | P0 |
| Cross-language functions | **Practical** | Rust wrapper can call C ABI libraries; Python/.NET embedding should be separate integration crates | P4 |

### 16. Security and enterprise deployment

| Capability | Rust feasibility | Proposed strategy | Priority |
|---|---:|---|---:|
| Minimize runtime dependencies | **Strong Rust advantage** | Statically link Rust runtime and keep core XLL self-contained | P0 |
| Signed binaries | **Practical** | Reproducible release workflow and Authenticode signing guidance | P2 |
| DLL search-path safety | **Direct** | Use absolute paths, safe loading flags, hashes, and versioned extraction paths | P1 |
| Macro security compatibility | **Practical** | Document trusted publishers/locations and enterprise deployment options | P2 |
| Memory safety across FFI | **Requires discipline** | Keep raw pointers isolated in `excel-api-sys`; test allocation, strings, arrays, and callbacks heavily | P0 |
| Supply-chain control | **Practical** | Minimal dependencies, locked CI, advisories, provenance/SBOM in releases | P2 |

## What should not be copied from Excel-DNA

The following Excel-DNA concerns exist because its primary purpose is hosting .NET inside Excel. They should not become design requirements for the Rust core:

- CLR discovery and hosting;
- loading .NET assemblies and resolving managed dependencies;
- .NET reflection-based method discovery;
- C#/F#/Visual Basic language-specific transformations;
- NuGet/MSBuild integration as the primary build system;
- managed AppDomain or AssemblyLoadContext behavior;
- managed exception and garbage-collection integration.

Rust-native equivalents should be designed around compile-time metadata, procedural macros, explicit ownership, and Cargo.

## Recommended scope boundary

The project should be split conceptually into layers:

1. **Core, portable Rust design**
   - value types;
   - conversion traits;
   - registration descriptors;
   - error model;
   - procedural macros.

2. **Windows Excel XLL runtime**
   - C API bindings;
   - exported lifecycle functions;
   - registration;
   - memory ownership;
   - main-thread dispatch;
   - asynchronous completion.

3. **Optional Windows integration crates**
   - COM automation;
   - Ribbon;
   - RTD;
   - task panes;
   - IntelliSense overlay.

4. **Developer tooling**
   - templates;
   - packaging;
   - metadata/documentation generation;
   - test harnesses;
   - CI release workflows.

## Sources and reference projects

This map is based on the public Excel-DNA ecosystem and should be revised as implementation work uncovers finer details:

- Excel-DNA core: <https://github.com/Excel-DNA/ExcelDna>
- Excel-DNA documentation: <https://excel-dna.net/docs/>
- Excel-DNA Registration helpers: <https://github.com/Excel-DNA/Registration>
- Excel-DNA IntelliSense: <https://github.com/Excel-DNA/IntelliSense>
- ExcelDnaDoc: <https://github.com/Excel-DNA/ExcelDnaDoc>
- Excel-DNA samples: <https://github.com/Excel-DNA/Samples>
- Microsoft Excel XLL SDK documentation: <https://learn.microsoft.com/en-us/office/client-developer/excel/excel-xll-sdk>

## Maintenance rule

When a new Excel-DNA feature is identified, add it to the matrix before implementing it. Record:

- whether the feature belongs to the Excel C API, Windows/COM integration, or .NET hosting;
- whether a Rust-native equivalent is desirable;
- its safety and deployment risks;
- its target crate and roadmap phase.
