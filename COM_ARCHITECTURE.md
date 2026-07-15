# COM Architecture Boundary

## Current scope

M18 establishes only the COM boundary needed to evaluate RTD. It does not add
general COM, Ribbon, or Office object-model support to `excel-api`.

A COM RTD server is activated by ProgID/CLSID, created through a class factory,
and governed by COM apartments, reference counting, marshaling, and server
unload rules. These obligations are independent of XLL registration and the
Excel C API callback table.

## M18.1 prototype server surface

A separate in-process prototype DLL exports the normal COM class-server
entry points, including `DllGetClassObject` and `DllCanUnloadNow`, and implement
`IUnknown`, `IClassFactory`, and Excel's Automation-compatible `IRtdServer`
interface with the verified IID/type-library contract. Registration is handled
by reversible PowerShell scripts rather than optional DLL exports or build-time
mutation.

The initial registration plan is per-user ProgID/CLSID registration under the
user classes hive. A packaged installer may later offer per-machine
registration with explicit elevation. Registration-free COM remains deferred
because activation manifests are controlled by the Excel host application,
not by an arbitrary XLL.

An in-process DLL must match Excel's bitness. The selected first target is
64-bit Excel and x86_64-pc-windows-msvc. A local out-of-process COM server is
the future isolation/cross-bitness alternative.

## Threading and marshaling

The prototype registers `ThreadingModel=Apartment`, subject to live
compatibility results. This is not a claim that RTD callbacks arrive on
Excel's main thread. COM chooses/delivers the call according to the caller and
registered apartment model.

Every thread that performs COM work initializes COM appropriately. Raw
interface pointers remain in their owning apartment. When a callback interface
must be used from another apartment, the server obtains a proxy through
standard COM marshaling or the Global Interface Table and releases it in the
correct apartment. Owned producer values cross threads through a bounded Rust
channel; neither XLOPER12 nor borrowed Automation data crosses threads.

The callback supplied to `ServerStart` is registered in the COM Global
Interface Table in the caller's apartment. The producer initializes MTA and
obtains its own proxy from the GIT; no raw interface pointer crosses threads.
Shutdown joins the producer before revoking the cookie. No project lock is held
while calling `UpdateNotify`. Every exported/vtable boundary contains panics
and maps invalid pointers, invalid state, allocation failure, unsupported
interfaces, and internal failure to explicit HRESULTs.

The raw ABI source is the locally registered Excel type library
`{00020813-0000-0000-C000-000000000046}` version 1.9, LCID 0, audited with
`LoadTypeLibEx`/`ITypeInfo`. Its hidden dual-interface vtables pin the
`IRtdServer` and `IRTDUpdateEvent` IIDs, inherited `IDispatch` slots, method
order, `stdcall` convention, widths, and parameter directions. Details and the
reproducible audit output are in
`docs/research/m18-1-excel-rtd-typelib-abi.md`.

## Dual-binary rule

The production XLL and RTD server remain separate binaries. Combining XLL
lifecycle exports with a COM class factory would make two independent hosts
and unload protocols share one module. Reducing artifact count does not prove
that `xlAutoClose`, COM reference counts, `DllCanUnloadNow`, background
producers, and Excel termination have reached the same quiescent boundary.

The default `excel-api`, `excel-api-sys`, `excel-api-macros`, and minimal XLL
builds gain no unconditional Windows COM dependency or RTD registration side
effect.

## Trust and deployment

COM registration, Office policy, XLL trust, code signing, and RTD formula
activation are separate observations in the validation matrix. Deployment
must record ProgID/CLSID, server path, bitness, registration scope, signing
state, Excel channel/build, and rollback result without committing private
machine paths.

The project will not modify Excel Trusted Locations, organization-wide COM
policy, or macro/XLM settings to make a prototype pass.

## M18.2 activation and cleanup refinement

The callback cookie is not cleared until `RevokeInterfaceFromGlobal` succeeds.
Failure enters `CallbackRevocationPending`; repeated `ServerTerminate` retries
the same cookie and `DllCanUnloadNow` remains `S_FALSE`. Producer joins are
inspected, and a panicked join is controlled `E_UNEXPECTED`, not a discarded
error. Producer and committed-notification counters are RAII-owned.

Activation diagnostics record class-factory and dual-interface negotiation
without pointer values. Registration inspection covers HKCU, HKLM, and merged
HKCR in both 64-bit and 32-bit views. Conflicts are reported rather than
overwritten in machine scope.
