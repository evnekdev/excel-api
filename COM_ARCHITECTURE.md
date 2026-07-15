# COM Architecture Boundary

## Current scope

M18 establishes only the COM boundary needed to evaluate RTD. It does not add
general COM, Ribbon, or Office object-model support to `excel-api`.

A COM RTD server is activated by ProgID/CLSID, created through a class factory,
and governed by COM apartments, reference counting, marshaling, and server
unload rules. These obligations are independent of XLL registration and the
Excel C API callback table.

## Required future server surface

A separate in-process prototype DLL must export the normal COM class-server
entry points, including `DllGetClassObject` and `DllCanUnloadNow`, and implement
`IUnknown`, `IClassFactory`, and Excel's Automation-compatible `IRtdServer`
interface with the verified IID/type-library contract. Optional registration
entry points must be reversible and must never silently change machine-wide
policy.

The initial registration plan is per-user ProgID/CLSID registration under the
user classes hive. A packaged installer may later offer per-machine
registration with explicit elevation. Registration-free COM remains deferred
because activation manifests are controlled by the Excel host application,
not by an arbitrary XLL.

An in-process DLL must match Excel's bitness. The selected first target is
64-bit Excel and x86_64-pc-windows-msvc. A local out-of-process COM server is
the future isolation/cross-bitness alternative.

## Threading and marshaling

The prototype will begin with `ThreadingModel=Apartment`, subject to live
compatibility results. This is not a claim that RTD callbacks arrive on
Excel's main thread. COM chooses/delivers the call according to the caller and
registered apartment model.

Every thread that performs COM work initializes COM appropriately. Raw
interface pointers remain in their owning apartment. When a callback interface
must be used from another apartment, the server obtains a proxy through
standard COM marshaling or the Global Interface Table and releases it in the
correct apartment. Owned producer values cross threads through a bounded Rust
channel; neither XLOPER12 nor borrowed Automation data crosses threads.

No project lock is held while calling `UpdateNotify` or returning from an
`IRtdServer` method. All FFI methods contain panics and translate failures to
documented HRESULT/RTD results once those exact mappings are verified.

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
