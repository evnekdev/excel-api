# COM Automation extraction roadmap

Status: architecture decision and staged implementation roadmap.

Baseline reviewed: `origin/master` at
`b159ad7d7e0fbb7c033101cebca0a932877e6224` (`feat: complete excel-com migration support`).

This document is intentionally architecture-first. It does not change the
existing `excel-com` public API, does not change the PyroApp migrator, and does
not make a new generic COM API stable.

## Executive decision

Extracting a reusable `com-automation` foundation is justified, but it should
be done as a staged internal refactoring rather than by moving the current
public Automation types out of `excel-com` in one step.

The repository already contains a real generic COM kernel inside `excel-com`:

- explicit STA ownership and same-thread enforcement;
- exactly-one-reference COM pointer ownership with `AddRef`/`Release`;
- private `QueryInterface` and canonical `IUnknown` identity;
- `IDispatch` invocation and structured `EXCEPINFO` handling;
- initialized `VARIANT` ownership and exactly-once `VariantClear` cleanup;
- `BSTR` and `SAFEARRAY` ownership;
- optional/missing Automation arguments;
- private `IEnumVARIANT` ownership;
- a working manually implemented COM object for `IMessageFilter`;
- conservative COM retry infrastructure.

Those mechanisms are not inherently Excel-specific. Future BYREF support,
connection points, event sinks, marshaling, generic activation, and early-bound
interfaces would otherwise add more generic Windows COM machinery to an Excel
crate and then have to be reimplemented for Word, PowerPoint, AutoCAD, or other
Automation servers.

However, the existing implementation also deliberately mixes generic COM
mechanics with Excel policy. In particular, activation is hard-coded to
`Excel.Application`, invocation diagnostics infer object types from Excel
member IDs, retry safety contains Excel-specific member exceptions, Range
value conversion has Excel-specific shape and `Value`/`Value2` policy, session
shutdown tracks Excel's process, and the object-model registry is Excel data.
Those concerns must remain above the generic layer.

The recommended dependency direction is therefore:

```text
                  developer-time metadata/tools
                  -----------------------------
                  Excel typelib / inventories
                            |
                            v
+--------------------------------------------------------------+
| excel-com                                                    |
|                                                              |
| curated Excel API                                            |
| Application / Workbook / Worksheet / Range / Chart / Pivot   |
| Excel enums, guards, process/session policy, retry policy     |
| Excel generated low-level member metadata                    |
+-------------------------------+------------------------------+
                                |
                                v
+--------------------------------------------------------------+
| com-automation                                               |
|                                                              |
| apartment + COM identity/ownership                           |
| IDispatch call engine + call frames                          |
| BSTR / VARIANT / SAFEARRAY                                   |
| IEnumVARIANT                                                 |
| generic activation / QI primitives                           |
| BYREF / OUT / INOUT                                          |
| sink/server primitives + connection points                   |
| explicit marshaling                                          |
| generic COM diagnostics                                      |
+-------------------------------+------------------------------+
                                |
                                v
                         Windows COM / OLE
```

The generic crate should initially be an implementation dependency of
`excel-com`, with a deliberately narrow or unpublished public surface. It
should become a separately reusable public crate only after its abstractions
have survived real Excel BYREF events and cross-apartment use.

## Evidence reviewed

The decision above is based on the current code and the repository's existing
COM research, not only on the desired future shape.

Important current implementation areas include:

- `excel-com/src/internal/apartment.rs`;
- `excel-com/src/internal/com_ptr.rs`;
- `excel-com/src/automation/{bstr,variant,safearray,value,argument}.rs`;
- `excel-com/src/automation/{dispatch,invocation,enumerator}.rs`;
- `excel-com/src/automation/{message_filter,retry}.rs`;
- `excel-com/src/error.rs`;
- `excel-com/src/excel/mod.rs` and `excel-com/src/excel/session.rs`;
- `excel-com/src/object_model/registry.rs`;
- `metadata/excel-object-model/`;
- `tools/excel-object-model-inventory/`;
- `docs/research/excel-com/01-com-automation-foundations.md`;
- `docs/research/excel-com/05c-client-implementation-knowledge-base.md`;
- `knowledge/excel-object-model/generated/client-implementations/`;
- `docs/architecture/excel-com-migration-support.md`;
- `excel-com/tests/migration_support_live.rs`.

The current migration work is especially important as a compatibility anchor:
the crate now has real Excel 16.0 evidence for private-process isolation,
Formula2, genuine CSE arrays, UsedRange/SpecialCells, workbook/global and
worksheet-local names, `.xlsm` to `.xlsx` SaveAs/reopen, VBA absence, and exact
owned-process shutdown. Generic refactoring must not regress that behavior.

## 1. Current low-level architecture

### 1.1 Apartment ownership

`ComApartment::sta` owns one successful `CoInitializeEx(...,
COINIT_APARTMENTTHREADED)` call, records the creating Rust thread, is neither
`Send` nor `Sync`, checks same-thread use, and pairs its successful
initialization with `CoUninitialize` in `Drop`.

This is already a generic COM concept. The Excel-specific part is its current
error type and the fact that only STA construction is exposed.

### 1.2 COM interface ownership and identity

`ComPtr<T>` is an exactly-one-reference private owner around a non-null COM
interface pointer. `Clone` calls `AddRef`, `Drop` calls `Release`, and
`query_interface` turns a successful `QueryInterface` result into a new owned
reference. Its `Rc` marker prevents accidental `Send`/`Sync`.

`DispatchObject::same_object` obtains canonical `IUnknown` pointers and
compares those identities. That identity rule is generic COM and should remain
the basis of object identity after extraction.

The current marker types (`Dispatch`, `Unknown`, `EnumVariantInterface`) are
private implementation categories rather than a general typed-interface
system. That is adequate today, but future generated early-bound interfaces
will need a more explicit internal interface-ID/vtable model.

### 1.3 Automation invocation

The current `IDispatch` engine correctly centralizes important call-frame
mechanics:

- logical positional arguments are converted to COM reverse order;
- a property put supplies `DISPID_PROPERTYPUT`;
- result `VARIANT`, `DISPPARAMS`, `EXCEPINFO`, and `puArgErr` storage remain
  alive for the call;
- server exception text is copied before its BSTR allocations are released;
- the original HRESULT, exception SCODE, argument index, flags, attempts, and
  elapsed time are retained;
- retries occur only after an explicit safety classification.

The engine is not yet generic in policy:

- activation is `Excel.Application` only;
- every call resolves the member name through `GetIDsOfNames` rather than
  accepting a generated DISPID or a per-object cache;
- arbitrary named arguments are not modeled; the only named-argument case is
  the special property-put DISPID;
- `MemberDescriptor` contains the Excel-specific `MemberId`;
- retry safety includes special cases such as `excel.workbooks.add`;
- diagnostic object names are derived from `excel.*` ID prefixes;
- locale behavior is fixed inside the engine rather than represented by an
  explicit generic policy.

These are good seams for extraction: keep the physical call engine below and
move member identity, retry classification, object labels, and Excel locale
compatibility above it.

### 1.4 Automation values

`OwnedVariant`, `Bstr`, and `SafeArray` are already private ownership types.
The cleanup model is strong: initialized `VARIANT`s are cleared once, BSTR
allocation is paired with the OLE allocator, COM references transferred into
`VARIANT`s are separately counted, and owned SAFEARRAYs are destroyed once.

The public `AutomationValue` is intentionally narrower and more semantic. It
is pointer-free and supports Empty, Null, Bool, Number, Text, Error, OLE Date,
Currency, and rectangular arrays. That is a good Excel-facing API, but it is
not a lossless model of every Automation `VARTYPE`.

Current limitations relative to a general Automation engine include:

- no `VT_BYREF` representation;
- no OUT/INOUT call slots;
- no exact public distinction among all signed/unsigned integer widths and
  `R4`/`R8`;
- no general owned object value for `VT_DISPATCH`/`VT_UNKNOWN`;
- SAFEARRAY creation is centered on `VT_VARIANT` and the Excel public array is
  rank two;
- Range shape validation and `Value` versus `Value2` date policy are mixed
  into the conversion layer.

The existing public `AutomationValue` should therefore stay in `excel-com`
during extraction. The generic layer should first own ABI storage and a
lossless internal value/call-slot model. A later release may decide whether a
reusable semantic value type is worth exposing publicly.

### 1.5 Enumeration

The private `EnumVariant` wrapper already owns `IEnumVARIANT`, calls `Next`,
handles `S_OK`/`S_FALSE`, and converts yielded `VT_UNKNOWN` or `VT_DISPATCH`
objects through `QueryInterface`.

The COM owner/vtable and one-item iteration are generic. The helper that starts
from an Excel `_NewEnum` member and labels errors with an Excel collection is
Excel-specific.

### 1.6 Existing COM-server precedent

`ComMessageFilterGuard` already contains a manually implemented COM object with
an `IUnknown` prefix, `QueryInterface`, `AddRef`, `Release`, and an
`IMessageFilter` vtable. This is useful evidence that the crate has already
crossed the client-only boundary internally.

It is not yet a general sink/server substrate. Its allocation remains owned by
the Rust guard, and its reference count exists to honor COM while registration
is active. A connection-point event sink needs a stronger COM-owned lifetime
model, callback state, panic containment, and apartment rules. The message
filter implementation should be treated as an implementation precedent to
audit, not copied blindly into an event system.

### 1.7 Excel sessions remain an Excel concern

`OwnedApplication` and `AttachedApplication` intentionally encode semantics
that should not move into a generic crate:

- a crate-owned Excel process versus a shared active Excel registration;
- exact Excel PID observation through `Application.Hwnd`;
- the rule that `Drop` releases references but never silently calls `Quit`;
- exact-process `quit_and_wait` behavior;
- Excel diagnostics such as Version, workbook/window counts, Ready,
  Interactive, and calculation state;
- the rule that PyroApp-style work must create a private Excel instance and
  never attach to the user's interactive process.

The generic crate may provide activation and ROT primitives. `excel-com` must
continue to decide which of those primitives are safe for each Excel session
API.

## 2. Extraction classification

| Current component | Classification | Recommended owner |
| --- | --- | --- |
| STA initialization / thread affinity | Generic | `com-automation` |
| `ComPtr<T>`, AddRef/Release, QI, canonical IUnknown | Generic | `com-automation` |
| raw `IDispatch` vtable and invoke frame | Generic | `com-automation` |
| BSTR owner | Generic | `com-automation` |
| initialized owned VARIANT | Generic | `com-automation` |
| owned/borrowed SAFEARRAY primitives | Generic | `com-automation` |
| IEnumVARIANT owner | Generic | `com-automation` |
| HRESULT / EXCEPINFO capture | Generic | `com-automation` |
| IMessageFilter ABI/server plumbing | Generic mechanism | `com-automation`; policy remains above |
| Missing argument encoding | Generic Automation | `com-automation` |
| `AutomationValue` public Excel semantics | Excel-facing compatibility surface | `excel-com` initially |
| Range shape validation and Value/Value2 date behavior | Excel-specific | `excel-com` |
| Excel `MemberId` / `ObjectId` inventory | Excel-specific | `excel-com` / metadata tools |
| Excel retry-safety overrides | Excel-specific | `excel-com` |
| Excel state guards | Excel-specific | `excel-com` |
| Owned/attached Excel session policy | Excel-specific | `excel-com` |
| private process observation/shutdown | Excel-specific | `excel-com` |
| typed Excel wrappers and enums | Excel-specific | `excel-com` |
| Excel event enum/facade | Excel-specific | `excel-com`, built on generic sinks |
| typelib parsing/generation machinery | Generic concepts, developer-time | tools first; separate crate only if later justified |

A useful rule is: the lower crate should understand COM contracts, Automation
ABI types, interface identities, and call mechanics, but it should not know
what a Workbook, Range, Excel formula, or Excel process is.

## 3. Capability comparison with mature COM clients

The repository's pywin32/comtypes source audit is the primary comparison here.
.NET COM interop is used as a capability-class reference rather than as a
requirement to copy its API or runtime behavior.

| Capability | Current `excel-com` | Mature-client capability class | Roadmap |
| --- | --- | --- | --- |
| Apartment ownership | Explicit STA, thread-bound, `!Send`/`!Sync` | COM initialization plus runtime/client rules | Preserve; add MTA only deliberately |
| Generic activation | Excel ProgID only; constrained active-object attachment | CLSID/ProgID, local/remote options, active objects | Generic primitives, public dynamic API later |
| `IUnknown` / QI | Strong private owner and generic QI | General interface discovery | Extract; add typed interface metadata |
| Late-bound IDispatch | Strong Excel use; name resolved each call | Name/DISPID paths, caches, generated descriptors | Common generic invoke engine |
| Generated dispatch metadata | Large Excel inventory but runtime descriptors are still manually mapped | pywin32 makepy/comtypes generated metadata | Generate private low-level descriptors |
| Arbitrary named arguments | Property-put special case only | General named DISPIDs where server supports them | Add generic call-frame model |
| Optional/missing arguments | Supported with exact Missing marker | Explicit Missing distinct from Null/Empty | Preserve |
| VARIANT scalar coverage | Deliberately bounded semantic subset | Broad VARTYPE coverage | Lossless internal model; keep Excel facade narrow |
| VT_BYREF / OUT / INOUT | Missing | Supported by mature interop systems | Required before general events |
| SAFEARRAY | Good ownership; public Excel rectangle uses variant elements | Variant and typed arrays, arbitrary rank/bounds | Generalize core storage without changing Range API |
| IEnumVARIANT | Supported privately | Standard collection iteration | Extract |
| Structured errors | Strong HRESULT/EXCEPINFO/arg-index preservation | Structured COM exceptions | Extract generic core, adapt to `ExcelComError` |
| DISPID caching/direct descriptors | Not used by current invoke path | Dynamic caches and generated DISPIDs | Add descriptor/cache path |
| Early-bound dual/vtable interfaces | Raw QI substrate only | Generated comtypes/.NET interop can use vtables | Add optional internal generated interfaces |
| COM server/sink implementation | Only dedicated IMessageFilter precedent | Event sinks and custom COM objects | General sink substrate |
| Connection points/events | Missing | Standard event subscription | Generic connection points + curated Excel events |
| Cross-apartment marshaling | Missing by design | Runtime or explicit COM marshaling | Explicit token plus Excel STA executor |
| ROT/monikers | Excel `GetActiveObject` only | Generic ROT/moniker access | Generic lower layer; conservative Excel facade |
| COM security/DCOM | Missing | Available in mature systems | Deferred unless a real client needs it |

Two mature-client design lessons from the existing research are especially
important:

1. Generated and dynamic invocation should share one physical call engine. A
   generated descriptor should supply knowledge such as DISPID, INVOKEKIND,
   parameter types/directions/defaults, and return type; it should not create a
   second ownership/cleanup implementation.
2. Generated bindings are an implementation aid, not a reason to generate the
   whole public Excel API. The current hand-designed wrappers preserve domain
   semantics that a mechanical typelib projection cannot express well.

## 4. Missing generic features and required designs

### 4.1 Neutral member and call-frame model

Before extracting the dispatch engine, remove its dependence on Excel member
IDs. The generic layer needs an internal target concept capable of expressing:

```text
member by known DISPID
member by name
explicit default member (DISPID_VALUE)
method / property-get / property-put / property-putref
locale policy
positional arguments
named argument DISPIDs
```

Do not expose this as `excel-com`'s normal public API. Excel wrappers should
continue to call generated or curated descriptors.

A generated DISPID should avoid a redundant `GetIDsOfNames` lookup where it is
valid evidence for the supported Office type library. A name-based path should
remain available for dynamic clients and compatibility cases. If a cache is
added, key it by the live dispatch object, locale, and exact requested name set;
do not create a process-global name-to-DISPID assumption.

The current LCID behavior must not be silently changed during extraction. The
research notes that pywin32/comtypes commonly use LCID 0 while the Rust engine
currently uses different default/user/system constants. Locale normalization
belongs in a separately validated milestone.

### 4.2 BYREF, OUT, and INOUT

`VT_BYREF` must remain outside the ordinary owning value enum.

There are two distinct borrowing directions:

1. **Client call slots**: Rust owns temporary backing storage for an outbound
   BYREF/OUT/INOUT parameter for exactly the duration of `Invoke`, then decodes
   the possibly modified result.
2. **Inbound sink arguments**: COM owns the caller's storage and passes a
   borrowed pointer to the Rust event sink. Rust may read or mutate it only
   during that callback.

The design should therefore introduce lifetime-bound call/sink views rather
than add something like `AutomationValue::ByRef(*mut T)`.

Conceptually, not as a frozen public API:

```text
OwnedAutomationValue          self-contained, independently droppable
CallArgument<'call>           Value | Missing | ByRef(CallSlot<'call>)
BorrowedVariant<'call>        decode-only borrowed inbound VARIANT
BorrowedByRef<'call>          checked mutable view for allowed BYREF types
Out<T> / InOut<T>             typed client-side backing slots where useful
```

The call frame, not the caller, should build the raw pointer-bearing VARIANTs.
That lets it guarantee backing-storage stability until `Invoke` returns and
prevents a BYREF pointer from escaping.

`WorkbookBeforeClose(Workbook, Cancel)` is the first Excel acceptance test. A
successful design must expose `Cancel` as a mutable Boolean for the callback
without allowing the reference to be retained after the event returns.

### 4.3 Type-library-driven internal metadata

The repository already has the hard part: an Excel typelib inventory and
machine-readable object/event metadata. The next step is to make low-level
runtime descriptors generated from that evidence instead of maintaining every
name/invoke-kind pairing manually.

The generator should be extended to retain at least:

- coclass CLSIDs;
- interface IIDs;
- outgoing/source IIDs and coclass-interface relationships;
- member DISPIDs;
- INVOKEKIND;
- parameter order;
- parameter direction (`in`, `out`, `in,out`, retval where applicable);
- raw and normalized parameter type;
- optional/default information;
- return type;
- interface inheritance / dual-interface relationship;
- default member;
- `_NewEnum` metadata;
- enum constants;
- event metadata and connection-point source interface.

Generated Rust descriptors should be checked into the repository or produced
by an explicit developer command from checked-in metadata. A normal
`cargo build` must not require Excel to be installed or read the machine's
registered type library.

Do not generate the public Excel wrapper surface. The target is:

```text
Excel typelib/inventory evidence
          |
          v
checked-in machine-readable metadata
          |
          v
generated private low-level descriptors
          |
          v
hand-written safe Excel facade
```

The current event inventory already contains `AppEvents`/`IAppEvents` event
members and DISPIDs. Extend that pipeline rather than create a separate event
list in Rust source.

### 4.4 Generic COM sink/server substrate

Connection points require Rust to implement COM objects whose lifetime may be
controlled jointly by Rust and a COM server.

The substrate must define and test:

- one stable allocation containing the vtable pointer and state;
- atomic or otherwise COM-correct reference counting for the supported
  threading model;
- `IUnknown::QueryInterface`, `AddRef`, and `Release`;
- canonical `IUnknown` identity;
- optional `IDispatch` implementation for outgoing dispinterfaces;
- custom-vtable interface implementation where needed;
- a static supported-interface set;
- no Rust panic across any COM vtable boundary;
- translation of callback failure/panic to a controlled HRESULT;
- explicit apartment/thread assertions for callbacks;
- callback-state lifetime independent of a temporary Rust stack borrow;
- a rule for shutdown when callbacks are currently executing or reentrant.

C4 should include a focused design spike on whether to build these vtables
manually on `windows-sys` or use the audited `windows-core`/`windows`
implementation support internally. The public crate API must not depend on
raw pointers either way. The decision should favor the smaller auditable
unsafe surface, not dependency purity for its own sake.

### 4.5 Connection points and subscriptions

Generic connection-point support should model:

- `IConnectionPointContainer`;
- `FindConnectionPoint`;
- `IConnectionPoint`;
- `Advise`;
- the returned subscription cookie;
- `Unadvise`;
- RAII subscription cleanup;
- explicit restoration/unsubscribe result where failure matters.

An `Advise` success transfers no permission to forget the sink lifetime. A
subscription object must keep the sink alive until `Unadvise` completes and
must prevent use-after-free if Excel reenters the sink.

The Excel layer should then translate raw event DISPIDs and arguments into a
curated API, conceptually:

```rust
excel.events().subscribe(|event| {
    match event {
        ExcelEvent::WorkbookOpen(workbook) => { /* ... */ }
        ExcelEvent::SheetChange(sheet, range) => { /* ... */ }
        ExcelEvent::WorkbookBeforeClose { workbook, cancel } => {
            *cancel = true;
        }
        _ => {}
    }
})?;
```

The actual event enum will probably need a callback lifetime parameter because
of borrowed BYREF fields. Do not erase that lifetime to make the syntax look
simpler.

Do not implement events by polling.

### 4.6 Explicit marshaling and an Excel STA executor

Do not make `ComPtr`, Excel wrappers, or public COM objects `Send`/`Sync`.

The generic core should eventually support an explicit one-shot marshaling
mechanism around `CoMarshalInterThreadInterfaceInStream` and
`CoGetInterfaceAndReleaseStream`. Conceptually:

```text
source apartment object
       |
       v
MarshaledInterface<T>     // transferable token, not a usable COM pointer
       |
       v
consume on destination apartment
       |
       v
new apartment-bound interface owner
```

The transferable token may be `Send`; the interface wrapper obtained from it
remains apartment-bound.

A Global Interface Table abstraction can be evaluated later for repeated
cross-apartment retrieval. It should not be the first marshaling API merely
because it exists.

For Excel specifically, the preferred high-level concurrency model should be a
single dedicated STA executor, not widespread marshaling of Workbook/Range
objects. The executor owns the STA, message pump, Excel session, and all Excel
wrappers. A sendable handle submits closures/commands and returns only values
that are safe to leave the STA. Excel wrapper objects must not escape the
executor closure.

The message pump is not optional once the STA must receive COM callbacks. Event
work should therefore define the reentrancy/message-pump contract before the
executor is considered complete.

### 4.7 Generic activation, ROT, and monikers

Separate generic primitives from Excel policy.

The lower layer may ultimately support:

- activation by CLSID;
- ProgID to CLSID resolution;
- local `CoCreateInstance`/`CoCreateInstanceEx` policy;
- `GetActiveObject`;
- Running Object Table access/enumeration;
- moniker binding;
- explicit QueryInterface to a requested interface.

A later dynamic client might support a surface such as
`ComObject::create("Word.Application")`, but that API should not be the basis
of the typed Excel facade.

`OwnedApplication::new` must continue to mean a fresh private Excel local
server. The presence of generic active-object/ROT APIs below it must never make
owned Excel construction fall back to a user's running instance.

### 4.8 Early-bound and dual interfaces

The current generic `QueryInterface` primitive is sufficient groundwork, but a
mature framework also needs a safe internal way to represent known interface
IIDs and vtables.

Generated metadata may allow selected dual interfaces to use early-bound
vtable calls. That can improve signature fidelity and remove dispatch-frame
work for members where the typelib is authoritative. It should remain an
implementation option, not a requirement that all Excel calls switch away from
Automation.

Rules:

- generated typed interfaces must use the same reference owner and apartment
  invariants as IDispatch;
- a generated vtable method must have an audited ABI signature;
- the public Excel API must not expose arbitrary raw interface pointers;
- dispatch-only interfaces remain on the common IDispatch engine;
- do not duplicate conversion/error ownership code between dynamic and
  early-bound paths.

### 4.9 MTA, COM security, and remote COM

These are important to a general COM framework but not prerequisites for the
Office Automation extraction.

Potential future support includes:

- explicit `ComApartment::mta` or equivalent;
- process-global `CoInitializeSecurity` configuration;
- proxy blanket control;
- remote COM activation/DCOM.

Do not implicitly initialize COM security as a side effect of ordinary Excel
construction. `CoInitializeSecurity` is process-wide and ordering-sensitive;
if implemented it needs an explicit process-level API and tests.

Remote COM should remain deferred until a real consumer requires it. It should
not delay BYREF, events, typelibs, or the Excel executor.

## 5. Recommended crate boundary

Create `crates/com-automation/` as a new Windows-only workspace member when C1
begins. Keep the existing top-level `excel-com/` path unchanged.

Suggested internal module layout, subject to implementation evidence:

```text
crates/com-automation/src/
    apartment.rs
    interface.rs          // owned refs, QI, identity
    activation.rs
    error.rs
    bstr.rs
    variant/
        owned.rs
        borrowed.rs       // later C3/C4
        byref.rs          // later C3/C4
    safearray/
        owned.rs
        borrowed.rs
    dispatch/
        object.rs
        member.rs
        arguments.rs
        invoke.rs
    enumeration.rs
    message_filter.rs
    server/               // C4
        unknown.rs
        dispatch.rs
    connection_point.rs   // C5
    marshal.rs            // C6/C7
    rot.rs                // C7
```

Do not create a mandatory runtime dependency on the typelib inventory tool.
Keep parsing/generation in `tools/` initially. If it later becomes broadly
useful for Word/PowerPoint/third-party typelibs, a separate
`com-automation-typelib` developer/build-time crate can be extracted without
putting it on the runtime dependency path.

The generic runtime should initially keep the same dependency philosophy as
`excel-com`: Windows-only, small, and based on audited Windows bindings. A
C4 sink implementation may justify an internal `windows-core` dependency if it
materially reduces unsafe vtable code; decide that from a focused comparison.

## 6. Compatibility rules during refactoring

These are hard gates, not preferences.

1. **No public Excel API break in C0-C7.** Existing wrapper names, argument
   types, return types, errors, guards, and session semantics remain usable.
2. **Keep public type identity in `excel-com` initially.** In particular, do
   not replace `excel_com::AutomationValue`, `AutomationArgument`,
   `AutomationArray`, or `ExcelComError` with generic types merely to reduce
   duplication.
3. **Preserve `excel_com::ComApartment` source compatibility.** If the lower
   implementation moves, retain an Excel facade/wrapper that maps generic
   failures into the existing `ExcelComError` contract. A direct re-export that
   changes `ComApartment::sta`'s error type is a breaking change.
4. **Never add unsafe `Send`/`Sync`.** Cross-thread use requires explicit COM
   marshaling or the executor.
5. **No raw escape hatch for ordinary Excel users.** `IUnknown*`, `IDispatch*`,
   `VARIANT*`, `SAFEARRAY*`, and BSTR ownership remain private.
6. **Private Excel stays private.** `OwnedApplication::new` must keep fresh
   local-server semantics; generic ROT support below it cannot change that.
7. **Drop does not become Quit.** Excel application shutdown remains explicit.
8. **Preserve structured diagnostics.** HRESULT, EXCEPINFO, SCODE, argument
   index, member context, and retry information must not be lost in adapters.
9. **Do not change locale, missing-argument, or retry behavior incidentally.**
   Each behavioral change needs its own evidence and acceptance tests.
10. **Keep the PyroApp pin viable.** Existing immutable reviewed revisions stay
    valid. New `excel-com` commits must retain a stable migration path; this
    roadmap does not require changes in `pyroapprs`.
11. **No typelib dependency at end-user build time.** Generated metadata must be
    reproducible from checked-in evidence.
12. **Mocks do not prove Excel semantics.** Native deterministic fixtures are
    valuable for COM mechanics; real desktop Excel remains required for Excel
    behavior.

Because `ExcelComError` is public and callers may pattern-match it, C1 should
map lower generic errors into the same existing variants rather than introduce
a new catch-all wrapper variant solely for convenience.

## 7. Migration strategy

Use a strangler-style extraction: make `excel-com` delegate downward one
well-tested primitive at a time while its public surface remains unchanged.

### Step A - freeze behavioral baselines

Before moving code, retain or add tests that characterize:

- STA lifetime and compile-time `!Send`/`!Sync`;
- `ComPtr` clone/drop/QI/identity;
- BSTR/VARIANT/SAFEARRAY cleanup;
- missing optional arguments and reverse call order;
- property get/put/putref frame rules;
- IEnumVARIANT early-drop cleanup;
- EXCEPINFO and `puArgErr` preservation;
- current locale and retry behavior;
- the complete real-Excel migration-support suite.

### Step B - extract ownership primitives without semantic expansion

Move/copy the private apartment, COM reference, BSTR, owned VARIANT, SAFEARRAY,
and IEnumVARIANT mechanics into `com-automation`. Adapt `excel-com` internally.
Do not add BYREF or events in the same commit as the first move.

The gate is behavioral equivalence, not fewer lines of code.

### Step C - extract the generic dispatch kernel

Introduce a neutral generic member/call-frame descriptor and move the physical
`IDispatch::Invoke` engine below `excel-com`. Keep Excel descriptor lookup and
retry policy above it. Preserve the current public error structure through an
adapter.

### Step D - make generated metadata drive Excel descriptors

Extend the inventory generator and replace hand-maintained low-level descriptor
facts incrementally. Compare generated output against the existing registry and
live behavior before deleting manual facts.

### Step E - add capabilities only after the seam is stable

Implement BYREF, sink/server infrastructure, events, and marshaling in separate
milestones. These features will validate whether the extracted generic
abstractions are genuinely reusable rather than just renamed Excel internals.

### Step F - stabilize the generic public API last

Only after Excel events and marshaling work should the project decide which
`com-automation` types are public, whether the crate is publishable, and which
APIs remain internal implementation details.

## 8. Milestones

The original C0-C8 outline is directionally correct, but the dependency order
benefits from separating the dispatch seam, generated metadata, and event
substrate.

### C0 - architecture and inventory

This document is the C0 architecture deliverable.

Tasks:

- classify the current low-level implementation by generic versus Excel policy;
- record the dependency direction and compatibility rules;
- inventory gaps against pywin32/comtypes/.NET capability class;
- define validation gates;
- make no public API change.

Gate: architecture reviewed; no implementation starts until the crate boundary
and compatibility rules are accepted.

### C1 - generic ownership/ABI extraction

Create `crates/com-automation` and extract, with behavior-preserving adapters:

- apartment ownership;
- owned COM references and QI/identity;
- BSTR;
- owned VARIANT;
- owned/borrowed SAFEARRAY substrate needed by current behavior;
- IEnumVARIANT base owner;
- generic HRESULT helpers.

Do not expand VARTYPE coverage merely because the code moved.

Gate: all existing unit/compile-fail tests and serial real-Excel tests pass
unchanged through `excel-com`.

### C2 - generic dispatch engine and diagnostics

Extract the physical IDispatch engine and introduce:

- neutral member target (`DISPID`, name, default);
- invocation kind independent of Excel `MemberId`;
- explicit call-frame representation;
- general named-argument storage even if Excel wrappers do not use it yet;
- generic structured invoke error;
- adapter preserving `ExcelComError`;
- optional per-object name/DISPID cache behind tests.

Keep Excel retry-safety classification and member labels in `excel-com`.

Gate: a deterministic native dispatch fixture and current Excel tests produce
behaviorally equivalent results. No incidental LCID change.

### C3 - typelib metadata and generated internal descriptors

Extend the existing inventory pipeline to generate the low-level facts used by
Excel wrappers:

- CLSIDs/IIDs/source IIDs;
- DISPIDs/INVOKEKIND;
- parameter type/direction/optionality/defaults;
- return type;
- inheritance/dual information;
- default members and `_NewEnum`;
- enums;
- event metadata.

Migrate the hand registry incrementally. Keep hand-written public wrappers.

Gate: generated metadata is deterministic, checked against checked-in source
metadata, and all converted members pass existing live tests.

### C4 - BYREF / OUT / INOUT Automation

Implement the separate borrowed/call-slot model.

Required cases include at least:

- `VT_BYREF | VT_BOOL`;
- `VT_BYREF | VT_I4`;
- `VT_BYREF | VT_VARIANT`;
- `VT_BYREF | VT_DISPATCH`;
- safe output decoding and ownership transfer;
- inbound borrowed BYREF decoding for future sinks.

Do not add `ByRef` to the ordinary owning `AutomationValue`.

Gate: deterministic native tests prove mutation and lifetime behavior; compile
fail tests demonstrate that borrowed references cannot escape.

### C5 - generic COM sink/server implementation

Build reusable IUnknown/IDispatch implementation infrastructure, informed by
but not copied blindly from `IMessageFilter`.

Required gates:

- QI identity and static interface set;
- correct reference-count destruction;
- concurrent AddRef/Release rules appropriate to supported apartments;
- panic containment;
- callback state lifetime;
- reentrancy tests;
- no callback after final teardown.

The `windows-sys` versus `windows-core` implementation decision is resolved
here with an explicit safety/maintenance comparison.

### C6 - connection points and Excel events

Add generic connection-point support, then an Excel facade.

First Excel acceptance events:

- WorkbookOpen;
- SheetChange;
- WorkbookBeforeClose with mutable `Cancel`.

Also test:

- multiple subscriptions where Excel permits them;
- explicit Unadvise;
- drop-based unsubscribe;
- callback error/panic containment;
- Excel shutdown with active/removed subscriptions;
- source-interface metadata correctness.

Gate: real desktop Excel proves event order/arguments and cancellation behavior.
Polling is not an acceptable substitute.

### C7 - cross-apartment marshaling and Excel STA executor

Generic layer:

- one-shot interface marshaling stream token;
- destination-apartment unmarshal;
- failure/cleanup tests;
- evaluate, but do not automatically require, the Global Interface Table.

Excel layer:

- dedicated STA executor;
- message pump suitable for COM callbacks;
- sendable command handle;
- no Excel wrapper escapes the STA;
- defined reentrancy/shutdown semantics.

Gate: cross-thread tests fail to compile for raw wrappers, pass through the
explicit marshal token where appropriate, and real Excel remains responsive to
calls/events through the executor.

### C8 - generic activation, ROT, monikers, and typed-interface expansion

Add reusable lower-level support for:

- CLSID and ProgID activation;
- active object lookup;
- ROT enumeration;
- moniker binding;
- typed QI helpers;
- selected generated dual/vtable interfaces where justified.

A public dynamic Automation API may be proposed here, but it must remain
separate from the curated Excel facade.

Gate: deterministic test server coverage plus at least one non-Excel
Automation smoke test when an appropriate server is available. Excel owned
session isolation must remain unchanged.

### C9 - hardening and public stabilization

The generic crate should not claim mature-client capability before this gate.

Tasks:

- fuzz/property tests for call-frame and VARIANT conversion where practical;
- Miri/static safety review for pure-Rust ownership portions where applicable;
- Windows x64 and, when available, x86/ARM64 ABI checks;
- documentation of unsafe invariants;
- API naming/stability review;
- decide publication/versioning policy;
- measure generated-vs-dynamic invocation overhead;
- decide which MTA/security/DCOM work belongs before versus after 1.0.

Gate: all generic native tests and all supported real-Office integration suites
pass, and no raw COM representation is required by ordinary typed clients.

## 9. Validation requirements

A general COM framework needs two complementary authorities.

### 9.1 Deterministic native COM fixture

Add a test-only COM server/fixture that can deliberately exercise mechanics
that are hard to force reliably through Excel:

- QI success/failure and canonical identity;
- precise AddRef/Release lifetime counts;
- method/get/put/putref;
- default member;
- known DISPID and name resolution;
- positional plus named arguments;
- Missing versus Empty versus Null;
- exact scalar VARTYPEs;
- typed and variant SAFEARRAYs with non-zero lower bounds;
- BYREF/OUT/INOUT mutation;
- IEnumVARIANT;
- structured EXCEPINFO and `puArgErr` failures;
- connection points and callback teardown;
- callback panic containment;
- cross-apartment marshaling;
- ROT/moniker behavior when those milestones land.

The fixture proves COM mechanics. It does not prove Excel semantics.

### 9.2 Real desktop Excel

Continue running the current serial live suites and keep the completed migration
support scenario as a permanent regression gate.

New generic-capability work requires targeted Excel tests:

- generated descriptors invoke the same real members as the current path;
- no regression in private-process isolation;
- no regression in optional arguments, Formula2, CSE arrays, names, SaveAs,
  reopen, or shutdown;
- event source IID and event DISPIDs are confirmed by real callbacks;
- WorkbookBeforeClose `Cancel` mutates Excel behavior through real BYREF;
- subscriptions are removed without use-after-free or zombie Excel processes;
- the STA executor pumps callbacks and shuts down deterministically.

### 9.3 Non-Excel evidence

Once generic activation is public enough to claim reuse, exercise at least one
second Automation server. A Word/PowerPoint smoke test is useful when installed,
but CI should not depend on optional Office products. The deterministic native
fixture remains the portable Windows mechanical authority.

## 10. Important non-goals

This roadmap does not authorize:

- a Rust formula parser or calculation engine;
- an OOXML serializer to replace Excel;
- public raw COM pointers;
- unsafe `Send`/`Sync` on apartment-bound wrappers;
- events implemented by polling;
- generated public wrappers for every Excel typelib member;
- DCOM/remote activation before a concrete requirement;
- PyroApp migrator changes;
- silent attachment to a user's interactive Excel process;
- a bulk rewrite of the existing working Excel object model.

`IDispatchEx`, ActiveX hosting, OLE embedding, drag/drop, structured storage,
and arbitrary COM server registration are also outside the initial Automation
roadmap unless a real client makes them necessary.

## 11. Risks and design traps

### Accidental public-type breakage

The largest near-term refactoring risk is not COM correctness but changing
public Rust type identity/signatures while extracting internals. Keep adapters
until a deliberate versioned API decision.

### Treating BYREF as an owned value

This is the most dangerous conceptual shortcut. A BYREF `VARIANT` borrows some
other storage. If it becomes a normal clonable owning enum variant, dangling
pointers and double ownership become easy.

### Over-generating from the typelib

The typelib knows ABI shape, not all Excel semantics. `Workbooks.Add`, Range
identity, formula behavior, global/local name scope, private-process ownership,
and version quirks still need curated wrapper policy and real Excel evidence.

### Assuming generated DISPIDs remove runtime compatibility concerns

Generated metadata should reduce hand-maintained facts and repeated lookups,
but it must preserve a deliberate fallback/version policy. Do not silently
assume every installed Office build is identical to the inventory source.

### Event lifetime and reentrancy

An event sink can be called while another Automation call is active. A callback
can call Excel again. Unsubscribe/shutdown can race with callbacks. These are
lifetime and state-machine problems, not only vtable problems; design them
explicitly before exposing subscriptions.

### Hiding marshaling behind `Send`

COM marshaling is an operation that can fail and can create proxies. Encoding
it as an unsafe trait implementation would hide exactly the boundary the
current crate correctly makes explicit.

### Process-global COM security

COM security initialization is process-wide and timing-sensitive. Never add it
as an invisible side effect of constructing an Excel object.

## 12. Final architecture recommendation

Proceed with the extraction.

The current `excel-com` implementation is already past the point where its
COM plumbing is merely a small private implementation detail. It has a
substantial, reusable ownership/invocation core, a mature evidence base, and
future requirements (BYREF, sinks, connection points, marshaling, generic
activation) that are fundamentally COM concerns rather than Excel concerns.

The split should nevertheless be conservative:

```text
C0        decide and document
C1-C2     extract existing generic mechanics without behavior change
C3        make typelib evidence drive low-level descriptors
C4-C6     validate the abstraction with BYREF + sinks + real Excel events
C7        add explicit marshaling and the Excel STA executor
C8        broaden generic client activation/ROT/typed interfaces
C9        harden and only then stabilize/publish the general framework
```

This keeps the strongest property of the current project: Excel users see a
safe, curated, apartment-correct object model, while the unsafe Windows COM ABI
work is centralized and reusable. It also gives future Word, PowerPoint,
AutoCAD, engineering-software, or custom Automation clients a common native
Rust foundation without forcing the Excel API to become a dynamic string-based
COM wrapper.
