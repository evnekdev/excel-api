# COM Automation foundations for a future Excel client

**Status:** research only; provisional internal design.

**Scope:** generic Windows COM Automation client substrate, not an Excel object
model or public API.
**Date:** 2026-07-21

## Purpose and repository fit

This document records the Windows contracts that a future `excel-com` crate
would have to preserve before it can safely automate Excel or another
`IDispatch` server. It does not create that crate, change the XLL core, or
freeze an Excel-facing public API.

The repository's core 1.0 boundary deliberately excludes general COM,
Ribbon, task panes, RTD streaming, and autonomous notification. The existing
`examples/minimal-rtd-server` is an unpublished Windows-only compatibility
prototype, not an implementation precedent for a general Automation client.
Its relevant current use of `windows` 0.62.2 is limited to raw COM/OLE
interfaces, `VARIANT`, `SAFEARRAY`, apartment setup, and the Global Interface
Table. In particular, its explicit `OwnedSafeArray` demonstrates that
Automation allocation must remain separate from XLL and `xlFree` ownership.

The contracts below come primarily from Microsoft Learn Win32 documentation.
The binding-specific observations were verified against the locked
`windows`/`windows-core` 0.62.2 source in this worktree and the corresponding
Microsoft `windows-rs` sources; they do not replace the Windows contracts.

## 1. COM identity, lifetime, and apartment boundary

### `IUnknown` is the base contract

Every COM interface derives directly or indirectly from
[`IUnknown`](https://learn.microsoft.com/en-us/windows/win32/api/unknwn/nn-unknwn-iunknown).
Its three methods are the protocol for capability discovery and lifetime:

- `QueryInterface(IID, out)` asks whether this object supports an interface and,
  when successful, returns an owned interface reference. A caller must treat a
  success as one reference to release. The returned pointer must be null on a
  failed query; a defensive wrapper should also discard an unexpected non-null
  output on failure. [Microsoft's `QueryInterface` rules](https://learn.microsoft.com/en-us/windows/win32/com/rules-for-implementing-queryinterface)
  require a static interface set for an object instance and successful
  navigation between any interfaces it supports.
- `AddRef` stabilizes every new copied interface pointer. Each successful extra
  ownership claim requires exactly one later `Release`; the numerical count is
  not application logic. See [`IUnknown::AddRef`](https://learn.microsoft.com/en-us/windows/win32/api/unknwn/nf-unknwn-iunknown-addref)
  and [`IUnknown::Release`](https://learn.microsoft.com/en-us/windows/win32/api/unknwn/nf-unknwn-iunknown-release).
- `Release` may destroy the object. Code must not read, write, or call through
  that pointer after the release which drops the final reference.

An interface pointer is a capability for one interface, not necessarily a
unique object address. Two different interface pointers can refer to the same
object and have different addresses. Object identity is established by asking
each pointer for `IID_IUnknown` and comparing those canonical pointers; the
`QueryInterface` rules require the same physical `IUnknown` pointer for one
object instance. Ordinary pointer equality can establish that two *same
interface* pointers are equal, but pointer inequality does not establish
different objects.

### Ownership conventions a wrapper must make explicit

COM's C signatures do not by themselves specify a Rust lifetime. The method's
documentation and direction annotations decide it:

- an interface returned through a successful `out` parameter is normally an
  owned reference; construct an owning interface wrapper exactly once;
- an `in` interface pointer is borrowed for the call unless the contract says
  otherwise; retain it only after taking a counted reference;
- passing an owning pointer to an `in,out` slot requires care because the
  callee can release the old value before overwriting it; `AddRef` first when
  preserving an alias, as required by [`AddRef`](https://learn.microsoft.com/en-us/windows/win32/api/unknwn/nf-unknwn-iunknown-addref);
- raw ownership transfer must have one named owner at every point. A transfer
  to COM or a `VARIANT`/`SAFEARRAY` consumes the former Rust owner; it must not
  run its destructor afterwards;
- an interface pointer is apartment-affine unless COM has marshaled it. Each
  COM-using thread initializes COM, and a pointer must never simply be copied
  from one single-threaded apartment to another. Microsoft documents the
  per-thread initialization and marshaling requirement in
  [Initializing the COM Library](https://learn.microsoft.com/en-us/windows/win32/learnwin32/initializing-the-com-library)
  and [Single-Threaded Apartments](https://learn.microsoft.com/en-us/windows/win32/com/single-threaded-apartments).

`ComApartment` must pair every successful `CoInitializeEx` call with one
`CoUninitialize` on the same thread. A later design must record whether it
owns that initialization or observed compatible prior initialization; it must
not blindly uninitialize somebody else's COM use. A single-threaded apartment
also needs a message loop for incoming COM calls. The MTA allows concurrent
calls and therefore shifts synchronization responsibility to the object or
client; COM does not serialize those calls. See
[Processes, Threads, and Apartments](https://learn.microsoft.com/en-us/windows/win32/com/processes--threads--and-apartments)
and [Multithreaded Apartments](https://learn.microsoft.com/en-us/windows/win32/com/multithreaded-apartments).

### `windows` crate ownership behavior and its limits

The locked version is `windows`/`windows-core` 0.62.2. Its
[`IUnknown` implementation](https://docs.rs/crate/windows-core/0.62.2/source/src/unknown.rs)
implements `Clone` by calling `AddRef` and `Drop` by calling `Release`.
`Interface::as_raw` remains owned by the wrapper; `into_raw` consumes that
wrapper and makes the caller responsible to release the pointer; unsafe
`from_raw` assumes the caller already owns a non-null, valid pointer for the
interface. These operations are implemented in
[`Interface`](https://docs.rs/crate/windows-core/0.62.2/source/src/interface.rs).

The wrappers remove the routine `AddRef`/`Release` pairing for normal owned
interfaces. They do **not** prove that a raw pointer is valid, that an
interface has been marshaled, that an ABI signature matches, or that an
apartment/message-pump requirement has been met. `as_raw`, `into_raw`,
`from_raw`, vtable access, manually constructed COM objects, and cross-thread
transport remain unsafe or require a separately audited safe abstraction.

### Rust invariants for `DispatchObject`

1. An owned `DispatchObject` holds exactly one counted `IDispatch` reference;
   cloning it creates another counted reference.
2. The only accepted raw input is a non-null `IDispatch*` whose ownership and
   apartment provenance are recorded by the immediately enclosing unsafe
   operation.
3. A raw pointer borrowed for an FFI call never escapes that call without a
   documented `AddRef`, marshaling operation, or deep copy.
4. `QueryInterface`/`cast` failures cannot leak an output pointer, and a
   successful result is represented by an owning wrapper immediately.
5. `DispatchObject` is neither `Send` nor `Sync` by default. Crossing an
   apartment requires an explicit COM-marshaling abstraction; a plain Rust
   pointer copy is forbidden.
6. No panic may cross a COM vtable or FFI boundary. A client boundary must
   preserve the returned `HRESULT`; a future server boundary must map panics to
   a controlled failure `HRESULT`.

## 2. Automation dispatch

### `IDispatch` and type information

[`IDispatch`](https://learn.microsoft.com/en-us/windows/win32/api/oaidl/nn-oaidl-idispatch)
is the Automation interface that exposes methods and properties for late-bound
clients. It inherits `IUnknown` and has four additional methods:

- `GetTypeInfoCount` reports either zero or one type-information interface. A
  zero result does **not** make the object undispatchable; it only means runtime
  type information is unavailable. See
  [`GetTypeInfoCount`](https://learn.microsoft.com/en-us/windows/win32/api/oaidl/nf-oaidl-idispatch-gettypeinfocount).
- `GetTypeInfo(0, lcid, out)` obtains the optional `ITypeInfo`; index zero is
  the only valid type-information index. `lcid` can select localized type
  information. See
  [`GetTypeInfo`](https://learn.microsoft.com/en-us/windows/win32/api/oaidl/nf-oaidl-idispatch-gettypeinfo).
- `GetIDsOfNames(IID_NULL, names, lcid, out_dispids)` maps the member name and
  optional parameter names to DISPIDs. The first name is the member and later
  names are parameter names. DISPIDs remain constant for the object's
  lifetime, so a cache may be scoped to one live `DispatchObject`, its locale,
  and the exact requested name set. A failed lookup can return
  `DISP_E_UNKNOWNNAME` and `DISPID_UNKNOWN` for affected entries. See
  [`GetIDsOfNames`](https://learn.microsoft.com/en-us/windows/win32/api/oaidl/nf-oaidl-idispatch-getidsofnames).
- `Invoke` calls a member by DISPID for a method, property get, property put,
  or property put-reference. Its complete call contract is documented by
  [`IDispatch::Invoke`](https://learn.microsoft.com/en-us/windows/win32/api/oaidl/nf-oaidl-idispatch-invoke).

`DISPID_VALUE` is zero, the reserved default or `Value` member DISPID. A
generic client should model "default member" as an explicit invoke target
(`DISPID_VALUE`) rather than guessing a name. Other negative DISPIDs are
special-purpose values; do not treat them as ordinary user member IDs.

### Invocation kinds

`Invoke` uses one or more `DISPATCH_*` flags:

| Kind | Flag | Meaning |
| --- | --- | --- |
| Method | `DISPATCH_METHOD` | Invoke the member as a method. |
| Property get | `DISPATCH_PROPERTYGET` | Retrieve a property or data member. |
| Property put | `DISPATCH_PROPERTYPUT` | Assign a value to a property or data member. |
| Property put-reference | `DISPATCH_PROPERTYPUTREF` | Assign an object reference to a property that accepts one. |

The method and property-get flags may be combined when a property and method
share the name. A put-reference is not a synonym for a normal put: it records
reference assignment semantics and is valid only for a property that accepts
an object reference. For a put or put-reference, the result pointer is
ignored.

### DISPIDs, arguments, locale, and missing values

`DISPPARAMS` holds the argument `VARIANT`s, optional named-argument DISPIDs,
and their counts. Microsoft defines its fields and the reverse-order rule in
[`DISPPARAMS`](https://learn.microsoft.com/en-us/windows/win32/api/oaidl/ns-oaidl-dispparams):

- `rgvarg` is in reverse parameter order. For positional arguments written as
  `(first, second, third)`, build `rgvarg` as `(third, second, first)`.
- `rgdispidNamedArgs` contains one DISPID per named argument, and the first
  `cNamedArgs` entries of `rgvarg` are its matching values. Named arguments
  may be supplied in any order as long as each value remains paired with its
  DISPID; positional values follow and remain reversed. This fuller layout is
  specified by the official [Automation protocol's dispatch consistency
  checks](https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-oaut/31e8ad0e-1d5f-4b60-9403-0082eb1a5080).
- A property put or put-reference must set `cNamedArgs = 1` and pass
  `DISPID_PROPERTYPUT` in `rgdispidNamedArgs[0]`. Its value is the first
  `rgvarg` entry. The required special initialization is stated in
  [`IDispatch::Invoke`](https://learn.microsoft.com/en-us/windows/win32/api/oaidl/nf-oaidl-idispatch-invoke).
- A positional omitted optional parameter is the missing-argument marker:
  `VARIANT { vt: VT_ERROR, scode: DISP_E_PARAMNOTFOUND }`. The OLE Automation
  protocol specifies this representation in
  [Handling Default Value and Optional Arguments](https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-oaut/be6e35f6-9327-4164-9bde-ffcd0fa0e07d).
  `VT_EMPTY`, `VT_NULL`, and a missing marker are distinct values.
- Pass the same intended locale to name lookup and `Invoke`. `lcid` gives the
  server the locale context for names and argument interpretation; servers may
  return `DISP_E_UNKNOWNLCID`. A future API should require an explicit
  `LCID` policy rather than silently assuming a process-wide locale.

The `riid` argument of both `GetIDsOfNames` and `Invoke` is reserved and must
be `IID_NULL`. `Invoke` can report an argument error only for
`DISP_E_TYPEMISMATCH` or `DISP_E_PARAMNOTFOUND`; `argErr` indexes the physical
`rgvarg` array, so it is also reverse-oriented. Preserve it without converting
it to a guessed source-language parameter index.

### Correct generic invocation sequence

The following pseudocode specifies the ordering and ownership boundary. It is
an internal algorithm, not a public API proposal.

```rust
fn invoke(
    object: &DispatchObject,
    member: MemberId,                 // known DISPID or a name to resolve
    kind: InvokeKind,
    locale: Lcid,
    args: InvokeArgs,                 // positional in source order; named pairs
) -> Result<AutomationValue, AutomationError> {
    let dispid = match member {
        MemberId::Dispid(id) => id,
        MemberId::Name(name) => object.get_ids_of_names(IID_NULL, [name], locale)?[0],
        MemberId::Default => DISPID_VALUE,
    };

    // Each item below is an owned, initialized VARIANT. It stays alive until
    // Invoke returns; its Drop/VariantClear path executes on every result.
    let mut rgvarg = Vec::<VARIANT>::new();
    let mut named_dispids = Vec::<DISPID>::new();

    match kind {
        InvokeKind::PropertyPut { value, by_reference } => {
            if !args.named.is_empty() {
                return Err(AutomationError::unsupported_named_property_put());
            }
            rgvarg.push(value.into_variant()?); // rgvarg[0], the property value
            named_dispids.push(DISPID_PROPERTYPUT);
            // Indexed-property arguments follow the value in reverse logical
            // order: for Prop[left, right] = value, use value, right, left.
            for index in args.positional.iter().rev() {
                rgvarg.push(index.into_variant()?);
            }
            let flags = if by_reference {
                DISPATCH_PROPERTYPUTREF
            } else {
                DISPATCH_PROPERTYPUT
            };
            return invoke_raw(object, dispid, locale, flags, rgvarg, named_dispids);
        }
        InvokeKind::Method | InvokeKind::PropertyGet => {
            // Named entries occupy the first cNamedArgs slots and each stays
            // paired with the matching name DISPID.
            for NamedArg { dispid, value } in args.named_in_dispatch_order() {
                rgvarg.push(value.into_variant()?);
                named_dispids.push(dispid);
            }
            for value in args.positional.iter().rev() {
                rgvarg.push(value.into_variant()?);
            }
            let flags = match kind {
                InvokeKind::Method => DISPATCH_METHOD,
                InvokeKind::PropertyGet => DISPATCH_PROPERTYGET,
                _ => unreachable!(),
            };
            return invoke_raw(object, dispid, locale, flags, rgvarg, named_dispids);
        }
    }
}

fn invoke_raw(/* ... owned arrays ... */) -> Result<AutomationValue, AutomationError> {
    let mut result = VARIANT::default();       // initialized `VT_EMPTY`
    let mut exception = EXCEPINFO::default();  // all fields null/zero
    let mut arg_err = 0_u32;
    let hr = unsafe {
        dispatch.Invoke(
            dispid, IID_NULL, locale, flags, &mut dispparams,
            if flags.is_put() { None } else { Some(&mut result) },
            Some(&mut exception), Some(&mut arg_err),
        )
    };

    if hr == DISP_E_EXCEPTION {
        // Run any documented deferred fill-in before copying the fields; then
        // copy BSTR fields into owned Rust strings and free their OLE storage.
        return Err(AutomationError::from_exception(hr, exception));
    }
    if hr.is_err() {
        return Err(AutomationError::from_invoke(hr, arg_err));
    }
    AutomationValue::try_from_owned_variant(result)
}
```

The real implementation must keep `rgvarg`, `rgdispidNamedArgs`,
`DISPPARAMS`, the result, and exception storage alive through the FFI call;
must use non-null array pointers only when their count is non-zero; and must
avoid invalidating vector storage after handing pointers to `DISPPARAMS`.
`EXCEPINFO::default()` must be checked against the generated binding layout
when implementation begins.

`EXCEPINFO` describes an exception from `Invoke`: source, description, help
file/context, and either `wCode` or `scode`. Its `pfnDeferredFillIn` may defer
the textual fields until the caller requests them. A same-process client must
honor the callback before copying the fields; for a cross-process call the
server invokes it before returning. These rules come from
[`EXCEPINFO`](https://learn.microsoft.com/en-us/windows/win32/api/oaidl/ns-oaidl-excepinfo).

## 3. Automation values and cleanup

### `VARIANT` is a tagged union, not a Rust enum

[`VARIANT`](https://learn.microsoft.com/en-us/windows/win32/api/oaidl/ns-oaidl-variant)
contains a `VARTYPE` tag (`vt`) and a union. Read or write only the union field
selected by the base type and modifier bits in `vt`. The `VARENUM` definitions
list the scalar tags and the `VT_ARRAY`/`VT_BYREF` modifiers in
[`VARENUM`](https://learn.microsoft.com/en-us/windows/win32/api/wtypes/ne-wtypes-varenum).

The generic boundary must distinguish at least:

| `VARTYPE` form | Meaning | Proposed internal representation |
| --- | --- | --- |
| `VT_EMPTY` | no value | `AutomationValue::Empty` |
| `VT_NULL` | Automation null | `AutomationValue::Null` |
| `VT_I1`/`I2`/`I4`/`I8`, `VT_UI1`/`UI2`/`UI4`/`UI8`, `VT_INT`/`UINT` | fixed-width scalar integer | exact signed/unsigned enum variants; never silently truncate |
| `VT_R4`, `VT_R8` | IEEE floats | `F32`, `F64` |
| `VT_BOOL` | 16-bit Automation Boolean | `Bool`; accept only `0` false and `-1` true |
| `VT_BSTR` | owned length-prefixed UTF-16 string | `String` or lossless `Vec<u16>` policy, preserving embedded NULs |
| `VT_DATE` | OLE Automation date/time `f64` | `OaDate(f64)`, not an unqualified Rust timestamp |
| `VT_CY` | fixed-point currency | `Currency { scaled_i64 }` |
| `VT_ERROR` | `SCODE`/error value | `Error(HRESULT)`; includes the optional-argument marker |
| `VT_DISPATCH` | `IDispatch*` | apartment-bound `DispatchObject` |
| `VT_UNKNOWN` | `IUnknown*` | apartment-bound `UnknownObject` or unsupported value |
| `VT_ARRAY | base` | `SAFEARRAY*` | `AutomationArray` retaining element type and bounds |
| `VT_BYREF | base` | borrowed pointer to caller-owned storage | separate borrowed only type, never the ordinary owning enum |

The table is deliberately a *lossless internal model*, rather than a promise
that every form becomes an ergonomic public Rust type. `VT_EMPTY`, `VT_NULL`,
and `VT_ERROR` are semantically different. `VT_ERROR` is both an error-valued
Automation value and the mechanism used for an omitted positional optional
argument; it must not be collapsed into a Rust `None`.

`VARIANT_BOOL` is not Rust `bool`: it is a 16-bit value where `0xFFFF` (`-1`)
is true and `0` is false; other values are invalid. See the `VARIANT` field
contract and [VARIANT_BOOL protocol definition](https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-oaut/7b39eb24-9d39-498a-bcd8-75c38e5823d0).

An OLE Automation `DATE` is a floating number of days, with a fractional
day for time, from 30 December 1899. It is a local-time representation with
well-known pre-epoch discontinuities; do not silently convert it to a UTC
instant. See [SystemTimeToVariantTime](https://learn.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-systemtimetovarianttime)
and [DATE Type](https://learn.microsoft.com/en-us/cpp/atl-mfc-shared/date-type?view=msvc-170).
`CY` is a signed 64-bit two's-complement integer scaled by 10,000, so retain
its scale exactly rather than converting through `f64`. See
[`CY`](https://learn.microsoft.com/en-us/windows/win32/api/wtypes/ns-wtypes-cy-r1).

### `BSTR`, interfaces, arrays, and `VariantClear`

`BSTR` is a length-prefixed wide string allocated by OLE Automation. It can
contain embedded NUL code units; C-string scanning is incorrect. A string from
`SysAllocStringLen` must be released with `SysFreeString`, as documented by
[`SysAllocStringLen`](https://learn.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-sysallocstringlen)
and [`SysFreeString`](https://learn.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-sysfreestring).

An initialized owned `VARIANT` must be cleared exactly once before its storage
is reused or freed. [`VariantClear`](https://learn.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-variantclear)
releases a `VT_BSTR`, releases a `VT_DISPATCH`, destroys an owned array, and
sets the tag to `VT_EMPTY`. It must only be called on initialized storage. It
does **not** release an interface behind `VT_DISPATCH | VT_BYREF` or
`VT_UNKNOWN | VT_BYREF`; that borrowed/reference-owned interface remains the
caller's responsibility. It can fail with `DISP_E_ARRAYISLOCKED`, so an owned
variant must never be dropped while one of its arrays is data-locked.

### `windows::Win32::System::Variant::VARIANT`

The inspected `windows` 0.62.2 extension implementation gives `VARIANT` the
following behavior:

- `Default` is zeroed and therefore initializes the normal union arm as
  `VT_EMPTY`;
- `Clone` calls `VariantCopy`;
- `Drop` calls `VariantClear` and discards its returned `HRESULT`;
- conversions construct selected supported scalar, `BSTR`, `IUnknown`, and
  `IDispatch` forms, while union fields use `ManuallyDrop` to prevent an
  unselected member from being dropped.

See the version-locked
[`VARIANT` extension source](https://docs.rs/crate/windows/0.62.2/source/src/extensions/Win32/System/Variant.rs)
and the `BSTR` owner in
[`windows-strings`](https://docs.rs/crate/windows-strings/0.5.1/source/src/bstr.rs).

This makes a local `VARIANT` convenient for the simple owned cases, but it is
not a complete safe Automation value model. In particular, the raw union,
`VT_BYREF`, raw `SAFEARRAY*`, raw ownership transfer, and a failed clear due to
a locked array remain the wrapper's responsibility. A future internal owner
must make it impossible to hold a data-lock guard across its owning variant's
drop.

## 4. SAFEARRAYs

[`SAFEARRAY`](https://learn.microsoft.com/en-us/windows/win32/api/oaidl/ns-oaidl-safearray)
is a descriptor for an Automation array. It records the rank (`cDims`),
features/element-category flags, byte size per element, lock count, data
pointer, and one bound per dimension. The descriptor is not proof of Rust
ownership: method direction and the enclosing `VARIANT` determine whether the
callee borrows it or whether the caller owns and must destroy it.

### Creation and ownership

`SafeArrayCreate(base_vartype, dimensions, bounds)` allocates and initializes
both a descriptor and data. Its base type cannot include `VT_ARRAY` or
`VT_BYREF`, and `VT_EMPTY`/`VT_NULL` are invalid base types. Rank cannot change
after creation. See [`SafeArrayCreate`](https://learn.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-safearraycreate).

A freshly created `SAFEARRAY*` needs one explicit owner. That owner calls
`SafeArrayDestroy` on every non-transferred path. Destruction releases the
descriptor and data, calls `VariantClear` for each `VT_VARIANT` element,
`SysFreeString` for each `BSTR`, and `Release` for contained objects. It fails
while the array is locked. See
[`SafeArrayDestroy`](https://learn.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-safearraydestroy).

Therefore `AutomationArray` must distinguish:

- `BorrowedSafeArray<'call>`: a non-null input pointer only valid for the
  documented call or owner lifetime; it never destroys the pointer;
- `OwnedSafeArray`: a freshly created or explicitly returned-and-transferred
  descriptor/data pair; `Drop` calls `SafeArrayDestroy` once after all locks
  have ended;
- `AutomationValue::Array(OwnedSafeArray)`: ownership moves into the
  `VARIANT`/value owner, so no second array destructor runs.

This is intentionally not a public API commitment. It makes raw transfer
auditable and prevents treating every `SAFEARRAY*` as an owned allocation.

### Bounds, element types, and access

Each dimension has its own lower and upper bound. Lower bounds need not be
zero, so a Rust `0..len` range is invalid unless the native bounds have been
translated and retained. Query dimensions and bounds with the automation APIs,
including [`SafeArrayGetLBound`](https://learn.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-safearraygetlbound)
and `SafeArrayGetUBound`; calculate counts with checked arithmetic.

The element type and cleanup behavior are part of the descriptor's feature
flags. A `VT_VARIANT` array has initialized `VARIANT` elements and receives
`VariantClear` element cleanup. A primitive array has fixed-size primitive
elements and must be decoded only as its declared base type. Never cast an
unknown array data pointer to `VARIANT*` merely because the surrounding value
is Automation data.

Use `SafeArrayGetElement`/`SafeArrayPutElement` for conservative per-element
copying. They automatically lock and unlock, and correctly copy strings,
objects, and variants; the destination storage still must be initialized and
cleared appropriately. See
[`SafeArrayGetElement`](https://learn.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-safearraygetelement).

For direct bulk access, `SafeArrayAccessData` increments the lock count and
returns a data pointer. Every successful call must pair with
`SafeArrayUnaccessData`, even on conversion failure or panic. An RAII guard is
required because `SafeArrayDestroy` and `VariantClear` can fail while the
array is locked. See
[`SafeArrayAccessData`](https://learn.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-safearrayaccessdata).
An instance is not generally thread-safe; the documented lock count is not a
replacement for synchronization.

### Multidimensional indexing

The shape has two related orders that a Rust abstraction must not confuse:

- in the descriptor, `rgsabound[0]` is the left-most dimension and
  `rgsabound[cDims - 1]` is the right-most dimension;
- for `SafeArrayGetElement` and `SafeArrayPutElement`, `rgIndices[0]` is the
  right-most (least significant) dimension and the left-most index is at
  `rgIndices[cDims - 1]`.

The latter rule is stated by
[`SafeArrayGetElement`](https://learn.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-safearraygetelement).
`AutomationArray` should expose dimensions and logical indices in one
documented Rust order, then perform the reversal exactly once at the FFI
boundary. It must preserve each native lower bound. This document makes no
Excel-specific claim about the orientation or bounds that Excel returns.

## 5. HRESULTs and Automation errors

An `HRESULT` is a signed status value: non-negative is success and negative is
failure, per [`SUCCEEDED`](https://learn.microsoft.com/en-us/windows/win32/api/winerror/nf-winerror-succeeded)
and [`FAILED`](https://learn.microsoft.com/en-us/windows/win32/api/winerror/nf-winerror-failed).
The high bit is severity and the facility identifies the producer; the
dispatch facility is specifically used for late-bound `IDispatch` errors. See
[Structure of COM Error Codes](https://learn.microsoft.com/en-us/windows/win32/com/structure-of-com-error-codes).

For `Invoke`, preserve at least these failure cases exactly:

| Result | Meaning for a generic client | Required preserved detail |
| --- | --- | --- |
| `DISP_E_MEMBERNOTFOUND` | requested member does not exist, or invocation kind is incompatible | member name/DISPID and invoke kind |
| `DISP_E_UNKNOWNNAME` | name lookup failed | requested names and locale |
| `DISP_E_BADPARAMCOUNT` | call-frame argument count differs from member contract | argument count and member |
| `DISP_E_TYPEMISMATCH` | an argument could not be coerced | raw `argErr` index and original call-frame order |
| `DISP_E_PARAMNOTFOUND` | parameter DISPID mismatch; also the missing-argument `SCODE` marker | raw `argErr` when supplied |
| `DISP_E_PARAMNOTOPTIONAL` | a required parameter was omitted | member and supplied call frame |
| `DISP_E_NONAMEDARGS` | server rejects named arguments | named DISPIDs and values' type tags |
| `DISP_E_UNKNOWNLCID` | server cannot interpret the locale | requested `LCID` |
| `DISP_E_EXCEPTION` | server supplied structured exception information | complete copied `EXCEPINFO`, including source/description/help/code |
| another failed `HRESULT` | activation, marshaling, interface, or server failure | exact `HRESULT`, operation, and context |

The `Invoke` documentation lists those dispatch errors and defines `argErr`'s
reverse-order index semantics. Do not reduce them to a formatted string. The
base error should retain the 32-bit `HRESULT`, a stable operation (`activation`,
`query_interface`, `name_lookup`, `invoke`, `array`, `conversion`), and
structured supplemental information. A display message can be derived later.

`AutomationError` should be provisionally shaped as:

```rust
struct AutomationError {
    hresult: HRESULT,
    operation: AutomationOperation,
    member: Option<MemberContext>,
    lcid: Option<LCID>,
    arg_err_rgvarg_index: Option<u32>,
    exception: Option<AutomationException>,
}

struct AutomationException {
    code: Option<u16>,
    scode: Option<HRESULT>,
    source: Option<String>,
    description: Option<String>,
    help_file: Option<String>,
    help_context: u32,
}
```

The raw result and every usable `EXCEPINFO` field must be copied before its OLE
allocation is cleaned up. `AutomationException` must retain either `wCode` or
`scode` rather than inventing a replacement error code.

## 6. Provisional internal boundary

This is a non-binding internal model. It draws the raw boundary where Windows
ownership, apartment, and tagged-union requirements can be audited without
committing to names or ergonomic behavior for an Excel public API.

| Internal type | Responsibility | Raw details hidden |
| --- | --- | --- |
| `ComApartment` | Initialize/uninitialize COM on one thread, record apartment model and initialization ownership, and provide the precondition for COM work. | `CoInitializeEx` flags/results, balancing `CoUninitialize`, message-loop requirement, and apartment provenance. |
| `DispatchObject` | Own an `IDispatch` reference in its valid apartment; resolve names and perform the raw call. | `IUnknown` reference accounting, `QueryInterface`, raw vtable pointer, IID nullability, and raw `IDispatch*`. |
| `AutomationValue` | Represent a self-contained Automation value suitable for a call frame or return value. | `VARIANT` union fields, `VARTYPE` modifiers, `BSTR` allocation, contained COM references, `VariantClear`, and lossless special values. |
| `AutomationArray` | Represent borrowed or owned SAFEARRAY data with checked shape, base type, bounds, and cleanup. | descriptor/data pointers, feature flags, element-copy mechanics, lock counts, index reversal, and `SafeArrayDestroy`. |
| `InvokeKind` | State whether a call is a method, property get, property put, or property put-reference. | `DISPATCH_*` flags and property-put special case. |
| `InvokeArgs` | Own initialized call arguments and encode the exact positional/named/missing layout for `DISPPARAMS`. | reverse `rgvarg`, named-DISPID pairing, `DISPID_PROPERTYPUT`, counts, and temporary array lifetimes. |
| `AutomationError` | Preserve an operation failure without losing COM/Automation diagnostics. | `HRESULT` bit pattern, `DISP_E_*`, raw `argErr`, `EXCEPINFO`, deferred fill-in, and OLE string cleanup. |

No type above is proposed as public yet. In particular, neither `Send`/`Sync`
behavior, Excel method naming, `Range` conversion policy, property APIs,
implicit coercion, date conversion, nor an error-display contract is frozen.

## 7. Required source audit

| Topic | Official source | Contract extracted | Open question |
| --- | --- | --- | --- |
| COM identity | [QueryInterface rules](https://learn.microsoft.com/en-us/windows/win32/com/rules-for-implementing-queryinterface) | `IID_IUnknown` is the canonical object-identity comparison; supported interfaces are static. | Which identity caching, if any, is valuable for a future Excel client? |
| Reference ownership | [`IUnknown::AddRef`](https://learn.microsoft.com/en-us/windows/win32/api/unknwn/nf-unknwn-iunknown-addref) and [`Release`](https://learn.microsoft.com/en-us/windows/win32/api/unknwn/nf-unknwn-iunknown-release) | Every copied interface pointer needs a matching release; an `in,out` pointer may be released by the callee. | Which Excel methods have unusual transfer annotations? |
| COM initialization | [Initializing the COM Library](https://learn.microsoft.com/en-us/windows/win32/learnwin32/initializing-the-com-library) | Initialize COM per using thread; pair successful calls with `CoUninitialize`; marshal rather than copy STA pointers. | Which apartment should create and drive Excel Automation in supported deployment modes? |
| Apartments | [Processes, Threads, and Apartments](https://learn.microsoft.com/en-us/windows/win32/com/processes--threads--and-apartments) and [STAs](https://learn.microsoft.com/en-us/windows/win32/com/single-threaded-apartments) | Objects live in apartments; STA cross-apartment calls use marshaling/message dispatch; MTA calls require synchronization. | What thread/message-pump integration is required beside an XLL host? |
| Dispatch interface | [`IDispatch`](https://learn.microsoft.com/en-us/windows/win32/api/oaidl/nn-oaidl-idispatch) | Automation exposes members through name-to-DISPID lookup and `Invoke`; type info is optional. | Which Excel interfaces expose useful `ITypeInfo` at runtime? |
| Name lookup | [`GetIDsOfNames`](https://learn.microsoft.com/en-us/windows/win32/api/oaidl/nf-oaidl-idispatch-getidsofnames) | Member and parameter names map to stable-for-object-lifetime DISPIDs under an LCID. | Are Excel DISPIDs safe to cache beyond one application/session? |
| Invocation flags/results | [`Invoke`](https://learn.microsoft.com/en-us/windows/win32/api/oaidl/nf-oaidl-idispatch-invoke) | Defines method/get/put/putref, IID_NULL, LCID, result, exceptions, and dispatch errors. | Which Excel members require combined method/get flags or put-reference? |
| Call-frame layout | [`DISPPARAMS`](https://learn.microsoft.com/en-us/windows/win32/api/oaidl/ns-oaidl-dispparams) and [dispatch consistency checks](https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-oaut/31e8ad0e-1d5f-4b60-9403-0082eb1a5080) | `rgvarg` is reversed; named values lead and pair with named DISPIDs; positional missing values use `VT_ERROR/DISP_E_PARAMNOTFOUND`. | Do Excel indexed properties expose any edge cases needing a concrete experiment? |
| Exception diagnostics | [`EXCEPINFO`](https://learn.microsoft.com/en-us/windows/win32/api/oaidl/ns-oaidl-excepinfo) | `DISP_E_EXCEPTION` can carry source, description, help, code, and deferred fill-in. | Which Excel errors reliably populate which fields? |
| Variant tag/union | [`VARIANT`](https://learn.microsoft.com/en-us/windows/win32/api/oaidl/ns-oaidl-variant) and [`VARENUM`](https://learn.microsoft.com/en-us/windows/win32/api/wtypes/ne-wtypes-varenum) | The tag selects the union and `VT_ARRAY`/`VT_BYREF` modify base types. | Which variant forms does Excel return for each planned object-model operation? |
| Variant cleanup | [`VariantClear`](https://learn.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-variantclear) | Clear initialized owned variants; it releases BSTR/dispatch/array contents but not by-reference interfaces. | What `VT_BYREF` forms occur in Excel out/ref arguments? |
| BSTR | [`SysAllocStringLen`](https://learn.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-sysallocstringlen) and [`SysFreeString`](https://learn.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-sysfreestring) | BSTR is length-based, may contain embedded NUL, and has OLE allocator ownership. | Should a public API preserve invalid UTF-16 losslessly? |
| Date/currency/bool | [`SystemTimeToVariantTime`](https://learn.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-systemtimetovarianttime), [`CY`](https://learn.microsoft.com/en-us/windows/win32/api/wtypes/ns-wtypes-cy-r1), [`VARIANT_BOOL`](https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-oaut/7b39eb24-9d39-498a-bcd8-75c38e5823d0) | DATE is a local floating day count, CY is scale-10,000 fixed point, and true is -1. | How should a public Excel API express date systems and currency without loss? |
| Array descriptor | [`SAFEARRAY`](https://learn.microsoft.com/en-us/windows/win32/api/oaidl/ns-oaidl-safearray) | Rank, bounds, data, lock count, features, and element category drive access and release. | What ranks/bounds/feature flags does Excel use for each result shape? |
| Array lifetime | [`SafeArrayCreate`](https://learn.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-safearraycreate) and [`SafeArrayDestroy`](https://learn.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-safearraydestroy) | Fresh owner destroys descriptor/data unless transferred; destruction clears complex elements and fails locked. | Which Excel methods transfer array ownership to the caller? |
| Array locking/access | [`SafeArrayAccessData`](https://learn.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-safearrayaccessdata) and [`SafeArrayGetElement`](https://learn.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-safearraygetelement) | Direct access locks/unlocks; element helpers copy values and manage their temporary locks. | Is bulk access worth the complexity for expected Excel transfers? |
| Multidimensional indexing | [`SafeArrayGetElement`](https://learn.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-safearraygetelement) | Index vector is right-most dimension first, unlike descriptor-bound order. | Which logical order should a Rust array view use for Excel matrices? |
| HRESULT semantics | [`SUCCEEDED`](https://learn.microsoft.com/en-us/windows/win32/api/winerror/nf-winerror-succeeded), [`FAILED`](https://learn.microsoft.com/en-us/windows/win32/api/winerror/nf-winerror-failed), and [COM error structure](https://learn.microsoft.com/en-us/windows/win32/com/structure-of-com-error-codes) | Success/failure is sign-based; facility identifies dispatch-specific errors. | Which HRESULTs need typed public classifications versus raw preservation only? |

Binding audit, separate from the Windows authority above: the repository locks
`windows` and `windows-core` 0.62.2. The version-locked source links for
[`IUnknown` ownership](https://docs.rs/crate/windows-core/0.62.2/source/src/unknown.rs),
[`Interface` raw transfer](https://docs.rs/crate/windows-core/0.62.2/source/src/interface.rs),
and [`VARIANT` cleanup](https://docs.rs/crate/windows/0.62.2/source/src/extensions/Win32/System/Variant.rs)
were inspected to describe binding behavior, not to infer an undocumented
Windows contract.

## 8. Open Excel-specific questions

1. Does the supported Excel Application object expose a stable Automation
   activation route and apartment/message-pump model suitable for an add-in
   process, and how does it behave during Excel shutdown?
2. Which interfaces are dual/dispatch-capable in the installed Excel type
   library, and what are their exact IIDs, DISPIDs, optional parameters, and
   `propget`/`propput`/`propputref` annotations?
3. Which DISPIDs and name mappings remain stable across supported Excel
   versions, locales, and application sessions? What cache scope is sound?
4. For each intended Excel property and method, which `LCID` is accepted and
   when do localized names or localized string/numeric coercions matter?
5. Which `VARIANT` scalar tags and error forms does Excel return for
   application, workbook, worksheet, and range operations? Are any values
   `VT_BYREF` or otherwise borrowed past the immediate call?
6. For range value transport, what SAFEARRAY rank, lower bounds, element type,
   and logical row/column order does Excel return and accept in every supported
   scenario? This requires type-library inspection plus real-Excel experiments.
7. How does Excel distinguish an omitted optional value
   (`VT_ERROR/DISP_E_PARAMNOTFOUND`) from `VT_EMPTY`, `VT_NULL`, and an Excel
   error cell for the APIs this crate might eventually expose?
8. Which Excel server failures provide `DISP_E_EXCEPTION`/`EXCEPINFO`, and are
   deferred exception fields observable for in-process Automation calls?
9. Are `VARIANT_BOOL`, OLE Automation date, and `CY` values accepted and
   returned with Excel-specific conversions or date-system behavior that a
   public API must model explicitly?
10. What lifecycle, reference, and marshaling behavior occurs when an Excel
    object reference is retained while workbooks close, the application quits,
    or callbacks re-enter client code?
11. Can a general Automation client coexist with the XLL runtime without
    granting any additional Excel C API callback capability or changing the
    core publication/dependency boundary?
12. Which real-Excel validations are required for activation, locale,
    optional-argument, property-put, SAFEARRAY, cleanup, and shutdown paths
    before any public `excel-com` API is proposed?
