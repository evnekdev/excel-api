# Codex Prompt 08 — Implement the Dynamic `IDispatch` Kernel

## Objective

Implement the generic COM Automation invocation layer required by all Excel-specific wrappers.

Do not add broad Excel object-model wrappers beyond the minimum fixture needed to validate dispatch.

## Repository and branch

Repository: `evnekdev/excel-api`

Branch:

```text
feature/excel-com-08-dispatch-kernel
```

## Required reading

Read the COM foundations research, Excel COM architecture, dynamic-dispatch ADR, apartment/lifecycle ADR, current skeleton, and official Microsoft documentation for `GetIDsOfNames` and `Invoke`.

## Required implementation

### `DispatchObject`

Implement a safe owning wrapper around an apartment-valid `IDispatch`.

It must:

- preserve COM reference ownership;
- remain `!Send` and `!Sync`;
- retain or reference the apartment token according to architecture;
- expose raw access only through documented advanced API;
- never allow borrowed raw pointers to outlive the wrapper.

### Member resolution

Implement name-to-DISPID resolution, UTF-16/BSTR conversion, locale handling, structured member-not-found errors, and safe DISPID caching.

Cache design must account for object/interface identity, member name, locale, and apartment confinement.

### Invocation

Implement internal invocation kinds:

```rust
Method
PropertyGet
PropertyPut
PropertyPutRef
```

Correctly construct reversed positional arguments, named arguments, `DISPID_PROPERTYPUT`, optional/missing values, `DISPPARAMS`, mutable argument storage, result `VARIANT`, `EXCEPINFO`, and `argErr`.

No panic may cross a COM boundary.

### Values

Implement the minimum approved `AutomationValue` variants:

- empty;
- null;
- missing;
- Boolean;
- signed integer;
- floating point;
- string;
- error code;
- dispatch object;
- unsupported/raw variant reporting if approved.

Do not implement rectangular SAFEARRAY transport yet unless needed for a scalar test.

### Errors

Preserve HRESULT, member name, invocation kind, argument index, exception source, description, and help context where available. Do not collapse errors into strings.

## Test COM server

Create a small test Automation object exposing enough `IDispatch` behavior to test method calls, property get/put, optional and named arguments, argument order, type mismatch, unknown member, server exception, and returned dispatch object.

Do not rely on Excel for deterministic tests.

## Safety documentation

Every unsafe block must state pointer validity, ownership, lifetime, argument-buffer stability, cleanup responsibility, and apartment assumption.

## Public API

Keep the public dynamic API narrow, for example:

```rust
DispatchObject::get(...)
DispatchObject::set(...)
DispatchObject::call(...)
DispatchObject::call_named(...)
```

Do not expose raw `DISPPARAMS` as the normal API.

## Validation and acceptance

Run full workspace validation and focused Windows tests. Use an MSRV-compatible implementation.

Acceptance requires proven argument order, correct property puts, exactly-once cleanup, correct returned-object ownership, preserved COM errors, Excel independence, and no Excel requirement for normal tests.

Commit, push, and open a draft PR.

Do not merge. Stop after reporting the draft PR.
