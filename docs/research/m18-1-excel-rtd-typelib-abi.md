# M18.1 Excel RTD type-library ABI audit

Date: 2026-07-15

## Source of truth

The prototype does not copy an internet sample. The audit script loads the
Excel type library registered on the validation host with `LoadTypeLibEx` and
reflects its dual interfaces through `ITypeLib`/`ITypeInfo`.

- Library GUID: `{00020813-0000-0000-C000-000000000046}`
- Version: 1.9
- LCID: 0
- Registered platform: `SYS_WIN32` Automation type library for Win64 Excel
- Observed Excel file version: `16.0.20131.20154`
- `IRtdServer`: `{EC0E6191-DB51-11D3-8F3E-00C04F3651B8}`
- `IRTDUpdateEvent`: `{A43788C1-D91B-11D3-8F39-00C04F3651B8}`

`scripts/inspect-excel-rtd-typelib.ps1` fails unless the installed definitions
match the pinned GUIDs, `stdcall` calling convention, method order, parameter
directions, and Automation shapes below.

## Raw dual-interface vtables

Both interfaces inherit the seven `IUnknown`/`IDispatch` slots. The hidden raw
`IRtdServer` partner then has six slots:

1. `ServerStart(IRTDUpdateEvent*, LONG* retval)`
2. `ConnectData(LONG, SAFEARRAY(VARIANT)**, VARIANT_BOOL*, VARIANT* retval)`
3. `RefreshData(LONG*, SAFEARRAY(VARIANT)** retval)`
4. `DisconnectData(LONG)`
5. `Heartbeat(LONG* retval)`
6. `ServerTerminate()`

The raw `IRTDUpdateEvent` partner has four slots:

1. `UpdateNotify()`
2. `get_HeartbeatInterval(LONG* retval)`
3. `put_HeartbeatInterval(LONG)`
4. `Disconnect()`

All methods return `HRESULT`. The reflected flags distinguish input, in/out,
and retval pointers. Rust compile-time assertions pin `GUID=16`,
`HRESULT/LONG=4`, `VARIANT_BOOL=2`, and the complete 13- and 11-pointer vtable
sizes. The prototype uses `windows`/`windows-core` 0.62.2 for authoritative
Windows COM/OLE base types and functions; only the Office interfaces are
declared locally from the audited type library.

## Ownership implications

Incoming topic SAFEARRAY/BSTR data is borrowed for the call and deep-copied.
Initial values are fresh VARIANTs. `RefreshData` returns a newly allocated
two-dimensional `VT_VARIANT` SAFEARRAY with bounds `[2, topic_count]`, lower
bounds zero, topic IDs at index `[0, n]`, and values at `[1, n]`. Successful
return transfers the SAFEARRAY to Excel; partial failure destroys it exactly
once. Automation memory is never mixed with XLOPER12 or `xlFree` ownership.

This audit establishes ABI shape, not Excel callback thread identity, formula
acceptance, deployment policy, or Excel C API legality.
