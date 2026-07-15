# M18 RTD working-host validation plan

Status: pending. The available host cannot create a plain workbook before any
XLL or RTD server is loaded.

## Current host gate result

On 2026-07-15, after building the release minimal XLL, the command

```powershell
powershell -File scripts/excel-stress-harness.ps1 -Preflight
```

failed with classification `plain-com-failed`. Both the plain-COM and
post-registration `Workbooks.Add` probes returned COM error `0x800A03EC`; XLL
registration itself reported success. Because the failure begins before RTD
activation and no workbook exists, no formula, callback, threading,
disconnect, or termination result can be claimed. Generated artifacts remain
uncommitted because they contain local paths and process details.

Run this matrix only with an isolated, supported 64-bit Excel installation and
a reviewed minimal RTD prototype. Do not weaken organization-wide security,
Trusted Locations, macro, XLM, or COM policy.

## Environment record

Record sanitized values for:

- Windows version and session type;
- Excel architecture, version, build, and channel;
- RTD server architecture, version, hash, signing state, and server model;
- ProgID/CLSID, per-user/per-machine scope, and registration rollback result;
- applicable COM/Office policy and trusted/ordinary location classification;
- command lines and test settings without user names or private paths.

## Compatibility sequence

1. Prove plain COM `Workbooks.Add` succeeds before registering the server.
2. Register the exact prototype and confirm COM activation through the ProgID.
3. Enter one RTD formula and record `ServerStart`, `ConnectData`,
   `UpdateNotify`, `RefreshData`, `DisconnectData`, and `ServerTerminate`
   ordering, HRESULTs, topic IDs, threads, and apartments.
4. Repeat with two topics, duplicate subscribers, and distinct subscribers.
5. Produce updates faster than Excel refreshes and prove bounded coalescing,
   correct latest values, and bounded CPU/memory/handles.
6. Disconnect one topic while another remains active; prove no later value or
   notification originates from the disconnected topic.
7. Exercise workbook close, formula removal, reopen/reconnect, calculation
   modes, recalculation, and Excel close.
8. Stop with notification pending; prove the producer joins, callback access
   is revoked, COM references reach zero, and no call occurs after termination.
9. Repeat load/close cycles and record memory, private bytes, handles, threads,
   crash events, and remaining server processes. Treat these as trend evidence,
   not proof of leak freedom.
10. Test ordinary versus policy-blocked locations and signed/unsigned artifacts
    where permitted. Record a blocked result rather than changing policy.

## Callback and capability evidence

For every COM method and `UpdateNotify`, record process ID, thread ID,
apartment initialization/model where observable, nesting, and timing. These
observations do not authorize Excel12/Excel12v calls. The prototype must make
no Excel C API call and must never construct an `excel-api` context.

## Required pass criteria

- registration and rollback are deterministic;
- RTD formulas connect and repeatedly receive correct owned values;
- multiple topics/subscribers and reconnect work;
- queues remain bounded and repeated updates coalesce;
- disconnect and termination prevent later updates/notifications;
- producers stop and COM references release on every close path;
- no Excel or server crash is reported;
- memory/handle trends are recorded;
- normal supported policy admits the chosen deployment;
- no XLL backend call occurs from any RTD thread.

Only after this matrix passes may an ADR approve a production RTD framework.
Even a pass does not approve RTD as an M17 wake adapter without separate
authoritative Excel C API capability evidence.
