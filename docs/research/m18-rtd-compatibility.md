# M18 RTD compatibility findings

Date: 2026-07-15

## Conclusion

The supported Excel RTD mechanism is a COM `IRtdServer`. A native XLL cannot
implement RTD merely through Excel12/Excel12v; it would have to become a COM
class server as well. No authoritative source located for M18 permits Excel C
API calls from an RTD method or turns `RefreshData` into an XLL callback
capability.

The selected future prototype shape is a separate 64-bit in-process COM DLL.
Implementation is deferred because the current host cannot create a workbook
and the remaining callback-thread, Automation ABI, activation-policy, and
termination behavior must be observed before writing unsafe Rust bindings.

The 2026-07-15 repository preflight reproduced classification
`plain-com-failed`: both pre- and post-XLL-registration `Workbooks.Add` calls
returned `0x800A03EC`. This is an environment gate, not RTD evidence.

## Question resolution

1. **Fundamental integration:** yes, the documented integration is COM
   `IRtdServer` plus the `IRTDUpdateEvent` callback.
2. **Native XLL alone:** no supported native-XLL-only contract was found. A DLL
   shared with an XLL would still need a registered COM class factory and all
   COM obligations.
3. **Callers:** Excel calls `ServerStart`, `ConnectData`, `RefreshData`,
   `DisconnectData`, `Heartbeat`, and `ServerTerminate`; the server calls
   `IRTDUpdateEvent::UpdateNotify`.
4. **Threads/apartments:** Microsoft does not specify physical threads for the
   RTD methods. COM apartment registration and marshaling govern delivery.
   Current Excel behavior remains a live-test item.
5. **Excel12/Excel12v:** no authorization was found. Calls are prohibited in
   the proposed server.
6. **RefreshData capability:** it is an RTD data-delivery COM method, not a
   verified Excel C API callback.
7. **UpdateNotify marshaling:** the callback is a COM interface. Cross-apartment
   use requires a proxy obtained through COM marshaling/GIT; copying the raw
   pointer is invalid.
8. **Refresh value types:** Microsoft specifies a Variant containing a
   two-dimensional array. The complete accepted current `VARTYPE` subset is
   not documented by the located RTD pages and remains a compatibility test.
9. **Topics and ownership:** Excel assigns Long topic IDs; `ConnectData`
   receives a Variant array of strings; `RefreshData` pairs IDs with values.
   Retained input is deep-copied and returned Automation storage is newly
   materialized.
10. **Close/termination:** `DisconnectData` retires unused topics and
    `ServerTerminate` ends the connection. Exact ordering for workbook close,
    Excel termination, calculation, and server unload remains a required live
    matrix.
11. **Registration:** normal activation requires ProgID-to-CLSID and server
    registration. Per-user registration is the non-elevated prototype plan;
    per-machine registration belongs to an installer. Registration-free COM
    cannot be assumed for an Excel-hosted component. An in-process DLL must
    match Excel bitness.
12. **Same DLL:** technically a DLL can export both surfaces, but no safe shared
    unload contract was established. It is rejected as the default.
13. **Signing/trust:** XLL trust and COM activation/policy are separate test
    dimensions. No source establishes that signing or trusting one surface
    authorizes the other.
14. **Optional dependency:** yes. A separate Windows-only package/artifact
    preserves the core workspace's platform and dependency boundary.
15. **M17 wake:** not approved. RTD solves topic notification/data refresh;
    no evidence permits arbitrary dispatcher operations from its callbacks.

## Evidence limits

The Microsoft pages describe the supported COM shape and Automation data, but
do not document every observed thread, accepted Variant subtype, deployment
policy, or close ordering for current Microsoft 365. Those questions are not
silently converted into API guarantees. No prototype or live success is
claimed in this research branch.
