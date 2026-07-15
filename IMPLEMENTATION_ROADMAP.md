# Implementation Roadmap

## Foundation

- M0 workspace and architecture
- M1 verified Excel 12 ABI
- M2 borrowed values
- M3 owned values (implemented)
- M4 safe return planning (implemented)
- M5 stable return allocation (implemented)
- M6 DLLFree handoff and AutoFree (implemented)
- M7 Excel-owned results and exactly-once xlFree (implemented); raw XLFree
  return transfer remains intentionally deferred pending a documented
  per-call root-lifetime cleanup mechanism
- M8 manual registration and first XLL (implemented; automated 64-bit Excel
  load/calculation/unload/reload passed; interactive UI checks remain)

## Ergonomics

- M9A procedural macro metadata (implemented)
- M9B generated ABI thunks (implemented; automated parity/build/export checks
  and two-process live 64-bit Excel rerun passed)
- M10 compile-time diagnostics (implemented)
- M11 execution contexts and call catalog (implemented; calculation-state query deferred pending an authoritative C API contract)
- M12 commands/lifecycle completeness (in progress)

## Production

- M13 diagnostics (implemented; real Excel observability validation pending)
- M14 packaging (implemented)
- M15 Excel integration/stress harness (implementation complete; live validation blocked)

## Advanced

- M16 async UDFs (implementation and lifecycle-race hardening complete with
  deterministic automated race/ABI tests; real Excel
  cancellation/recalculation/unload validation pending; started by
  explicit maintainer direction while the M15 live-smoke gate remains blocked)
- M17 main-thread dispatcher (cooperative callback-drained implementation
  complete with deterministic automated tests; live pump validation blocked;
  enqueue does not wake Excel and no production autonomous wake is approved)
- M18 RTD
- M19 COM/Ribbon
- M20 1.0 review
