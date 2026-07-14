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
  return integration deferred to M8 pending root-lifetime cleanup proof
- M8 manual registration and first XLL (implemented; automated 64-bit Excel
  load/calculation/unload/reload passed; interactive UI checks remain)

## Ergonomics

- M9 procedural macros
- M10 compile-time diagnostics
- M11 execution contexts and call catalog
- M12 commands/lifecycle completeness

## Production

- M13 diagnostics
- M14 packaging
- M15 Excel integration/stress harness

## Advanced

- M16 async UDFs
- M17 main-thread dispatcher
- M18 RTD
- M19 COM/Ribbon
- M20 1.0 review
