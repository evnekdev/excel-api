# Memory-safety review checklist

## ABI and parsing

- [ ] Layout and calling convention verified against the supported SDK.
- [ ] Ownership bits are masked before base-tag matching.
- [ ] Union access matches the validated tag.
- [ ] Pointers, alignment, signed dimensions, lengths, and arithmetic are checked.
- [ ] Unsafe functions document pointer, lifetime, ownership, and thread requirements.

## Borrowed inputs

- [ ] Callback inputs are never freed.
- [ ] Borrowed wrappers carry the callback lifetime.
- [ ] Safe code cannot retain or send them beyond the call.
- [ ] Async/worker use deep-copies first.

## Returns

- [ ] Logical validation occurs before allocation/pointer publication.
- [ ] Every published pointer targets stable final storage.
- [ ] No pointer targets stack memory or a growable buffer.
- [ ] Handoff consumes ownership exactly once.
- [ ] DLL-free flags are applied centrally and only at handoff.
- [ ] No fallible work occurs after handoff.

## Cleanup

- [ ] `xlAutoFree12` reconstructs the exact original allocation.
- [ ] Cleanup is panic-free and thread-independent.
- [ ] Excel-owned results never reach `xlAutoFree12`.
- [ ] XLL returns never use an Excel allocator.
- [ ] Debug allocation counts return to zero.

## Failure and tests

- [ ] Exported callbacks catch unwinding.
- [ ] All pre-handoff errors clean up through RAII.
- [ ] Destructors do not panic.
- [ ] Overflow, malformed dimensions, limits, and allocation-stage failures are tested.
- [ ] Recalculation, workbook close, add-in unload, MTR, and panic stress tests exist.
