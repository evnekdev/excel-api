# ADR-0015: Array ownership

DLL-owned return multis deep-copy every pointer-bearing element and reject
arrays-of-arrays/references initially.

## Status

Partially implemented through M4. M3 provides semantic `ExcelArray`: immutable boxed
row-major elements, exact checked shape, deep-copied strings, and no nested
arrays or references. M4 validates logical flat return arrays and accounts for
one future contiguous element block plus independent text buffers. M5
materializes that storage with stable pointers and local RAII cleanup. DLLFree
handoff remains pending.
