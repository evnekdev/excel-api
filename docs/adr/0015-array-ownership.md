# ADR-0015: Array ownership

DLL-owned return multis deep-copy every pointer-bearing element and reject
arrays-of-arrays/references initially.

## Status

Implemented through M6 for DLL-owned returns. M3 provides semantic `ExcelArray`: immutable boxed
row-major elements, exact checked shape, deep-copied strings, and no nested
arrays or references. M4 validates logical flat return arrays and accounts for
one future contiguous element block plus independent text buffers. M5
materializes that storage with stable pointers and local RAII cleanup. M6
marks only the top-level multi with DLLFree, preserves base-only element tags,
and reclaims the exact root-first owner through AutoFree without traversing the
raw array.
