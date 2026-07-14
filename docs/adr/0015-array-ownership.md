# ADR-0015: Array ownership

DLL-owned return multis deep-copy every pointer-bearing element and reject
arrays-of-arrays/references initially.

## Status

Partially implemented in M3 for semantic `ExcelArray`: immutable boxed
row-major elements, exact checked shape, deep-copied strings, and no nested
arrays or references. DLL-owned return multis remain pending.
