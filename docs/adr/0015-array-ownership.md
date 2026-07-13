# ADR-0015: Array ownership

DLL-owned return multis deep-copy every pointer-bearing element and reject
arrays-of-arrays/references initially.
