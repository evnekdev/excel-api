# Final Excel-error write policy

All seven tested worksheet errors are returned as `VT_ERROR` with a physical signed `SCODE`: `0x800A0000 | ExcelErrorNumber` (for example, `#N/A`: `2042` -> `0x800A07FA` -> `-2146826246`). Excel/CVErr error numbers are not directly writable raw SCODEs.

| Representation | Scalar Value | Scalar Value2 | 1×1 array | Mixed array | Homogeneous array |
| --- | ---: | ---: | ---: | ---: | ---: |
| Short Excel number | rejected | rejected | rejected | rejected | not tested |
| Full signed SCODE | complete | complete | complete | complete | complete |
| Formula-returned raw copy | complete | complete | complete | complete | complete |
| VBA CVErr control | not run | not run | not run | not run | not run |

Outcome A applies to the seven tested errors: future internal work must preserve the exact physical signed SCODE, expose symbolic kinds only as a construction aid, encode `VT_ERROR` with that SCODE, and support scalar and rectangular-array writes.
