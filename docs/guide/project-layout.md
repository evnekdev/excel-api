# Project layout

| Package | Role | Status |
| --- | --- | --- |
| `excel-api-sys` | Raw `XLCALL.H` ABI definitions. | Stable target; unsafe. |
| `excel-api` | Safe values, returns, registration, contexts, runtime, and macros re-export. | Stable target, with preview areas noted below. |
| `excel-api-macros` | Closed signature model and generated thunks. | Stable target. |
| `examples/minimal-xll` | Recommended native XLL integration example. | Example only. |
| `examples/minimal-rtd-server` | Windows COM RTD compatibility prototype. | Experimental, unpublished. |

The core crate never makes a raw pointer safe merely by wrapping its type. Its
separation is deliberate: callback views borrow Excel storage; owned semantic
values contain no callback pointers; return planning and `xlAutoFree12` own the
DLL allocation lifecycle.
