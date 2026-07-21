# Minimal versus full call-trace diff

Trace records omit pointer, HWND, PID, and local-path values.

| Sequence | Trace events | Add HRESULT |
| --- | ---: | --- |
| A0 | 8 | `0x00000000` |
| Full | 19 | `0x00000000` |

Additional full-sequence operations:

- `Application:Version` — `0x00000000`
- `Workbooks:GetTypeInfoCount` — `0x00000000`
- `Workbooks:QueryInterface(IUnknown)` — `0x00000000`
- `Workbooks:QueryInterface(IDispatch)` — `0x00000000`
