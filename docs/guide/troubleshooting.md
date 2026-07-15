# Troubleshooting

| Symptom | Check |
| --- | --- |
| XLL will not load | Architecture, exports, dependencies, signing/trust, and Excel policy. |
| `#VALUE!` from a function | Generated registration signature, input conversion, and thunk diagnostics. |
| Text is corrupt | Counted versus NUL-terminated UTF-16 contract; do not retain callback pointers. |
| Async result never arrives | Executor installed for this open generation, capacity, cancellation, and lifecycle events. |
| Dispatch ticket remains pending | Expected unless a legal pump callback drains it; enqueue never wakes Excel. |
| Close/unload fails | Preserve diagnostics; do not reactivate the old generation or unload with live work. |

Include the Excel version/build, bitness, exact command, exported symbols, and
bounded diagnostic events in a bug report. Do not include workbooks or private
paths unless they have been sanitized.
