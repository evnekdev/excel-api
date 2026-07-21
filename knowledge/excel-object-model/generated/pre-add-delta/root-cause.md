# Root cause

**Classification: high-level runtime/path sensitivity, inconclusive.** The isolated prefix matrices contain no failing prefix, while 6 recovery-path rows fail before `Open` with the Excel exception. The current independent controls keep the lower-level generic `IDispatch` and native paths successful. No individual pre-`Add` operation, ownership transition, or storage behavior is established as causal; no repair is applied.
