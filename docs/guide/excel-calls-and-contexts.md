# Excel calls and contexts

The typed call catalogue is intentionally small. A call descriptor records its
function ID, argument roots, ownership, legal contexts, return-code policy, and
threading classification. Call it only through the matching context:

| Context | Intended capability |
| --- | --- |
| `ThreadSafeContext` | Verified thread-safe worksheet operations only. |
| `WorksheetContext` | Worksheet-callback operations. |
| `MacroContext` | Legal macro/command operations and approved weaker operations. |
| `LifecycleContext` | Lifecycle-safe operations only. |

Context is a capability, not a main-thread assertion. `xlAbort` is cancellation
polling; it does not report `Application.CalculationState`. `xlretUncalced` is
an Excel return code, not a calculation-state query.
