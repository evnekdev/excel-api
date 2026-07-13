```mermaid
flowchart TD
RAW[Excel call ID] --> META[Call metadata]
META --> CHECK{Context permits?}
CHECK -->|No| ERR[ExcelCallError]
CHECK -->|Yes| CALL[Excel12v]
CALL --> OWN[ExcelOwnedValue]
```
