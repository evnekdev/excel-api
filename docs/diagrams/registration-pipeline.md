```mermaid
flowchart LR
SIG[Rust signature] --> META[Typed registration metadata]
META --> TYPE[Verified type text]
TYPE --> ARGS[xlfRegister arguments]
ARGS --> REG[Registration ID]
```
