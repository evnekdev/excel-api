```mermaid
flowchart LR
REF[xltypeRef/xltypeSRef] --> BORROW[ExcelReference]
BORROW --> COERCE[xlCoerce]
COERCE --> OWN[ExcelOwnedValue]
OWN --> VALUE[ExcelValue/ExcelArray]
```
