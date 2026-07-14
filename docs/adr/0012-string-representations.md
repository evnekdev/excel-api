# ADR-0012: String representations

Normalize multiple Excel UTF-16 ABI forms into `ExcelStr`, with `ExcelString`
for owned UTF-16 and `String` for UTF-8.

## Status

Implemented for borrowed and owned semantic values in M2-M3. `ExcelString`
stores payload-only `Box<[u16]>`; strict UTF-8 conversion is fallible and lossy
conversion is explicit. ABI return storage remains pending.
