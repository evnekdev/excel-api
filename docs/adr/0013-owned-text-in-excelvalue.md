# ADR-0013: Owned text

Use `ExcelValue::Text(ExcelString)` initially.

## Status

Implemented in M3. Callback text is copied directly from UTF-16 to UTF-16;
there is no intermediate UTF-8 conversion.
