# Rectangular arrays

`AutomationArray` is explicit `rows × columns` row-major storage, with logical index `row * columns + column`. Decoding requires rank two `SAFEARRAY(VARIANT)`, normalizes its lower bounds away, and maps physical dimension 1 to rows and dimension 2 to columns. Encoding uses deterministic one-based bounds.
