# Excel Call Architecture

## Raw layer

Internal support for:

```text
Excel12
Excel12v
```

The vector form is preferred for generated/variable argument lists.

## Call result

Every call returns:

```rust
Result<ExcelOwnedValue, ExcelCallError>
```

or a no-value result for calls whose documented result is void.

The wrapper validates both:

- C API return code;
- returned `XLOPER12` type/content.

## Call classification

```rust
enum ExcelCallClass {
    CApiOnly,
    WorksheetFunction,
    MacroSheetFunction,
    Command,
}
```

Every function ID carries metadata:

- allowed contexts;
- thread safety;
- result ownership;
- argument rules.

## Legality

The library must encode the book's central rule: not every C API function is
legal from every callback.

Illegal calls fail before invoking Excel where possible.

## `xlFree`

`xlFree` is modeled as a release operation, not a normal value-producing call.

## `xlCoerce`

Coercion is explicit and may return allocated memory requiring `xlFree`.

## Runtime linking

Linking/resolution happens during idempotent initialization, not static
construction and not before `xlAutoOpen`.

Unlinking occurs only after objects that might call Excel have been destroyed.
