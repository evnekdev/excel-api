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

M7 introduces the narrow crate-private `ExcelReleaseBackend`; it accepts the
stable top-level root and returns an owned `ExcelReleaseError`. It is not a
general call catalogue and the owner has no global mutable function pointer.

The production adapter is intentionally left to Prompt 08, where a linked and
callback-scoped `Excel12v` capability can call
`Excel12v(xlFree, null, 1, [root])`. The root is caller-supplied storage.
`xlFree` releases auxiliary Excel storage, nulls its contained pointer, and
leaves the root allocation itself intact.

## `xlCoerce`

Coercion is explicit and may return allocated memory requiring `xlFree`.

## Runtime linking

Linking/resolution happens during idempotent initialization, not static
construction and not before `xlAutoOpen`.

Unlinking occurs only after objects that might call Excel have been destroyed.
## M8 implementation

The production backend mirrors SDK `XLCALL.CPP`: it resolves `MdCallBack12`
from the host executable and accepts `SetExcel12EntryPt`. An atomic stores the
linked entry; safe public code cannot call arbitrary function integers.

The initial typed catalogue contains `xlGetName`, `xlfRegister`, `xlfSetName`,
`xlfUnregister`, and `xlFree`, including context, result-root, argument-count,
thread-safety, and release metadata. Exact C API return-code bits are retained
in `ExcelReturnCode`. `xlGetName` and lifecycle results are represented by
`ExcelOwnedValue` and receive one top-level `xlFree` attempt.
