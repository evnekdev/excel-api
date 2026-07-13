# String Architecture

## ABI forms

The modern C API exposes multiple wide-string forms.

### `xltypeStr`

- `XLOPER12` union member;
- first UTF-16 code unit stores payload length;
- payload is not generally NUL-terminated;
- maximum 32,767 UTF-16 code units.

### Counted direct wide string

Registered direct pointer form with a prepended length.

### Null-terminated direct wide string

Registered direct pointer form scanned to NUL.

### Modify-in-place string

Excel allocates the output buffer and the function modifies it. This is a
separate advanced ABI and is deferred.

## Safe types

```rust
ExcelStr<'call>  // borrowed UTF-16 payload
ExcelString      // owned Box<[u16]> payload
String           // owned UTF-8
```

ABI-specific return storage stays internal.

## Unicode policy

- `ExcelStr` and `ExcelString` preserve arbitrary UTF-16 code units.
- Strict conversion to `String` is fallible.
- Lossy conversion is explicit.
- Counted forms preserve embedded NUL.
- Null-terminated forms stop at the first NUL.

## Return policy

General dynamic strings return as:

```text
XLOPER12 xltypeStr | xlbitDLLFree
```

Direct dynamic simple-string returns are not supported initially because they
lack the general `xlAutoFree12` ownership callback.

## Hybrid strings

The book discusses appending a NUL after counted payloads for C-library
convenience. The safe semantic types do not depend on such a hybrid layout.

An internal buffer may optionally reserve a trailing NUL, but:

- it is not part of the logical payload;
- it is never written beyond the Excel maximum/capacity;
- the prefix remains authoritative.

## Arrays

Borrowed multis yield `ExcelStr` without allocation. Owned arrays deep-copy text
to `ExcelString`. DLL-owned return multis own every string payload.
