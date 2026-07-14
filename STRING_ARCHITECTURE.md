# String Architecture

## Status

- **Status:** Borrowed, owned, planned, and stable local return strings
  implemented through M5.
- **Implemented in:** `borrowed.rs` (`ExcelStr`) and `value.rs`
  (`ExcelString`), bounded callback-copy conversion in `convert.rs`, and
  `return_plan.rs` (`ReturnText` and `PlannedText`), and `return_alloc.rs`
  (stable counted return buffers).
- **Test coverage:** empty, ASCII, BMP, surrogate pairs, unpaired high and low
  surrogates, embedded NUL, strict/lossy decoding, UTF-8 encoding, and source
  independence.
- **Remaining limitations:** DLLFree handoff, Excel-owned API strings, and
  modify-in-place/direct dynamic returns.

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

`ExcelString` stores exactly `Box<[u16]>` payload units. It stores neither an
Excel length prefix nor a trailing terminator and exposes no mutable units.
Direct construction from UTF-16 is infallible because arbitrary code units are
valid semantic data. Callback-copy entry points separately apply configurable
resource limits before allocation.

The callback-borrowing layer implements three distinct audited parser entry
points:

- counted `xltypeStr` payloads reached through the `XLOPER12` decoder;
- counted direct UTF-16 callback arguments;
- null-terminated direct UTF-16 callback arguments.

All return `ExcelStr<'call>` over the original code units. Parsing allocates
nothing, preserves embedded NUL and unpaired surrogates, and performs no UTF-8
or lossy conversion. Counted lengths and null scans are bounded by the Excel 12
32,767-code-unit limit.

## Unicode policy

- `ExcelStr` and `ExcelString` preserve arbitrary UTF-16 code units.
- Strict conversion to `String` is fallible.
- Lossy conversion is explicit.
- Counted forms preserve embedded NUL.
- Null-terminated forms stop at the first NUL.

Strict `ExcelString` to `String` conversion returns
`Utf16ConversionError`; lossy conversion is available only through the
explicit `to_string_lossy` method. UTF-8 input is encoded infallibly to UTF-16.

## Return policy

General dynamic strings return as:

```text
XLOPER12 xltypeStr | xlbitDLLFree
```

Direct dynamic simple-string returns are not supported initially because they
lack the general `xlAutoFree12` ownership callback.

M4 planning retains either the original valid UTF-8 `String` or arbitrary
`ExcelString` UTF-16 payload. UTF-8 planning counts `encode_utf16()` without
allocating a UTF-16 payload. Each planned text records exact payload units and
prefix-plus-payload units, enforces Excel's 32,767-unit hard limit and separate
project limits, and permits embedded NUL. M5 consumes that metadata directly.

M5 materialization creates exactly `[XCHAR length][payload units]` in a final
`Box<[XCHAR]>`. It appends no terminator. UTF-8 sources encode directly into the
payload; UTF-16 sources are copied without transcoding. Planned payload and
storage lengths are rechecked before a pointer is used. The `xltypeStr` pointer
targets the prefix at index zero and remains stable for the `ExcelReturn`
lifetime.

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
