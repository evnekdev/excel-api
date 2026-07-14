# ABI Architecture

Status: Implemented for the Windows x86_64 MSVC Excel 12+ definitions scoped
by Prompt 01. The standalone checker verifies them against the checked-in SDK.

## Supported target

Initial target:

```text
Windows x86_64
MSVC ABI
Excel 12+ C API
XLOPER12 / Excel12 / Excel12v
```

## Authoritative source

Raw definitions must be verified against official `xlcall.h` headers.

The book confirms important Excel 12 changes:

- `xltype` is a 32-bit field;
- rows and columns are 32-bit;
- `xltypeInt` and error storage are 32-bit;
- strings use `XCHAR*`;
- `xltypeRef` stores an `XLMREF12*` and sheet ID;
- `xltypeBigData` has distinct input/output pointer/handle meaning.

## Binding strategy

- Generate or transcribe once from official headers.
- Commit curated Rust definitions.
- Do not require bindgen or Clang in normal builds.
- Maintain an optional C/Rust ABI checker.

## Required raw types

- `XLOPER12`;
- `XLREF12`;
- variable-length `XLMREF12`;
- `FP12`;
- `XCHAR`, row, column, Boolean aliases;
- `xltypeBigData` union;
- callback/function pointer types.

## ABI checker

The checker compares:

- size;
- alignment;
- offsets;
- constant values;
- function pointer compatibility.

If the SDK is unavailable, normal tests continue and SDK verification is
reported as skipped.

## Legacy boundary

The old `xloper`/Excel4 API differs materially in:

- string width and maximum length;
- row/column widths;
- integer/error widths;
- registration type text.

Legacy support must not contaminate the initial modern safe API.
