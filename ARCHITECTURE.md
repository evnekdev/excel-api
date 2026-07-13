# Project Architecture

## Mission

`excel-api` is a Rust-first framework for native Microsoft Excel XLL add-ins.

The project is not merely a set of raw bindings. Its central purpose is to wrap
the Excel C API in a safe ownership, lifetime, threading, and capability model
that prevents common XLL failures:

- invalid union access;
- memory leaks;
- double frees;
- returning unstable pointers;
- illegal C API calls from the wrong context;
- retaining callback-owned memory;
- thread-unsafe static return storage;
- mismatched registration type text.

## Workspace

```text
crates/
  excel-api-sys
  excel-api
  excel-api-macros

examples/
  minimal-xll

tools/
  abi-check
```

### `excel-api-sys`

Raw ABI only:

- `XLOPER12`;
- `XLREF12`, `XLMREF12`;
- `FP12`;
- constants and function IDs;
- callback signatures;
- `Excel12`/`Excel12v` function pointer types.

No ownership policy or safe constructors.

### `excel-api`

Safe public API:

- borrowed values;
- owned semantic values;
- strings and arrays;
- references;
- conversion traits;
- call contexts;
- Excel-owned result RAII;
- XLL-owned return memory;
- registration descriptors;
- lifecycle runtime;
- diagnostics.

### `excel-api-macros`

Generated glue only:

- exported thunks;
- signature validation;
- registration metadata;
- panic boundary integration.

Runtime logic remains in `excel-api`.

## Dependency direction

```text
excel-api-sys <- excel-api <- user XLL
excel-api-macros -> generated references into excel-api
```

## Core invariants

1. No panic crosses an FFI boundary.
2. Unsafe code is isolated to ABI, raw parsing, calls, and return materialization.
3. Borrowed callback values cannot outlive the callback.
4. Excel-owned and XLL-owned memory use separate RAII types.
5. XLL return pointers are published only after backing storage is stable.
6. Excel-owned results are released with `xlFree` or transferred with
   `xlbitXLFree`, never `xlAutoFree12`.
7. XLL-owned return memory uses `xlbitDLLFree` and `xlAutoFree12`.
8. Registration type text is generated from verified Rust signatures.
9. Context types restrict which C API operations are legal.
10. First production target is Windows x64 MSVC and Excel 12+ ABI.

## Interface/core separation

Exported thunk code is thin:

```text
Excel ABI
  -> validate/borrow input
  -> convert
  -> call ordinary Rust function
  -> logical return
  -> materialize
  -> handoff
```

Business logic must remain independent from Excel.

## Version strategy

The initial implementation targets Excel 12+ only. Legacy Excel4/xloper support
is a possible future compatibility crate or feature, not a constraint on the
first safe API.
