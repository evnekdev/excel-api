# Project Architecture

This document is the architectural blueprint for `excel-api`.

## Core principles

- Safe-first wrapper around the Excel C API.
- Exact ABI isolation in `excel-api-sys`.
- Ergonomic public API in `excel-api`.
- Procedural macros generate glue, not runtime logic.
- Panic-safe FFI boundary.
- Self-contained return allocations.
- Explicit registration before distributed registration.

## Proposed workspace

```text
crates/
  excel-api-sys
  excel-api
  excel-api-macros
```

Future crates include async, ribbon, RTD, COM, UI and packaging.

## Layering

`excel-api-sys <- excel-api <- excel-api-macros <- user XLL`

## Ownership model

Use separate types for borrowed inputs, owned values and returned values:
- ExcelValueRef
- ExcelValue
- ExcelReturn

## Registration

Static descriptors -> validation -> xlfRegister -> runtime registry.

## Lifecycle

xlAutoOpen initializes runtime.
xlAutoClose unregisters and releases runtime.
xlAutoFree12 owns all return allocations.

## Threading

Typed execution contexts enforce legal API usage.

