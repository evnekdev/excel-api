# Microsoft Excel 2013 XLL SDK

This directory contains the Microsoft Excel 2013 XLL SDK used as the
authoritative ABI reference for `excel-api-sys`.

## Contents

- `INCLUDE/XLCALL.H` — structures, constants, function identifiers, and
  registration type codes;
- `SRC/XLCALL.CPP` — callback bridge implementation;
- `LIB/` — import libraries;
- `SAMPLES/` — example and framework projects;
- `DOC/` — SDK documentation;
- `ExcelSDK_eula.rtf` — Microsoft licence terms.

## Provenance

- Toolkit/header version: 15.0
- Intended API: Excel 2013 / Excel 12+
- Repository import date: 2026-07

Keep SDK files unmodified. Project code belongs in sibling directories.

## Licensing

These files are Microsoft material governed by `ExcelSDK_eula.rtf`. Confirm
redistribution rights before public mirroring or release. If redistribution is
not permitted, retain setup instructions and checksums but remove the payload.

## ABI checker

```powershell
cargo run --manifest-path tools/abi-check/Cargo.toml
```

Override the SDK root with `EXCEL_XLL_SDK_DIR`.
