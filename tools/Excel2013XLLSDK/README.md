# Microsoft Excel 2013 XLL SDK

This directory contains the Microsoft Excel 2013 XLL SDK used as the authoritative ABI reference for `excel-api-sys`.

## Contents

The checked-in SDK layout includes:

- `INCLUDE/XLCALL.H` — Excel C API structures, constants, function identifiers, and registration type codes;
- `SRC/XLCALL.CPP` — Microsoft callback bridge implementation;
- `LIB/` — import libraries for supported architectures;
- `SAMPLES/` — Microsoft example and framework projects;
- `DOC/` — SDK documentation;
- `ExcelSDK_eula.rtf` — the accompanying Microsoft licence terms.

## Provenance

- Product: Microsoft Excel XLL SDK
- Toolkit/header version: 15.0, as identified by `INCLUDE/XLCALL.H`
- Intended Excel generation: Excel 2013 / Excel 12+ C API
- Repository import date: 2026-07

The files should remain unmodified so that ABI comparisons can be made against the original SDK material. Project-specific code belongs in sibling directories such as `tools/abi-check`.

## Licensing

The SDK is Microsoft material and is governed by `ExcelSDK_eula.rtf`. Before distributing releases or mirrors of this repository, confirm that the intended redistribution is permitted by those terms.

If redistribution is not permitted, replace this directory with a setup script, checksums, and local installation instructions, and keep the SDK itself outside version control.

## ABI checker

Run the project ABI checker from the repository root:

```powershell
cargo run --manifest-path tools/abi-check/Cargo.toml
```

The checker uses this directory by default. Set `EXCEL_XLL_SDK_DIR` to use a different local SDK installation.
