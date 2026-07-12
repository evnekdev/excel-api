# Initial implementation notes

This package establishes the first compile-ready workspace outline.

## Included

- `excel-api-sys`: raw ABI constants, error discriminants, lifecycle signatures,
  and initial `XLOPER12`-related structures.
- `excel-api`: safe value types, conversion traits, typed execution-context
  tokens, and registration descriptors.
- `excel-api-macros`: transparent placeholder attributes that reserve the public
  macro names without pretending ABI thunk generation is finished.
- `examples/minimal-xll`: a `cdylib` example exporting lifecycle placeholders.

## Deliberately incomplete

The following require authoritative SDK comparison and dedicated implementation
work before they should be exposed as working functionality:

1. exact ABI verification for every `XLOPER12` union member;
2. loading or linking `Excel12v`;
3. actual `xlfRegister` calls;
4. UTF-16 length-prefixed strings;
5. arrays and references crossing the ABI;
6. self-contained `ExcelReturn` allocation and `xlAutoFree12`;
7. panic containment around every exported callback;
8. generated ABI thunks in `excel-api-macros`.

The placeholder XLL lifecycle exports validate static metadata only. They do not
yet register `RUST.ADD` with Excel.
