# Excel worksheet-error SCODE runtime (Prompt 05J)

## Scope and baseline

This focused research branch starts from `origin/master` `27291f87c6c28380859c77290f5ab7e8c80f4445`. It adds no public API and leaves the 05H/05I corpora unchanged.

## Documentation and representation

Microsoft documents `xlErr*` values and `CVErr(errornumber)` as worksheet-error identifiers, while the Automation `VARIANT` layout identifies the `VT_ERROR` payload as an `SCODE`. The controlled runtime result distinguishes them: physical SCODE is `0x800A0000 | error_number`, signed as `i32`.

## Formula-returned values

All seven controlled formulas returned `VT_ERROR` through both `Value` and `Value2`: `#NULL! 0x800A07D0`, `#DIV/0! 0x800A07D7`, `#VALUE! 0x800A07DF`, `#REF! 0x800A07E7`, `#NAME? 0x800A07ED`, `#NUM! 0x800A07F4`, and `#N/A 0x800A07FA`.

## Write matrix and policy

Short Excel numbers failed for scalar and array writes with `0x80020009`. Constructed full signed SCODEs and exact formula-returned `VariantCopy` values completed for `Value`, `Value2`, 1×1 arrays, mixed arrays, and homogeneous arrays for every tested error. Outcome A applies: future internal work must preserve exact signed SCODEs and may support scalar and rectangular array writes.

## Controls and audit

pywin32 controls are client-visible only: its explicit error wrapper accepted both supplied numeric forms, so that result does not establish physical VARTYPE/SCODE conversion. comtypes formula-copy calls failed and its installed VARIANT union exposes no SCODE member. VBA automation was deliberately not enabled, so no macro workbook or Office security change was made.

The Rust constructor starts with `VariantInit`, writes `VT_ERROR` and the signed union `scode`, and has raw-bit/`VariantCopy`/`SafeArrayPutElement` unit coverage. No pointers, HWNDs, PIDs, local paths, generated wrappers, or temporary workbooks are persisted.

## Validation and remaining blocker

The opt-in raw and Python controls each required zero pre-existing Excel processes and verified natural exit. The only remaining non-decision is the optional VBA control, which was not run because its security setting was not changed. Prompt 06 must use the signed full-SCODE policy rather than the short Excel number.
