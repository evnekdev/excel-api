# M8 64-bit Excel smoke test

Status: **automated live Excel smoke passed; interactive UI cases pending**.

## Artifact

```powershell
pwsh -File scripts/build-minimal-xll.ps1 -Profile release
dumpbin /exports target/release/minimal_xll.xll
```

The produced file is `target/release/minimal_xll.xll`. It targets Windows x64
MSVC and requires Excel 2007 or later's Excel12 ABI. Visual Studio Build Tools
are needed only for normal MSVC linking and optional `dumpbin`; the SDK C++
bridge is not compiled because its verified host-resolution behavior is
implemented directly in Rust.

## Manual procedure

Record Excel version/build and confirm **64-bit** in File > Account > About.
Copy the XLL to a short local path, unblock it in Windows file properties if
necessary, then load it through File > Options > Add-ins > Manage Excel
Add-ins > Go > Browse.

1. Confirm these functions appear in the Function Wizard: `RUST.ADD`,
   `RUST.ECHO`, `RUST.ARRAY.ECHO`, `RUST.REFERENCE.KIND`, and
   `RUST.OPTION.KIND`.
2. Evaluate `=RUST.ADD(2,3)` and expect `5`.
3. Evaluate `=RUST.ECHO("Aé水😀")` and compare exact text.
4. If a formula/UI path can create embedded U+0000, confirm counted UTF-16 is
   preserved; otherwise record this case as not representable through that UI.
5. Place mixed number/text/Boolean/error/blank values in `A1:B3`, evaluate
   `=RUST.ARRAY.ECHO(A1:B3)`, and confirm the spilled 2x3 values.
6. Evaluate `=RUST.REFERENCE.KIND(A1:B3)` and expect `SRef` or `Ref`; confirm
   `=RUST.ARRAY.ECHO(A1:B3)` receives values (`Q`) rather than a reference.
7. Compare `=RUST.OPTION.KIND()` (`missing`) with an explicit blank cell
   argument (`nil` or `value`, recording Excel's supplied representation).
8. Recalculate repeatedly with F9 and Ctrl+Alt+F9. Enable multi-threaded
   calculation and repeat with many calls to the pure `$` functions.
9. Open and close workbooks containing the formulas, deactivate/remove the
   add-in, close Excel, reload Excel, then reload the XLL.
10. Confirm Excel remains stable and formulas re-register exactly once.

## Result record

- Excel version/build: `16.0`, build `20131` (`EXCEL.EXE` file version
  `16.0.20131.20126`).
- Architecture: `Windows (64-bit) NT 10.00`.
- Function results: add `5`; Unicode echo asserted exact; reference kind
  `SRef`; omitted option `missing`; mixed 2x2 spill preserved number, text,
  Boolean, and `#N/A`.
- Registration/unregistration behavior: `Application.RegisterXLL` succeeded in
  two fresh Excel processes. Each process closed cleanly, exercising
  `xlAutoClose`; the second load registered and calculated again.
- Repeated/MTR recalculation: Excel reported MTR enabled with 28 threads; 500
  pure formulas plus `CalculateFullRebuild` completed, last value `501`.
- Unload/reload stability: passed across two fresh Excel processes.
- Interactive cases still pending: visible Function Wizard inspection,
  Add-in Manager deactivate/remove UI, and an embedded-NUL input (ordinary
  formula text cannot express it directly).

## M8.5 reference-freeze update

The automated record above remains the live-Excel evidence for the frozen M8
oracle. Function Wizard and Add-in Manager are still pending visible UI
inspection. The documented direct probes are `RUST.ADD`, Unicode
`RUST.ECHO`, `RUST.ARRAY.ECHO`, `RUST.REFERENCE.KIND`, and
`RUST.OPTION.KIND`; their automated results are recorded above. An embedded
NUL remains pending because ordinary Excel formula text cannot supply one.
CI deliberately validates the build and export table only; it does not attempt
to automate a live Excel UI session.

## M9B generated-thunk rerun

The same automated smoke script was rerun against the M9B release XLL after
the five handwritten worksheet thunks were replaced by generated exports.

Command:

```powershell
powershell -File scripts/smoke-minimal-xll.ps1 -XllPath target/release/minimal_xll.xll
```

Result: **passed in two fresh Excel processes**.

- Excel version/build: `16.0`, build `20131`.
- Architecture: `Windows (64-bit) NT 10.00`.
- MTR: enabled with 28 calculation threads in both passes.
- Add: `5`.
- Unicode echo: exact in both passes.
- Reference kind: `SRef`.
- Omitted option: `missing`.
- Mixed array spill: number `1`, text `text`, Boolean `TRUE`, error `#N/A`.
- Repeated calculation: 500 formulas, last value `501`.
- Reload: the second fresh process registered, calculated, and exited cleanly.

The visible Function Wizard, Add-in Manager deactivate/remove UI flow, and an
embedded-NUL formula input remain pending interactive checks; ordinary formula
text still cannot express the embedded NUL case directly.
