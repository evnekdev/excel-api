# excel-com

`excel-com` is an experimental, unpublished foundation for safe Excel COM
Automation. Its semantic and wrapper APIs may change before a first release.

The initial implemented path is `Application -> Workbooks -> Workbook`. It
supports creating a local Excel instance, inspecting and setting visibility,
getting the version, creating a workbook, and closing or quitting explicitly.
It does not claim complete Excel object-model support.

The crate is layered as Excel wrappers, object-model member descriptors,
Automation values and dispatch invocation, then private `windows-sys` COM
ownership. Research tools exercise this crate but it does not depend on those
tools or their evidence formats.

Excel wrappers are apartment-bound and are neither `Send` nor `Sync`. Callers
create an explicit `ComApartment::sta()` and pass it to `Application::new`.
`Drop` releases COM references but never calls `Quit`; application shutdown is
an explicit operation. Raw COM pointers, `VARIANT`, and `SAFEARRAY` values are
not exposed by the ordinary API.

`AutomationValue` preserves Automation scalar distinctions and rectangular
arrays. `ExcelComError` preserves HRESULT and invocation context without
recording pointer addresses.

See `../docs/excel-object-model/README.md` for the generated inventory and
`../docs/architecture/excel-com-project-layout.md` for repository boundaries.

Live tests are opt-in because they launch a new Excel process:

```powershell
cargo test -p excel-com --test live -- --ignored --test-threads=1
```

Worksheets, ranges, events, charts, macros, existing-session attachment,
marshaling, generic collections, and a stable public API are intentionally out
of scope for this first crate slice.
