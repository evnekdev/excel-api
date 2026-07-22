# Excel COM project layout

`excel-com/` is the experimental reusable library. It owns apartment-bound
COM wrappers, generic Automation values, invocation policy, and the bounded
`Application -> Workbooks -> Workbook -> Worksheets -> Worksheet -> Range`
path. It depends only on `windows-sys`.

`tools/excel-com-range-probe/` remains runtime research and compatibility
evidence. `tools/excel-com-microsoft-sample/` is the Microsoft-reference
Automation oracle. `tools/excel-object-model-inventory/` reads the registered
Excel type library and generates the maintained inventory; it is not a
production-library dependency.

`metadata/excel-object-model/` is the machine-readable source of truth for
the generated pages in `docs/excel-object-model/`. Manual prose outside
generated-region markers is preserved by the generator.

Historical evidence remains authoritative for its scoped questions:

- Prompt 05H: `docs/research/excel-com/05h-value-safearray-runtime-matrix.md`
- Prompt 05I: `docs/research/excel-com/05i-python-client-differential.md`
- Prompt 05J: `docs/research/excel-com/05j-error-scode-runtime.md`
- Prompt 06: `docs/research/excel-com/06-internal-automation-value-layer.md`
- Prompt 06A: `docs/research/excel-com/06a-microsoft-cpp-port.md`

The inventory is a structural implementation guide, not a replacement for
those experimental records and not a claim that every Excel object is wrapped.
