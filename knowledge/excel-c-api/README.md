# Microsoft Excel C API — canonical declaration catalogue

This is the **raw Excel C API evidence/inventory layer** for `excel-api`, distinct from the higher-level [Excel-DNA subsystem map](../../docs/architecture/excel-dna-subsystems-and-crates.md). Its existence does **not** mean the Rust project or Excel-DNA implements every C API entry.

## Files

| File | Responsibility | Current completeness |
| --- | --- | --- |
| [`symbols.jsonl`](symbols.jsonl) | Persistent `CAPI-xxxxxx` identifier per inspected C API declaration, source anchors and related EDNA IDs | **1,132 provisional source-derived records, including the 15 original IDs. Independent completeness review pending.** |
| [`schema.md`](schema.md) | Required fields, identity, evidence and verification rules | Versions 1 and 2 |
| [Architecture and coverage model](../../docs/architecture/c-api-inventory-layer.md) | Canonical layering, behavior evidence, mapping, closure and implementation boundary | Defined |
| [Microsoft SDK locally](../../tools/Excel2013XLLSDK/README.md) | Checked-in primary header, SDK samples, EULA, libraries and SDK documentation | Present; redistribution rights still require review |

## Evidence baseline

- **Primary declaration header:** [`tools/Excel2013XLLSDK/INCLUDE/XLCALL.H`](../../tools/Excel2013XLLSDK/INCLUDE/XLCALL.H); Excel 2013 XLL SDK, header version **15.0**. Header Git blob: `4eb50e8fd3aea40d283181b7dcc74871ba88d854`. Initial immutable source commit: `6c8354d6536e18b58559f9cc3134f25a9e58d478`.
- **Header provenance and license:** [#114](https://github.com/evnekdev/excel-api/issues/114) and unresolved rights review [#120](https://github.com/evnekdev/excel-api/issues/120); sample-only headers, Windows SDK headers and Excel-DNA native host headers are **separate evidence classes**, not Excel C API symbol declarations.
- **Excel-DNA reference implementation:** [Excel-DNA/ExcelDna pinned at `f484749`](https://github.com/Excel-DNA/ExcelDna/tree/f484749260d5cc43c35f7aea5f7a0b0f8f74a78f/Source). Its C#/native bridges and higher-level features are *supplemental evidence*.
- **Practical behavior and difficult cases:** Steve Dalton, *Financial Applications Using Excel Add-in Development in C/C++*, 2nd edition, Wiley, 2007. Book evidence task [#115](https://github.com/evnekdev/excel-api/issues/115). The full copyrighted PDF is privately stored outside public GitHub; **not** a present-day Excel spec.
- **Rust baseline:** exact current commit, crate paths and validation evidence must be recorded **per assessed record**, not inferred from this source-only starter inventory.

## Use

Search the JSONL catalogue by exact C API symbol or persistent `CAPI-xxxxxx` ID; link its ID from [EDNA parent issues](https://github.com/evnekdev/excel-api/issues/78) and future implementation/test issues. The same raw C API identifier can underlie several EDNA subsystem behaviors. A higher-level Excel-DNA or Office COM feature can have no matching C API entry.

Do not create one issue for each JSONL row automatically. Open a focused issue only for a real inventory uncertainty, source incompatibility, missing Rust functionality or independent validation task. The raw completeness audit is pending under [EDNA-02 #90](https://github.com/evnekdev/excel-api/issues/90). Until it is reconciled with the complete SDK source, **no percentage of C API coverage is meaningful**.

All current `rust_status` fields are `unassessed`; all current `excel_host_status` fields are `not-run`. These are conservative placeholders, **not findings that source is missing or host behavior fails**.

The [header provenance inventory](../../docs/evidence/native-xll-header-provenance.md) records the pinned Microsoft and Excel-DNA headers and licensing uncertainties. Run the local [catalogue validator](../../tools/excel-c-api-inventory/README.md) for integrity, explicitly bounded reconciliation and freshness reports; the exhaustive source census remains [#117](https://github.com/evnekdev/excel-api/issues/117).

The [source-census working report](../../docs/evidence/excel-c-api-source-census.md) states the candidate counts, accounting rules and outstanding independent review.
