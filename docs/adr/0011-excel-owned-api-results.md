# ADR-0011: Excel-owned API results

- Status: Proposed

Results produced by `Excel12v` use a dedicated `ExcelOwnedValue` wrapper carrying a verified release policy. The wrapper allows borrowing or copying into `ExcelValue` and releases through the correct Excel path. It never uses `xlAutoFree12`, and raw ownership flags are not exposed publicly. Ownership behavior is documented per supported call category.
