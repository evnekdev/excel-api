# Internal value-model requirements

This derived report records research requirements only; it defines no public Rust API.

1. Retain raw VARTYPE, HRESULT, EXCEPINFO, `puArgErr`, SAFEARRAY rank, and bounds independently from any Python-client conversion.
2. A client-visible `None`, numeric value, `datetime`, `Decimal`, or tuple is not evidence of a raw returned VARTYPE.
3. Preserve the exact mixed-array position and client-side input metadata for all 41 Python and 50 raw rows.
4. Preserve `Value` and `Value2`, OA serial, and NumberFormat context for all 120 Python and 24 raw date rows.
5. Shape/rank rejection or coercion must retain the target range and source sequence shape (185 rows).
6. Formula2 spill and blocked-spill results remain capability observations, not assumptions (30 rows).
7. Each live run must begin with no Excel process and finish with an owned natural exit; 26 recorded run(s) meet that condition.
8. No public enum, conversion policy implementation, or Excel-specific vtable is introduced.
