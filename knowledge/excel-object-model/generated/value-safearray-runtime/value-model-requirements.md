# Value-model requirements

This is a derived research requirements report, not a public API design. It is based on 13 case definitions, 28 `Value2` scalar rows, 28 `Value` scalar rows, 13 rectangular reads, 1 mixed-array rows, 10 precision rows, and 14 string rows.

1. Scalar physical VARTYPEs observed: 0, 5, 8, 10, 11.
2. Rectangular physical owner VARTYPEs observed: 5, 8204.
3. Integer inputs must preserve their input tag in evidence; Excel normalizes accepted integer cell values to floating-point reads in the scalar data.
4. Future policy must distinguish never-written/cleared cells, `VT_EMPTY`, `VT_NULL`, and empty BSTR input.
5. Excel cell errors are physical `VT_ERROR` values with an exact signed `scode`; they must not be stringified in a raw value layer.
6. Date and currency require member- and NumberFormat-aware interpretation; the evidence retains `Value` and `Value2` reads separately.
7. Rectangular reads are observed as rank-2 `SAFEARRAY(VARIANT)` with lower bounds, upper bounds, and per-element VARTYPEs retained.
8. Rows map to physical SAFEARRAY dimension 1 and columns to dimension 2 in the proven marker case; storage policy must remain explicit.
9. Values at and above IEEE-754 exact-integer boundaries, non-finite values, and negative zero require explicit loss/normalization policy.
10. BSTR UTF-16 length, embedded NUL transformation, Unicode, formula-like strings, and Excel cell-length limits require explicit policy.
11. Unsupported or rejected writes must retain HRESULT, EXCEPINFO, `puArgErr`, and post-failure read-back.
12. Prompt 05I confirms that client-visible Python values are not physical return-VARTYPE evidence; retain the raw kernel result alongside any pywin32 or comtypes differential.
13. The fixed-position 05I mixed-array control accepts Empty, Null, I4, Date, and Currency replacements in this environment, while a VT_ERROR replacement is rejected with preserved failure diagnostics.
14. Negative OA doubles through `Value2` and negative VT_DATE through `Value` are distinct operations; preserve member, source VARTYPE, NumberFormat, and both read-backs.
15. No finalized Rust enum or public API is defined here.
