# SAFEARRAY layout

Rectangular reads: 13. Rectangular writes: 9. The 3×4 marker case records this runtime layout:

```json
{"access_balanced":true,"dimensions":[{"element_count":3,"lower_bound":1,"physical_dimension":1,"upper_bound":3},{"element_count":4,"lower_bound":1,"physical_dimension":2,"upper_bound":4}],"element_vartype":12,"metadata_dimension_order":"SafeArrayGetLBound/GetUBound dimensions 1..rank","rank":2,"storage_traversal":"SDK SafeArrayGetElement/PutElement index order [physical_dimension_1, ...]"}
```

The marker traversal is preserved in `rectangular-read-observations.jsonl`; it establishes physical dimension 1 as Excel rows and dimension 2 as Excel columns for this environment.
