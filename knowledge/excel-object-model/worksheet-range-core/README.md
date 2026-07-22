# Worksheet and Range core evidence

This directory records the bounded Prompt 08 wrapper contract. It does not
duplicate the full type-library inventory or expose raw COM addresses.

`runtime-observations.jsonl` is an owned-process, version-specific observation:
the live test begins only with zero existing Excel processes, makes the created
instance visible, calls `Quit`, and verifies natural exit. The structural
member identity and DISPIDs remain in `metadata/excel-object-model/`.
