# Excel C API inventory validator

Run from the repository root with Python 3 and Git, without network access:

```text
python tools/excel-c-api-inventory/validate.py check
python tools/excel-c-api-inventory/validate.py reconcile
python tools/excel-c-api-inventory/validate.py freshness
python -m unittest discover -s tools/excel-c-api-inventory -p test_validate.py -v
```

`check` reads the documented implicit schema v1 (an optional explicit integer `schema_version: 1` is also accepted), validates fields, categories, IDs, references and evidence presence, and checks identity against the first catalogue commit `172a003172e5d90a26ea4a8345218988d4b189cc`. Use `--baseline REV` to compare another historical revision. Missing baseline history is a warning and stability is **not assessed**. The comparison detects reassignment of IDs present in that baseline; it does not prove stability for IDs absent from it or forbid removal. Keep old IDs reserved.

Anchors are loaded with Git from the exact commit and checked against the path and Git blob SHA, never the working copy. The checker recognizes balanced `typedef struct` names, a narrow prototype form, and object-like macros after stripping comments. It validates macro expression syntax only, **not effective numeric values, C type signatures, ABI, guards, aliases, or conditional availability**. Unknown schema fields, including proposed alias or guard representations, are rejected pending a reviewed schema extension. It cannot authenticate evidence links or establish that a host run really happened. Evidence arrays must be nonempty for assessed statuses and empty for unassessed/not-run statuses. A `pass` host claim is rejected even with a nonempty string: add a reviewed structured, independently checked real-Excel evidence format before accepting one.

`reconcile` without a census reports `NOT ASSESSED — complete source census pending #117`. A separately reviewed, independently sourced JSON file may be supplied with `--census FILE`. It has `complete` (boolean) and `entries` (array), each entry with `path`, `kind`, `symbol`, and `disposition` (`eligible`, `excluded`, or `ambiguous`); exclusions and ambiguities require a `reason`. Optional `alias_of` and `guard` are reported for review; optional `source_blob` detects a changed blob; optional `expression` compares the source macro expression lexically, ignoring whitespace, without evaluating it. Eligible missing entries, duplicate identities, ambiguities, excluded-but-catalogued entries, mismatched expressions/blobs, and (only for complete censuses) extra catalogue entries are reported as failures. The tool does not extract declarations or verify that the supplied census is independent, complete or reviewed. Even a matching census requires that external review; no complete census is currently supplied. Source changes inside a struct or prototype are not compared semantically. An incomplete census cannot establish extras or coverage.

`freshness` separates pinned-source integrity, source completeness, Rust assessments, native Excel evidence and EDNA functional coverage. Seed assessments are not assessed. Existing schema v1 evidence strings do not supply machine-readable dependency revisions, so assessed records are reported unknown rather than declared current or stale based on unrelated HEAD changes. This is a local report, not a CI or release gate.
