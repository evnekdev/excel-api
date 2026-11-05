# Excel C API inventory validator

Run from the repository root with Python 3 and Git, without network access:

```text
python tools/excel-c-api-inventory/validate.py check
python tools/excel-c-api-inventory/validate.py reconcile
python tools/excel-c-api-inventory/validate.py freshness
python -m unittest discover -s tools/excel-c-api-inventory -p test_validate.py -v
```

All three commands first validate the catalogue's UTF-8 JSONL, schema v1 (an optional explicit integer `schema_version: 1` is accepted), required fields, categories, IDs, references, evidence presence and supported declaration anchors against pinned Git blobs. They also check historical identities before reporting reconciliation or freshness. An invalid catalogue cannot produce a successful census comparison.

Historical ID protection reads catalogue-changing Git snapshots reachable from `HEAD` and `refs/remotes/origin/master`, including both sides of merged history, beginning at the original catalogue commit `172a003172e5d90a26ea4a8345218988d4b189cc`. Every recorded ID remains reserved for its original `(path, kind, symbol)` identity. The checker rejects conflicting historical assignments, current removal of any recorded ID, and historical removal along a parent-to-child transition. Schema v1 has **no retirement/supersession mechanism**: retain the record until a reviewed extension exists. `--baseline REV` includes an additional history tip; it cannot replace the required baseline, current branch or `origin/master`. This protects locally reachable history, not unknown or deleted branches or remote changes not fetched locally. A shallow clone, missing original baseline, missing upstream ref, unavailable object or incompatible history yields **NOT VERIFIED**, never PASS. Fetch the relevant complete history and retry. A detached HEAD can supply the current tip, but the normal project workflow uses an attached branch.

Anchors are loaded with Git from the exact commit and checked against the path and Git blob SHA, never the working copy. The checker recognizes balanced `typedef struct` names, a narrow prototype form, and object-like macros after stripping comments. It validates macro expression syntax only, **not effective numeric values, C type signatures, ABI, guards, aliases, or conditional availability**. Unknown schema fields, including proposed alias or guard representations, are rejected pending a reviewed schema extension. It cannot authenticate evidence links or establish that a host run really happened. Evidence arrays must be nonempty for assessed statuses and empty for unassessed/not-run statuses. A `pass` host claim is rejected even with a nonempty string: add a reviewed structured, independently checked real-Excel evidence format before accepting one.

`reconcile` without a census reports `NOT ASSESSED — complete source census pending #117`. An independently prepared JSON file may be supplied with `--census FILE`. It has `complete` (boolean) and `entries` (array), each entry with `path`, `kind`, `symbol`, and `disposition` (`eligible`, `excluded`, or `ambiguous`); exclusions and ambiguities require a `reason`. Optional `alias_of` and `guard` are reported for review; optional `source_blob` detects a changed blob; optional `expression` compares source macro expression text, ignoring whitespace, without evaluating it. Eligible missing entries, duplicate identities, ambiguities, excluded-but-catalogued entries, mismatched expressions/blobs, and (only when `complete` is true) extra catalogue entries are reported as mismatches. `complete` means only that the caller asks for a set comparison against all supplied entries. It is **not proof** that the source was exhaustively enumerated. An incomplete census can compare its listed entries but cannot establish extras or coverage.

A matching supplied census is reported as `MATCHES SUPPLIED CENSUS — SOURCE COMPLETENESS NOT VERIFIED`; this tool cannot authenticate independence, review or exhaustive extraction. For #117, the JSON can carry additional metadata for source repository, pinned revision, header path and blob SHA, included/excluded declaration categories, independently established counts, reviewed exclusions/ambiguities and review provenance. These additional metadata are currently **unvalidated annotations**, not proof; #117 must define and implement a reviewed verification method before the validator can report independently verified source completeness. There is no exhaustive extractor or verified census yet. Struct/prototype semantic changes and effective macro values remain unverified.

Exit codes are stable for scripts:

| Code | Meaning |
| --- | --- |
| `0` | `check` or `freshness` input integrity and available history passed. For `reconcile`, reserved for a future independently verified complete reconciliation; currently never returned. |
| `1` | Invalid catalogue, historic identity conflict/removal, or invalid input/census. |
| `2` | Valid catalogue but supplied census has a reconciliation mismatch or ambiguity. |
| `3` | Unverified/incomplete evidence: no census, supplied matching census without verified completeness, or required Git history unavailable. |

Thus a code of `0` for `check` or `freshness` is not an SDK-coverage or Excel-runtime claim, and a code of `3` is not a successful exhaustive reconciliation.

`freshness` validates catalogue integrity first and separates pinned-source integrity, source completeness, Rust assessments, native Excel evidence and EDNA functional coverage. Seed assessments are not assessed. Existing schema v1 evidence strings do not supply machine-readable dependency revisions, so assessed records are reported unknown rather than declared current or stale based on unrelated HEAD changes. This is a local report, not a CI or release gate.
