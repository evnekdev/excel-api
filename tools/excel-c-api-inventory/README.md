# Excel C API inventory validator

Run from the repository root with Python 3 and Git, without network access:

```text
python tools/excel-c-api-inventory/validate.py check
python tools/excel-c-api-inventory/validate.py reconcile
python tools/excel-c-api-inventory/validate.py freshness
python tools/excel-c-api-inventory/source_census.py --output census.json
python tools/excel-c-api-inventory/audit_source.py census.json
python tools/excel-c-api-inventory/validate.py reconcile --census census.json
python -m unittest discover -s tools/excel-c-api-inventory -p "test*.py" -v
```

All three commands first validate the catalogue's UTF-8 JSONL, schema v1 or v2 (`schema_version: 2` is explicit), required fields, categories, IDs, references, evidence presence and supported declaration anchors against pinned Git blobs. They also check historical identities before reporting reconciliation or freshness. An invalid catalogue cannot produce a successful census comparison.

Historical ID protection reads catalogue-changing Git snapshots reachable from `HEAD` and `refs/remotes/origin/master`, including both sides of merged history, beginning at the original catalogue commit `172a003172e5d90a26ea4a8345218988d4b189cc`. Every recorded ID remains reserved for its original `(path, kind, symbol)` identity. The checker rejects conflicting historical assignments, current removal of any recorded ID, and historical removal along a parent-to-child transition. Schema v1 has **no retirement/supersession mechanism**: retain the record until a reviewed extension exists. `--baseline REV` includes an additional history tip; it cannot replace the required baseline, current branch or `origin/master`. This protects locally reachable history, not unknown or deleted branches or remote changes not fetched locally. A shallow clone, missing original baseline, missing upstream ref, unavailable object or incompatible history yields **NOT VERIFIED**, never PASS. Fetch the relevant complete history and retry. A detached HEAD can supply the current tip, but the normal project workflow uses an attached branch.

Anchors are loaded with Git from the exact commit and checked against the path and Git blob SHA, never the working copy. Version 1 retains its narrow original declaration checks. Version 2 verifies each record against the source-driven candidate's identity, exact span and applicable detail fields, including macro expression and integer value, scalar/pointer/aggregate typedef names, supported callback/prototype signatures, conditional C linkage and the XLOPER conditional member spellings. Integer evaluation accepts literals, identifiers, parentheses, unary plus/minus and bitwise OR only; it never executes source expressions. Unsupported syntax, names, cyclic definitions or conditional forms are errors, not guessed values. It is not a general C parser and does not establish full aggregate semantics, ABI, signedness, external type widths or actual Excel runs. A `pass` host claim is rejected even with a nonempty string until verifiable real-Excel evidence has a reviewed format.

`source_census.py` independently reads the pinned header rather than the catalogue and emits a candidate census without CAPI IDs. `populate.py --write` preserves existing identities and anchors, appends new source candidates with unused IDs, verifies the generated records, and refuses conflicts. Review its output; this is a proposed catalogue, not independent review. `audit_source.py` uses a distinct lexical scanner and expression evaluator to compare source names, spans and macro values with that census, including unaccounted source lines. Both tools were authored in the same run, so their agreement is a mechanical cross-check, not independent human review or proof of exhaustive source semantics. See the [working report](../../docs/evidence/excel-c-api-source-census.md).

`reconcile` without a census reports `NOT ASSESSED — complete source census pending #117`. An independently prepared JSON file may be supplied with `--census FILE`. It has `complete` (boolean) and `entries` (array), each entry with `path`, `kind`, `symbol`, and `disposition` (`eligible`, `excluded`, or `ambiguous`); exclusions and ambiguities require a `reason`. Optional `alias_of`, `guard`, typedef/signature details and source line coordinates are compared with extracted source details when provided; optional `source_blob` detects a changed blob; optional `expression` compares source macro expression text, ignoring whitespace; an accompanying `value` is compared with the supported source-derived effective integer. Eligible missing entries, duplicate identities, ambiguities, excluded-but-catalogued entries, mismatched expressions/blobs, and (only when `complete` is true) extra catalogue entries are reported as mismatches. `complete` means only that the caller asks for a set comparison against all supplied entries. It is **not proof** that the source was exhaustively enumerated. An incomplete census can compare its listed entries but cannot establish extras or coverage.

A matching supplied census is reported as `MATCHES SUPPLIED CENSUS — SOURCE COMPLETENESS NOT VERIFIED`; this tool cannot authenticate independence, review or exhaustive extraction. For #117, the JSON can carry additional metadata for source repository, pinned revision, header path and blob SHA, included/excluded declaration categories, independently established counts, reviewed exclusions/ambiguities and review provenance. These additional metadata are currently **unvalidated annotations**, not proof; #117 must define and implement a reviewed verification method before the validator can report independently verified source completeness. There is no independently reviewed complete census yet. Struct layout and general C semantics remain unverified; supported v2 macro expressions and values are checked.

Exit codes are stable for scripts:

| Code | Meaning |
| --- | --- |
| `0` | `check` or `freshness` input integrity and available history passed. For `reconcile`, reserved for a future independently verified complete reconciliation; currently never returned. |
| `1` | Invalid catalogue, historic identity conflict/removal, or invalid input/census. |
| `2` | Valid catalogue but supplied census has a reconciliation mismatch or ambiguity. |
| `3` | Unverified/incomplete evidence: no census, supplied matching census without verified completeness, or required Git history unavailable. |

Thus a code of `0` for `check` or `freshness` is not an SDK-coverage or Excel-runtime claim, and a code of `3` is not a successful exhaustive reconciliation.

`freshness` validates catalogue integrity first and separates pinned-source integrity, source completeness, Rust assessments, native Excel evidence and EDNA functional coverage. Seed assessments are not assessed. Existing schema v1 evidence strings do not supply machine-readable dependency revisions, so assessed records are reported unknown rather than declared current or stale based on unrelated HEAD changes. This is a local report, not a CI or release gate.
