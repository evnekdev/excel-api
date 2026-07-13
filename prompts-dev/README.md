# Development prompts for Codex

These prompt files define the planned implementation sequence for `excel-api`.

Use them in order unless a later prompt explicitly states that it can be
executed independently.

## Recommended sequence

1. `01-verify-xloper12-abi.md`
2. `02-borrowed-excel-values.md`
3. `03-owned-excel-values.md`
4. `04-return-planning.md`
5. `05-stable-return-allocation.md`
6. `06-xlautofree12-handoff.md`
7. `07-excel-owned-api-results.md`
8. `08-manual-function-registration.md`

## General usage

Before loading a prompt into Codex:

1. check out a fresh feature branch from the latest `master`;
2. ensure the working tree is clean;
3. copy the complete prompt into the Codex instruction field;
4. allow Codex to inspect the repository before editing;
5. review all changes before committing.

The prompts deliberately require Codex to:

- avoid unrelated edits;
- preserve the existing architecture and ADRs;
- document every unsafe assumption;
- keep each stage independently reviewable;
- run formatting, linting, and tests;
- report unresolved ABI or SDK uncertainties instead of guessing.

## Branch suggestions

| Prompt | Suggested branch |
|---|---|
| 01 | `verify/xloper12-abi` |
| 02 | `feature/borrowed-excel-values` |
| 03 | `feature/owned-excel-values` |
| 04 | `feature/return-planning` |
| 05 | `feature/stable-return-allocation` |
| 06 | `feature/xlautofree12` |
| 07 | `feature/excel-owned-api-results` |
| 08 | `feature/manual-function-registration` |
