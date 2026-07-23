# Contributing

Contributions should preserve the crate's Windows COM safety boundary: public
wrappers must stay apartment-bound, raw COM pointers must remain private, and
new unsafe code requires a local `SAFETY:` explanation and focused tests.

Before proposing a change, run the repository validation commands in
[`AGENTS.md`](AGENTS.md). Do not commit workbook credentials, connection
strings, user paths, process IDs, or machine-specific runtime logs. Changes to
the public `excel-com` API must update the release API snapshot and relevant
documentation.
