# Security policy

`excel-com` is experimental software. Report suspected vulnerabilities
privately to the repository maintainers through GitHub's private vulnerability
reporting facility; do not include credentials, connection strings, workbook
passwords, or sensitive workbooks in an issue.

The crate does not sandbox Excel. Opening a workbook, attaching to an existing
Excel session, refreshing external data, following Excel-side links, exporting
files, or enabling the optional `macro-runtime` feature can have security or
privacy consequences. Review workbook provenance and external-data settings
before running these operations. Supported release lines and disclosure timing
will be recorded in release notes once published.
