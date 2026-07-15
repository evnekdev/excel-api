# Windows XLL packaging

Run `pwsh -File scripts/package-minimal-xll.ps1` on a fresh Windows x64 MSVC
environment with Rust, PowerShell, and Visual Studio `dumpbin`. It produces a
versioned unsigned package containing the XLL, required export report, notices,
manifest, and SHA256 sums. The checked-in Excel SDK is build provenance only
and is not redistributed.

Signing is optional and external: pass `-CertificateThumbprint` and, normally,
`-TimestampServer`. No certificate, password, or secret is stored in the repo.
Verify a signed package with `signtool verify /pa /all minimal_xll.xll`.
