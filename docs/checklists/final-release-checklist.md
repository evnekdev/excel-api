# Final release checklist

This is a publication gate for `excel-api-sys`, `excel-api-macros`, and
`excel-api`. It does not authorize publication, tagging, or a GitHub release.

## Current RC4 outcome

- [ ] All RC4 release blockers resolved
- [ ] RC4 audit rerun after documentation correction
- [ ] Explicit human publication approval received

Current state: **NOT READY**. See the
[RC4 publication review](../research/rc4-publication-review.md).

## Repository and API

- [x] Branch began from clean `origin/master`
- [x] Core publication boundary is the three publishable crates
- [x] RTD, COM/Ribbon, task panes, and autonomous notification are excluded
- [x] Preview async and dispatcher status is documented
- [ ] Remove or correct obsolete placeholder release documentation
- [ ] Reconcile historical release-audit claims with accepted RC evidence

## Safety, ABI, documentation, and security

- [x] Unsafe-code lints deny undocumented unsafe behavior
- [x] Ownership boundaries are documented and deterministically tested
- [x] ABI checker passed all 145 checks
- [x] Release XLL build passed
- [x] Export inspection reports exactly 18 required production exports
- [x] Workspace Rustdoc and doctests passed
- [x] Core manifests include README, license, repository, documentation,
  keywords, categories, and MSRV metadata
- [x] No committed credential, private-key, certificate, or absolute-user-path
  evidence was found
- [ ] Render and visually inspect corrected repository and package READMEs

## Packaging and staged publication

- [x] `excel-api-sys` package and dry-run publish passed
- [x] `excel-api-macros` package and dry-run publish passed
- [ ] Publish `excel-api-sys` after explicit approval
- [ ] Publish `excel-api-macros` after explicit approval
- [ ] Wait for both crates to index on crates.io
- [ ] Rerun `excel-api` package and dry-run against indexed dependencies
- [ ] Publish `excel-api` after explicit approval

## Release communications and CI

- [x] Draft changelog, release summary, migration note, and limitations exist
  in the RC4 audit
- [x] Latest observed Windows XLL CI workflow is green
- [ ] Re-run CI after release-blocker corrections
- [ ] Obtain explicit authorization before creating a tag or release
