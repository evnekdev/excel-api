# Signing and trust

Excel and organizational policy can block an otherwise valid XLL. Treat code
signing, trusted publishers/locations, endpoint protection, and Office policy
as deployment concerns outside this crate. Do not weaken global macro, XLM, or
Office security merely to run a test.

Use an isolated test machine/account for policy experiments, record the exact
Excel build and architecture, and restore any test-only setting. The core
library makes no claim that unsigned XLLs will load in every enterprise policy.
