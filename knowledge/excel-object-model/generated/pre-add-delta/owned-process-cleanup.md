# Owned-process cleanup

| Control | Exit bucket | Owned process exited | Forced termination |
| --- | --- | --- | --- |
| `minimal-high-level` | `1-15-seconds` | `true` | `false` |
| `full-high-local-0400` | `1-15-seconds` | `true` | `false` |
| `lower-level-windows-sys` | `1-15-seconds` | `true` | `false` |
| `rust-to-native-c-abi-shim` | `immediate` | `true` | `false` |
| `native-direct-executable` | `immediate` | `true` | `false` |
| `state-b-manual-workbook-closed-controls` | `1-15-seconds` | `true` | `false` |
| `state-c-automation-warmup-controls` | `1-15-seconds` | `true` | `false` |
| `state-d-retained-reference-controls` | `1-15-seconds` | `true` | `false` |
| `state-e-unique-prefix-matrix` | `1-15-seconds` | `true` | `false` |
| `state-f-current-high-level-controls` | `1-15-seconds` | `true` | `false` |
