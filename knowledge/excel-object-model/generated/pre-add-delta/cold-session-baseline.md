# Cold-session baseline

The records below are separate from warm-session and prefix observations. No raw process identity, local path, account name, HWND, or pointer is persisted.

| Order | Control | Add | Owned process exited | Session clean |
| ---: | --- | --- | --- | --- |
| 1 | `minimal-high-level` | `0x00000000` | `true` | `true` |
| 2 | `full-high-local-0400` | `0x00000000` | `true` | `true` |
| 3 | `lower-level-windows-sys` | `0x00000000` | `true` | `true` |
| 4 | `rust-to-native-c-abi-shim` | `0x00000000` | `true` | `true` |
| 5 | `native-direct-executable` | `0x00000000` | `true` | `true` |

Interpretation: a successful full high-level local/0x0400 cold control is **cold-session success; prior session contamination is a credible hypothesis**, not a code repair.
