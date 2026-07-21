# Excel session-state transition

| State | Classification | Minimal high-level | Full high-level | Cleanup |
| --- | --- | --- | --- | --- |
| `A-cold-boot-excel-never-opened` | cold-session-observation | `0x00000000` | `0x00000000` | `passed` |
| `B-manual-workbook-closed` | clean-warm-session-observation | `0x00000000` | `0x00000000` | `passed` |
| `C-automation-warmup` | clean-warm-session-observation | `0x00000000` | `0x00000000` | `passed` |
| `D-retained-workbooks-reference` | clean-warm-session-observation | `0x00000000` | `0x00000000` | `passed` |
| `E-unique-prefix-matrix` | fresh-process-prefix-matrix-success | `0x00000000` | `0x00000000` | `passed` |
| `F-current-high-level-failure-controls` | high-level-runtime-path-sensitive | `0x80020009` | `0x80020009` | `passed` |
