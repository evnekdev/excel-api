# Prompt 05H transport gate

The initial non-cold session observation is retained as historical evidence only. It is not attributed to an ABI or product defect.

After reboot, before any manual Excel use, the zero-process controls-first baseline passed: minimal high-level A0, full high-level A8, lower-level `windows-sys`, native C++ direct, and Rust-to-native C-ABI shim all created and closed a workbook. The raw L/S/X baseline then passed `Workbooks.Add`, controlled `Workbooks.Open`, and the `A1.Value2 = 42` write/read/clear smoke in every mode.

Following the raw-kernel modularization, ten fresh processes also passed with zero-process pre/post gates: L×4, S×3, and X×3. Every run verified `Add`, `Open`, `Value2`, and owned-process exit. The Prompt 05H transport gate is therefore satisfied and matrix work may proceed.
