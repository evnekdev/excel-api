# Prompt 15 live baseline

`workbooks_add_baseline_live` was launched three times as separate ignored-test
processes. Each began with zero `EXCEL.EXE` processes, constructed the Excel
Application successfully, observed `Visible = false`, `DisplayAlerts = true`,
and A1 reference style, then failed at `Workbooks.Add` with outer
`DISP_E_EXCEPTION` (`0x80020009`) and Excel SCODE `0x800A03EC`. Each transient
Excel process exited during the test's passive wait; no process termination,
Office setting change, or broad host diagnosis was performed.

This is an environmental blocker before workbook creation, not a claim that a
Prompt 15 wrapper call was live-validated. The three Prompt 15 live acceptance
tests remain checked in and ignored until Excel accepts `Workbooks.Add` again.
