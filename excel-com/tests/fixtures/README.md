# Excel COM live-test fixtures

These small, repository-owned Open XML workbooks let controlled live tests use
`Workbooks.Open` when a fresh `Workbooks.Add` is unavailable. Tests must verify
the expected archive structure, copy a fixture to a unique temporary path, and
never overwrite a tracked file.

`tools/create-excel-com-fixtures.py` creates every file deterministically with
the Python standard library. It does not start Excel, read a user profile, or
copy a user workbook. Regenerate from the repository root with:

```powershell
python tools/create-excel-com-fixtures.py
Get-ChildItem excel-com/tests/fixtures -File | Get-FileHash -Algorithm SHA256
```

| Fixture | Intended controlled use | SHA-256 |
|---|---|---|
| `blank.xlsx` | File-lifecycle and generic blank-workbook fallback | `0E635B45583C00955F95A44AAC95BC4CE12B28DF8DC659B388E32604B4957FB8` |
| `blank.xlsm` | Macro-enabled file lifecycle without VBA content | `0DEF27553F11F38009ED95619F38BBBB3CDFE8C97BF80F32CABFFB5E747B7820` |
| `chart-source.xlsx` | Source range for chart creation and formatting | `E03895C5E903B68F9519782BAD9C1B0288A303DF125F722DD21C8A95329C9C04` |
| `pivot-source.xlsx` | Local tabular source for PivotTable scenarios | `680CE55DB1919C4E8A682F242E176A43AACC03DFAD77B07A8DABF9BB55B60B6F` |
| `local-query-source.xlsx` | Local, provider-free QueryTable input | `37821CB413F9C6A26B48F0923DFDA3B36398E28DCA08B48550D284A14CF6E3E8` |

All files have one visible `Sheet1`, no hidden worksheets, no external links,
no data connections, no printer settings, and no document-personal metadata.
`blank.xlsm` has the macro-enabled workbook content type for lifecycle testing,
but deliberately has no VBA project or macro relationship.
