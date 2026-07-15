# Excel Call Architecture

## Raw layer

Internal support for:

```text
Excel12
Excel12v
```

The vector form is preferred for generated/variable argument lists.

## Call result

Every call returns:

```rust
Result<ExcelOwnedValue, ExcelCallError>
```

or a no-value result for calls whose documented result is void.

The wrapper validates both:

- C API return code;
- returned `XLOPER12` type/content.

## Call classification

```rust
enum ExcelCallClass {
    CApiOnly,
    WorksheetFunction,
    MacroSheetFunction,
    Command,
}
```

Every function ID carries metadata:

- allowed contexts;
- thread safety;
- result ownership;
- argument rules.

## Legality

The library must encode the book's central rule: not every C API function is
legal from every callback.

Illegal calls fail before invoking Excel where possible.

## `xlFree`

`xlFree` is modeled as a release operation, not a normal value-producing call.

M7 introduces the narrow crate-private `ExcelReleaseBackend`; it accepts the
stable top-level root and returns an owned `ExcelReleaseError`. It is not a
general call catalogue and the owner has no global mutable function pointer.

The production adapter is intentionally left to Prompt 08, where a linked and
callback-scoped `Excel12v` capability can call
`Excel12v(xlFree, null, 1, [root])`. The root is caller-supplied storage.
`xlFree` releases auxiliary Excel storage, nulls its contained pointer, and
leaves the root allocation itself intact.

## M11 typed catalogue

`xlCoerce` is explicit: it accepts a callback-owned source and an explicit
target root type, and returns a separate `ExcelOwnedValue` requiring `xlFree`.
`xlfCaller` is available only from worksheet and macro contexts and returns an
owned result. `xlSheetId` with no argument is named `active_sheet_id` because
the documented result is the active/front sheet; `xlSheetNm` with an external
reference whose sheet ID is zero is separately named `current_sheet_name`.
Sheet-name and caller roots retain their documented `xlFree` obligation.

## M11 research boundary: cancellation is not calculation state

`xlAbort` is the selected cancellation-polling call. It returns an immediate
`xltypeBool`: zero arguments, or an explicit TRUE, preserve a pending break;
an explicit FALSE clears it. The Boolean reports only an Esc/CANCEL break
request. It is not Excel calculation progress or application state, and it
does not create an Excel-owned result or an `xlFree` obligation.

Microsoft documents `Application.CalculationState` as a VBA/COM property. No
authoritative `Excel12`/`Excel12v` callback ID and full contract has been
identified for Done/Calculating/Pending, so the C API catalogue intentionally
does not expose `calculation_state`, `is_calculating`, or a placeholder enum.
`xlretUncalced` remains a return code meaning a requested dependency was not
yet calculated; it is not a query result. Revisit only with an authoritative
Microsoft C API ID, selector/arguments, result ownership, and callback-context
contract.

## Runtime linking

Linking/resolution happens during idempotent initialization, not static
construction and not before `xlAutoOpen`.

Unlinking occurs only after objects that might call Excel have been destroyed.

## M17 xlcOnTime compatibility boundary

The raw SDK layer now mirrors checked-in `xlfNow = 74` and
`xlcOnTime = 148 | xlCommand = 32916`. The safe crate exposes only hidden,
experimental methods for the historical two-argument schedule, schedule with
latest time, and four-argument cancellation forms. They preserve the raw
Excel12v return code and immediate Boolean/error/other tag; counted command
text and every argument root remain live for the call, and no immediate scalar
result receives an `xlFree` owner.

This is not a general XLM command escape hatch. Current Microsoft material does
not document the complete modern `xlcOnTime` contract, so lifecycle legality,
security, cancellation, and unload reliability remain live research questions.
The experimental descriptor must not be treated as production catalogue
approval.
## M8 implementation

The production backend mirrors SDK `XLCALL.CPP`: it resolves `MdCallBack12`
from the host executable and accepts `SetExcel12EntryPt`. An atomic stores the
linked entry; safe public code cannot call arbitrary function integers.

The initial typed catalogue contains `xlGetName`, `xlfRegister`, `xlfSetName`,
`xlfUnregister`, and `xlFree`, including context, result-root, argument-count,
thread-safety, and release metadata. Exact C API return-code bits are retained
in `ExcelReturnCode`. `xlGetName` and lifecycle results are represented by
`ExcelOwnedValue` and receive one top-level `xlFree` attempt.

## M16 narrow background callback

`xlAsyncReturn` is not added to general worksheet/context capabilities. The
async controller alone can call it with exactly two roots: a reconstructed
opaque `xltypeBigData` handle and a locally owned result. The controller
decodes Excel's Boolean acceptance result and preserves the exact C API return
code, including `xlretInvAsynchronousContext`. No `xlFree` or DLLFree transfer
is involved. `xlEventRegister` is lifecycle-only and accepts the documented
procedure-name string plus calculation event integer.
