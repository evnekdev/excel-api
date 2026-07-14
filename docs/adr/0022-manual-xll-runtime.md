# ADR-0022: Manual Excel12 XLL runtime

## Status

Accepted and implemented. Automated live 64-bit Excel integration passed;
interactive Function Wizard/Add-in Manager UI checks remain.

## Decision

Mirror the checked-in SDK `XLCALL.CPP` bridge in Rust: resolve the host
executable's `MdCallBack12` export and accept `SetExcel12EntryPt` for the SDK's
cluster-host case. The x64 `XLCALL32.LIB` is verified as an x64 import library,
but it exports only the legacy bridge functions and does not provide
`Excel12v`. This avoids a C++ build step while retaining Microsoft's loading
model. The SDK remains governed by `ExcelSDK_eula.rtf`; no SDK source is copied
into a produced XLL.

The runtime uses `Uninitialized -> Initializing -> Initialized -> Closing`, a
mutex for transitions, an atomic callback entry point, no `static mut`, and no
lock held over an Excel call. Registration IDs and copied module identity are
runtime-owned. Failed initialization unregisters successful earlier entries in
reverse order before unlinking.

Typed signatures generate registration text. `Q` is value-only and `U` is
reference-preserving; modifiers follow argument codes in `!`, `#`, `$`, `&`
canonical order, with `#$` rejected. Pure thunks use fresh DLLFree roots.

Raw XLFree transfer remains deferred. Microsoft documents freeing auxiliary
memory but provides no callback to reclaim a unique per-call root carrying
only `xlbitXLFree`, and does not document combining XLFree and DLLFree. Live
returns therefore deep-copy into the proven DLLFree owner.
