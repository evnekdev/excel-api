//! Windows COM implementation for the isolated compatibility DLL.

mod abi;
#[expect(
    clippy::undocumented_unsafe_blocks,
    reason = "centralized OLE calls are covered by automation's module safety contract"
)]
mod automation;
#[expect(
    clippy::undocumented_unsafe_blocks,
    reason = "query-only Win32 calls are covered by diagnostics' module safety contract"
)]
mod diagnostics;
#[expect(
    clippy::undocumented_unsafe_blocks,
    reason = "raw COM calls are covered by server's module safety contract"
)]
mod server;
