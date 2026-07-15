# Excel Call Context Review

- [x] Call class is known for the M8 catalogue.
- [x] Allowed callback contexts are documented for the M8 catalogue.
- [x] Thread-safe availability of xlFree is verified.
- [x] Return ownership and release-policy source are documented for M7.
- [x] Abort/uncalculated/not-thread-safe release errors are preserved.
- [x] C API result is released on every owned path.
- [x] Call is not made before runtime linking or after unlinking.
- [x] `xlAbort` is documented as cancellation polling, not calculation state;
  preserve/clear modes, Boolean result, MTR restriction, and no-release policy
  are explicit.
# M17 dispatcher review additions

- [ ] The drain entry point requires a genuine typed callback context; no
      context-free or arbitrary-closure API is exposed.
- [ ] The operation's exact requirement is compatible with that context under
      the explicit dispatcher matrix.
- [ ] Queue/controller locks are released before executing the operation or
      calling Excel.
- [ ] Callback waits and nested drains are rejected/suppressed.
- [ ] Shutdown removes the generation and retires queued/selected work before
      unregister and backend unlink.
- [ ] Documentation states that enqueueing does not wake Excel.
