# Excel Call Context Review

- [ ] Call class is known.
- [ ] Allowed callback contexts are documented.
- [x] Thread-safe availability of xlFree is verified.
- [x] Return ownership and release-policy source are documented for M7.
- [x] Abort/uncalculated/not-thread-safe release errors are preserved.
- [x] C API result is released on every owned path.
- [ ] Call is not made before runtime linking or after unlinking.
