# Packaging

```powershell
powershell -File scripts/build-minimal-xll.ps1 -Profile release
powershell -File scripts/inspect-minimal-xll-exports.ps1
```

Build the sample release XLL on Windows with:

```powershell
powershell -File scripts/build-minimal-xll.ps1 -Profile release
powershell -File scripts/inspect-minimal-xll-exports.ps1
```

The normal package contains the XLL and its documented exports only. It does
not bundle or register the experimental RTD COM prototype. Keep the native
target, architecture, dependency closure, signing policy, and deployment path
explicit; loading an XLL is not equivalent to installing a managed add-in.
