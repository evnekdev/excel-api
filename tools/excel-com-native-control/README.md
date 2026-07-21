# Native Excel COM control

This is an unpublished, x64-only C++17 control for Prompt 05E. It uses only
Windows SDK COM/Automation APIs; it does not use ATL, MFC, Office interop,
`#import`, or generated Excel wrappers.

Configure and build from an empty local directory:

```powershell
cmake -S tools/excel-com-native-control -B <temporary-build-dir> -G Ninja
cmake --build <temporary-build-dir> --config Release
```

The executable accepts `--mode` with one of
`native-cocreate-local-lcid-0400`, `native-cocreate-server-lcid-0000`, or
`native-cocreateex-server-lcid-0000`, plus an optional temporary `--fixture`.
It emits copied scalar JSON only; paths, HWNDs, and pointers are not emitted.
