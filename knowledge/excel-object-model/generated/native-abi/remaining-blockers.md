# Remaining blockers

1. Isolate the full high-level local/0x0400 pre-Add sequence against the passing minimal sequence before changing production behavior.
2. Re-run the production range probe only after that bounded repair.
3. The current windows-rs source checkout could not compile this released-API reproduction because its feature model no longer exposes the released Win32 feature names.

Version matrix:

{"schema_version":1,"id":"released-current","windows":"0.62.2","windows-core":"0.62.2","windows-result":"0.4.1","windows-strings":"0.5.1","windows-sys":"0.61.2","minimal_high_level_add_hresult":0,"classification":"Runtime-observed"}
{"schema_version":1,"id":"released-preceding","windows":"0.62.1","windows-core":"0.62.1","windows-result":"0.4.0","windows-strings":"0.5.0","windows-sys":"0.61.1","minimal_high_level_add_hresult":0,"classification":"Runtime-observed"}
{"schema_version":1,"id":"source-head-447078ea771a97277b710de1e3149c5146af1dc8","windows":"0.62.2 source checkout","minimal_high_level_add_hresult":null,"classification":"Inconclusive","blocker":"current source no longer exposes the 0.62 Win32 feature names required by this released-API reproduction"}

Additional unresolved rows:

{"schema_version":1,"id":"native-cpp-runner-local-lcid-0400","classification":"Inconclusive","detail":"The final standalone C++ runner returned DISP_E_EXCEPTION/0x800A03EC although the C ABI DLL built from the same source succeeded through Rust in every mode.","effect":"Prevents a clean Case D conclusion."}
{"schema_version":1,"id":"windows-rs-source-head","classification":"Inconclusive","detail":"The isolated source-head checkout could not compile the released Win32-feature reproduction because its feature model no longer exposes those feature names.","effect":"No source-head runtime claim."}
