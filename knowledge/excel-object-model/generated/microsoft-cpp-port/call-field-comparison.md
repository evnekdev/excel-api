# Call-field comparison

The C++ and Rust controls agree on initialization, activation, IID, lookup and invocation LCIDs, `IDispatch` flags, zero-argument and property-put `DISPPARAMS`, reverse COM argument order, result initialization, null output pointers, interface-release order, `SAFEARRAY` cleanup, and `CoUninitialize` timing.

The traces record DISPIDs, VARTYPEs, HRESULTs, counts, and ownership roles only. They do not record raw pointer values or PIDs.
