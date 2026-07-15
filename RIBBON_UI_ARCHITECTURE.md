# Ribbon UI Architecture

Status: **deferred optional post-1.0 integration**.

Ribbon customization is not required for ordinary native Excel worksheet
functions, commands, lifecycle behavior, asynchronous UDFs, or cooperative
dispatch. It has a distinct Office UI callback, XML/metadata, deployment,
trust, and compatibility contract from the Excel 12 C API core.

No Ribbon code, Office object-model dependency, COM registration, or packaging
behavior is present in the core crates or minimal XLL. A future E3 design must
define callback legality, ownership, threading, host-version support, signing,
and installation before implementation begins.
