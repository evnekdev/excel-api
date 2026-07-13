# excel-api

A Rust-first project for building native Microsoft Excel add-ins through the Excel C API.

The intended long-term outcome is an idiomatic Rust ecosystem for native `.xll` add-ins: raw C API bindings, safe values and calls, generated worksheet-function registration, arrays, commands, asynchronous calculation, and optional Windows integrations.

## Current status

The project is in the architecture and ABI-research stage. The first implementation target is a minimal 64-bit Rust XLL that Excel can load and that registers a scalar worksheet function through the Excel C API.

## Architecture references

- [Overall architecture](ARCHITECTURE.md)
- [Architecture index](ARCHITECTURE_INDEX.md)
- [Excel 12 ABI architecture](ABI_ARCHITECTURE.md)
- [Memory and ownership architecture](MEMORY_OWNERSHIP_ARCHITECTURE.md)
- [String architecture](STRING_ARCHITECTURE.md)
- [Array and reference architecture](ARRAY_REFERENCE_ARCHITECTURE.md)
- [Type conversion architecture](TYPE_CONVERSION_ARCHITECTURE.md)
- [Threading architecture](THREADING_ARCHITECTURE.md)
- [Excel call architecture](EXCEL_CALL_ARCHITECTURE.md)
- [Registration architecture](REGISTRATION_ARCHITECTURE.md)
- [Implementation roadmap](IMPLEMENTATION_ROADMAP.md)
- [Excel-DNA capability map](EXCELDNA_CAPABILITY_MAP.md)
- [Codex development prompts](prompts-dev/README.md)

These documents are living references. They should be updated whenever implementation experience changes the proposed architecture or priority of a capability.

## Workspace

```text
crates/
  excel-api-sys/
  excel-api/
  excel-api-macros/

examples/
  minimal-xll/

docs/
  adr/
  checklists/
  diagrams/
  research/

prompts-dev/

tools/
  Excel2013XLLSDK/
  abi-check/
```

The Microsoft SDK material is used as the authoritative ABI reference. See [`tools/Excel2013XLLSDK/README.md`](tools/Excel2013XLLSDK/README.md) for provenance and licensing notes.
