# Excel-DNA Capability Map for Rust

This is a high-level planning map.

| Capability | Rust feasibility | Initial strategy |
|---|---|---|
| XLL lifecycle | High | Native exports and runtime state machine |
| Worksheet UDF registration | High | Manual descriptors, then macros |
| Type conversion | High | Borrowed/owned/return value layers |
| Strings | High | UTF-16-first borrowed and owned types |
| Mixed arrays | High | Flat row-major multis |
| Floating arrays | High | Dedicated `FP12` path |
| References | High | Separate reference types and coercion |
| Thread-safe UDFs | High | Context whitelist and per-call returns |
| Async UDFs | High but later | Native async architecture |
| RTD | Feasible, COM-heavy | Separate crate |
| Ribbon/COM | Feasible, Windows-specific | Separate crates |
| IntelliSense | Feasible later | Metadata/provider integration |
| Packing/signing | Feasible | Build/package tool |
| .NET reflection features | Not directly applicable | Rust metadata/macros instead |

The project should not port Excel-DNA wholesale. It should reproduce user-facing
capabilities through Rust-native ownership and type-system guarantees.
