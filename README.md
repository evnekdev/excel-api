# excel-api

A Rust-first project for building native Microsoft Excel add-ins through the Excel C API.

The intended long-term outcome is an idiomatic Rust ecosystem for native `.xll` add-ins: raw C API bindings, safe values and calls, generated worksheet-function registration, arrays, commands, asynchronous calculation, and optional Windows integrations.

## Architecture references

- [Excel-DNA capability map](EXCELDNA_CAPABILITY_MAP.md) — bird's-eye inventory of Excel-DNA functionality, Rust feasibility, implementation strategies, scope and priorities.
- [Implementation roadmap](IMPLEMENTATION_ROADMAP.md) — proposed workspace, design principles, staged milestones, risks, testing strategy and initial issues.

These documents are living references. They should be updated whenever implementation experience changes the proposed architecture or priority of a capability.

## Current status

The project is in the architecture and ABI-research stage. The first implementation target is a minimal 64-bit Rust XLL that Excel can load and that registers a scalar worksheet function through the Excel C API.
