# Optional Integrations Roadmap

This roadmap is outside the core 1.0 stabilization sequence. Work is optional
and must not change the core callback, ownership, or packaging contract without
a separate reviewed design.

| Item | Scope | Entry gate |
|---|---|---|
| E1 | RTD clean-host activation comparison | Supported clean 64-bit Excel host; Rust/control comparison through formulas |
| E2 | RTD production API design | E1 evidence plus deployment, signing, lifecycle, and support decision |
| E3 | Ribbon metadata and UI | Separate Office/Ribbon callback, packaging, and trust design |
| E4 | Custom task panes and general COM | Separate COM apartment/lifetime/deployment design |
| E5 | Autonomous notification adapters | Authoritative Excel-issued callback capability; issue #30 remains open |

The existing M18 research is parked and M19 is deferred. These items do not
renumber historical milestones and do not block M20 core 1.0 review.
