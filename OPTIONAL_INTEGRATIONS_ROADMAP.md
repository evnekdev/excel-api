# Optional Integrations Roadmap

This roadmap is outside the core 1.0 stabilization sequence. Work is optional
and must not change the core callback, ownership, or packaging contract without
a separate reviewed design.

The coarse `E` buckets below are retained for historical planning. The concrete
execution plan for replacing a managed Excel integration layer is now
`docs/architecture/native-xll-hosting-roadmap.md`, whose `H0-H8` milestones do
not renumber the historical core or optional milestones.

| Item | Scope | Entry gate |
|---|---|---|
| E1 | RTD clean-host activation comparison | Supported clean 64-bit Excel host; Rust/control comparison through formulas |
| E2 | RTD production API design | E1 evidence plus deployment, signing, lifecycle, and support decision; for a pure-Rust application this is attempted only if H7 proves native async insufficient |
| E3 | Ribbon metadata and UI | H5 design gate: reviewed Office type-library/interface contract, selected binary topology, callback/lifetime model, packaging and trust design |
| E4 | Custom task panes and general COM | Separate COM apartment/lifetime/deployment design coordinated with `docs/architecture/com-automation-roadmap.md` |
| E5 | Autonomous notification adapters | H3 is the preferred concrete path: current-host bridge plus hidden/message-only window, narrow `Application.Run` wake, and genuine helper-macro callback; issue #30 `xlcOnTime` research remains evidence until H3 live validation succeeds |

The native-host roadmap deliberately allows H1-H4 to proceed without waiting
for every generic COM milestone. Ribbon/server work should preferentially reuse
the generic COM ownership and `IDispatch` foundation rather than create a
second permanent unsafe kernel.

The existing M18 research is parked and M19 is deferred. These optional items
do not block M20 core 1.0 review.