# Excel COM Codex Prompt Series

This package contains research and implementation prompts for investigating, designing, implementing, and reviewing an `excel-com` Rust crate in `evnekdev/excel-api`.

## Recommended order

1. `01-com-automation-foundations.md`
2. `02-object-model-knowledge-base.md`
3. `03-object-model-analysis.md` - analyze the object model from the knowledge base
4. `04-range-values-and-safearray.md` - `VARIANT`, `SAFEARRAY`, and Range runtime research
5. `05-threading-activation-lifecycle.md` - apartments, activation, and lifecycle
6. `06-typelib-and-events.md` - type-library and events audit
7. `07-architecture-synthesis.md` - architecture synthesis

The existing implementation prompts retain their established identifiers until a
separate implementation-plan change is approved:

- `07-crate-skeleton.md`
- `08-dispatch-kernel.md`
- `09-excel-application.md`
- `10-core-object-model.md`
- `11-range-array-transport.md`
- `12-reliability-and-release-review.md`

Each prompt is intended to produce one focused feature branch and one draft pull request. Review and merge each prompt's output before starting the next prompt unless you deliberately choose to parallelize independent research work.

Prompt 02 builds the pinned, attributed, machine-readable evidence base. Prompt
03 analyzes that evidence before the Range runtime work in Prompt 04; neither
prompt chooses a public wrapper API.
