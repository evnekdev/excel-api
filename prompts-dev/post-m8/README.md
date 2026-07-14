# Excel API post-M8 Codex prompt queue

These prompts assume M1-M8 are merged on `master` at or after commit `f62002a047101a857c0995aabc2935817fced4bb`.

Execute in filename order. Each prompt creates one branch and one draft PR, then stops. Never begin the next prompt before the previous PR is reviewed and merged.

M16-M19 are gated by earlier milestones and require ADR/research before implementation when official behavior is incomplete.
