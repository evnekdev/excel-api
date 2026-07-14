# Codex orchestration: execute prompts one at a time

You are working in `evnekdev/excel-api`. This directory is an ordered queue of milestone prompts.

## Startup
1. Inspect repository, branch, worktree, remotes, recent commits, and open PRs.
2. Fetch `origin`; confirm `master` is clean and current.
3. Read this file and this directory's `README.md`.
4. Identify the first filename-ordered prompt whose milestone is not merged on `origin/master`. Confirm by implementation/commit/PR state, not branch names alone.
5. Read that prompt fully.

## Per-milestone protocol
1. Run baseline validation.
2. Create the exact branch named by the prompt from latest `origin/master`.
3. Implement only that prompt.
4. Commit in logical reviewable commits; separate formatting-only changes where practical.
5. Run all required validation and record exact results.
6. Push the branch.
7. Open a **draft PR** targeting `master`.

The PR body must include: prompt filename; design decisions; changed modules; safety/ownership/context invariants; tests; exact validation results; live Excel status; docs/ADRs; deferred work/risks; and an acceptance checklist.

## Mandatory stop
After opening the draft PR, print branch, commit SHA, and PR URL; summarize human-review points; then STOP. Do not merge and do not begin the next prompt.

## Resume
Only after the user confirms merge: verify the PR is merged, update local `master`, identify the next unmerged prompt, and repeat.

## Failure rules
- Stop on pre-existing baseline failure and report it.
- Stop rather than guess unclear Excel behavior.
- If live Excel is unavailable, mark validation pending and keep the PR draft.
- If authentication prevents push/PR, commit locally and provide exact commands; never pretend a PR exists.
- Never force-push master, rewrite merged history, or merge your own PR.
