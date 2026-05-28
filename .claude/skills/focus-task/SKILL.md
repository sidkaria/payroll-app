---
name: focus-task
description: Use when the user asks to start, build, implement, or work on something — "implement X", "let's build Y", "what's next", "next steps", "continue", "pick up Z". Do NOT trigger on brainstorming/ideas (use capture-idea) or plain questions.
---

# focus-task

Bridge the `.backlog/` task files to actual implementation.

## Pipeline

1. **See the backlog** — `node tools/task list`. `.backlog/*.md` is the source of
   truth; `ROADMAP.md` is generated (never hand-edit).
2. **Pick the task**
   - "what's next" / "continue" → `node tools/task next` (deterministic).
   - Names a task → `node tools/task show <ID>`. If `blocked`, tell the user what
     it needs (`depends_on` / `needs_keys`) and ask before proceeding.
   - Not in backlog → create it (`status: ready`, best-guess priority/tier),
     `render`, then proceed; tell the user it was added.
3. **Claim** — `node tools/task set <ID> status in_progress`.
4. **Plan** — break into 3–7 sub-steps with the task tools. One in-progress at a time.
5. **Build** — follow `CLAUDE.md` (locked decisions) and any `ARCHITECTURE.md`.
   New code dirs ship an `ARCHITECTURE.md`. Parallel tool calls when independent.
6. **Verify** — run `./check.sh`; it must exit 0. Never edit it to pass. For UI
   changes, say explicitly a human still needs to eyeball it.
7. **Finish** — `node tools/task done <ID>` → `node tools/task render` → commit in
   the project's message style → `git push origin main`.
8. **Report** — one paragraph: what changed, what's verified, what's next.

## Rules

- Never invent scope; if acceptance is unclear, ask.
- `./check.sh` green before every push. Main-only, no PRs.
- Stop and surface any conflict with a locked decision in `CLAUDE.md`.
- `.backlog/` + `tools/task` + `check.sh` are the contract any external runner
  consumes — keep them clean and the project stays automatable, no coupling.
