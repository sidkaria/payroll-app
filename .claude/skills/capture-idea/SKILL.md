---
name: capture-idea
description: Use whenever the user shares a new idea, feature, design/UX direction, engineering pattern, or any "what if we did X" / "we should also" / "another thought". Do NOT trigger on questions, pure discussion, or a request to implement something already decided (use focus-task for that).
---

# capture-idea

File a new idea into the backlog as `status: idea`. Append-only, minimum edits,
dedup first.

## Pipeline

1. **Dedup check** — read `IDEAS.md` (if present) and `node tools/task list`.
   If the idea is already captured, reply
   `Already covered — <id/section>. Not adding.` and stop.
2. **Create the task** — find the highest `T###` in `.backlog/`, write
   `.backlog/T<next>.md` with frontmatter: `status: idea`, a best-guess
   `priority` (P0/P1/P2) and `tier`, `depends_on: []`, `needs_keys: []`, `tags`,
   `created: <today>`. Body: 1–3 sentences capturing the idea.
3. **Index it** — append a one-liner to `IDEAS.md` (create it if missing):
   `- T### — <short title> (<YYYY-MM-DD>)`.
4. **Render + confirm** — `node tools/task render`; reply terse:
   `Added T### (idea): "<short title>"`.

## Rules

- Never set a new idea to `ready` — grooming does that. Leave it `idea`.
- Multiple ideas in one message → capture each separately, dedup each.
- "discuss only, don't add" → respect it, capture nothing.
- A pure preference (a colour, a wording) that isn't real work → note in `IDEAS.md`
  only, no backlog file.
- If an idea conflicts with a locked decision in `CLAUDE.md`, surface the conflict
  and ask before filing.
