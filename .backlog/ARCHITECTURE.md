# .backlog

One markdown file per task — the source of truth for open work. `ROADMAP.md` is a
generated view (`node tools/task render`); never hand-edit it.

Frontmatter: `id` (matches filename, e.g. T010.md), `status`
(idea|scoped|ready|in_progress|blocked|done), `priority` (P0|P1|P2), `tier`,
`depends_on: [ids]`, `needs_keys: [ENV names]`, `tags: []`, `created`.

`tools/task`: `next` (deterministic pick) · `list` · `show <id>` · `set <id> <k> <v>`
· `done <id>` · `render` · `validate`. `next` returns the highest-priority `ready`
task whose deps are all `done` and whose `needs_keys` are present (env or
`.backlog/keys.available`).
