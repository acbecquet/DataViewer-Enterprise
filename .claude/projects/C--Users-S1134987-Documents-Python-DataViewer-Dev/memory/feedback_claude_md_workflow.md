---
name: CLAUDE.md workflow rules
description: User added CLAUDE.md to DataViewer-Enterprise with strict workflow rules — always refer to it before coding
type: feedback
---

User added `CLAUDE.md` to the project root (`DataViewer-Enterprise/CLAUDE.md`). It defines mandatory workflow rules:

1. **Plan mode default** — Enter plan mode for any non-trivial task (3+ steps). Stop and re-plan if things go sideways.
2. **Subagent strategy** — Use subagents liberally for research/exploration. One task per subagent.
3. **Self-improvement loop** — After ANY user correction, update `tasks/lessons.md` with the pattern.
4. **Verification before done** — Never mark complete without proving it works (build, test, demonstrate).
5. **Demand elegance** — For non-trivial changes, pause and ask "is there a more elegant way?" Skip for simple fixes.
6. **Autonomous bug fixing** — Just fix bugs without asking for hand-holding.
7. **Task management** — Write plans to `tasks/todo.md`, verify root causes, minimal impact changes.

**Why:** User wants disciplined, senior-engineer-level workflow with traceability.
**How to apply:** Always read CLAUDE.md at session start. Follow plan-first approach. Update lessons.md on corrections.
