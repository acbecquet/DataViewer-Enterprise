## Workflow Orchestration

### 1. Plan Mode Default
- Enter plan mode for ANY non-trivial task (3+ steps or architectural decisions)
- if something goes sideways, STOP and re-plan immediately - don't keep pushing
- Use plan mode for verification steps, not just building
- Write detailed specs upfront to reduce ambiguity

### 2. Subagent Strategy
- Use subagents liberally to keep main context window clean
- offload research, exploration, and parallel analysis to subagents
- For complex problems, throw more compute at it via subagents
- One tack per subagent for focused executiion

### 3. Self-Improvement Loop
- After ANY correction from the user: update 'tasks/lessons.md' with the pattern
- Write rules for yourself that prevent the same mistake
- Ruthlessly iterate on these lessons until mistake rate drops
- review lessons at session start for relevant project

### 4. Verification before done
- Never mark a task complete without proving it works
- Diff behaviour between main and your changes when relevant
- ask yourself: "Would a staff engineer approve this?"
- run tests, check logs, demonstrate correctness

## 5. Demand elegance (balanced)
- for non-trivial changes: pause and ask "is there a more elegant way?"
- if a fix feels havcky: "Knowing everything I know now, implement the elegant solution"
- skip this for simple, obvious fixes - don't over-engineer
- challenge your own work before presenting it

### 6. Autonomous bug fixing
- when given a bug report: just fix it. Don't ask for hand-holding.
- Point at logs, errors, failing tests - then resolve them
- zero context switching required from the user
- go fix failing CI tests without being told how

## task management
1. **Plan first**: write plan to 'tasks/todo.md' with checkable items
2. **verify plan**: find root causes. No temporary fixes. Senior developer standards.
3. **minimal impact**: changes should only touch what's necessary. Avoid introducting bugs.