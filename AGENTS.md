# AGENTS.md

Shared guidance for the Excel and PowerPoint productivity add-ins in this workspace.

## Communication

- When reporting information to me, be extremely concise and sacrifice grammar for the sake of concision.
- Lead with the result, then include only the context needed to act or review.
- Keep progress updates short and concrete.

## Product design principles

- Treat these add-ins as productivity tools: reduce clicks, repetition, and mental overhead.
- Make the common action the default path.
- Keep advanced options available through progressive disclosure; do not force them into the normal workflow.
- Remember the user’s last deliberate choice when repeating an action, so non-default choices do not need to be reselected.
- Use the current selection and Office context wherever that makes the action more predictable.
- Preserve existing content, formatting, and native Office behavior unless a change is intentional.
- Avoid confirmation prompts for safe, reversible, expected actions. Confirm risky, destructive, or surprising actions.
- Do not hijack native Office shortcuts or behavior. New shortcuts must remain configurable and avoid collisions.
- Make features discoverable through existing UI patterns, search, and command interfaces.
- Keep similar features consistent across Excel and PowerPoint.
- Do not add functionality beyond the approved scope just because it is technically possible.

## Implementation principles

- Prefer simple, readable functions over clever abstractions.
- Add short comments where they help a later reviewer follow the intent.
- Combine related helpers when doing so reduces complexity and interaction cost.
- Keep UI callbacks thin and place behavior in testable modules.
- Reuse existing project patterns before introducing new ones.
- Consult official Office documentation when relying on existing Office APIs or behavior.
- Keep plans and source comments focused on design intent; keep exact function/API details in source and tests.

## Validation and release principles

- Treat source files as the source of truth; do not edit Office binaries directly.
- Validate source contracts before running host tests.
- Treat Excel/PowerPoint smoke tests as host validation, not as a substitute for source review.
- Do not claim a build or publication is complete when a required host validation fails.
- Preserve recoverable backups when an Office build or synchronization step damages an artifact.
- When Office host errors repeat, report the blocker clearly instead of hiding it with repeated retries.

## Working style

- If asked for a plan or review, stay read-only until implementation is approved.
- Once implementation is approved, finish the safe in-scope work and validate it proportionally.
