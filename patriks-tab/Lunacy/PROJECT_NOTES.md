# Patriks Tab project notes

- Source files under `vba-files` are authoritative; XVBA syncs them into the add-in.
- Preserve approved behavior, Ribbon callback names, form/control names, and keyboard mappings.
- Optimize for a simple, keyboard-friendly productivity tool with clear grouping and low interaction cost.
- Prefer direct, readable VBA over abstraction layers. Consolidate only when it clearly reduces duplication or lifecycle risk.
- Comments should explain intent, lifecycle constraints, Office quirks, or error-handling decisions. Do not narrate obvious statements.
- Do not edit the `.xlam` binary directly.
- Do not run Excel/COM smoke tests. Validate source contracts, Ribbon XML/callbacks, and the `xvba_unit_test` suite instead.
- Existing user changes are authoritative and must be preserved unless they directly conflict with the requested quality pass.

