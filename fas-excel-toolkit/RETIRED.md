# Retired — superseded by `patriks-tab`

This tree is no longer the Excel add-in source. **`patriks-tab` is canonical.**

Both trees were the same codebase, forked in two directions. As of 2026-08-29 the
merge went into `patriks-tab`, which already held the newer feature work
(configurable shortcuts, settings dialog, undo manager, IFS/capex/group-by tools,
the fuller Ribbon). What this tree still had, and where it went:

| Was here | Now in `patriks-tab` |
|---|---|
| `round` / `round <n>` / `iferror` palette commands | `vba-files\Module\modFormulaTools.bas`, merged onto the undo-aware wrapper this tree lacked |
| Palette reachability for `visible`, `linkcopy`, `linkpaste`, `colorcapture`, `colorapply`, `delta`, `autofit` | `vba-files\Module\modCommandRegistry.bas` (the macros already existed here, Ribbon-only) |
| Build / import / export / smoke / publish / watch / sync / run-macro / test-instance scripts | `patriks-tab\tools\` |
| `tests\source_contract.ps1` | `patriks-tab\tests\source_contract.ps1`, extended with palette-reachability assertions |

All 29 shared modules were compared. 10 were byte-identical once VBE case-rewriting
(`cells` → `Cells`) and line endings were normalised; the other 19 were reviewed
individually, as were the help catalog (72 entries here, all present in the 91-entry
`patriks-tab` catalog) and the command dispatch table. One further behaviour was
carried across after that review: `FAS_CopyVisibleOnly` now re-raises its error when
`Application.Visible` is false, so the smoke test can detect a failed copy instead of
seeing only a status-bar message.

Three things were deliberately **not** carried over, because `patriks-tab` does them
better:

- `modAddinLifecycle.bas` hardcoded seven `Application.OnKey` bindings. `patriks-tab`
  routes all shortcuts through `FAS_GetShortcutDefinitions` so users can rebind them.
- `modFormulaTools.bas` wrapped formulas with no undo support. `patriks-tab` wraps
  inside `FAS_UndoBeginRange` / `FAS_UndoCommit` / `FAS_UndoCancel`.
- `modProductivityTools.bas` prompted for confirmation before appending a delta sign
  to a non-empty cell. `patriks-tab` appends silently but undoably, which matches the
  AGENTS.md rule against confirming safe, reversible actions. It also adds guards that
  this tree lacked, refusing to write a delta into a formula or error cell.

## Why this is still here

The ported build could not be host-validated: this machine has
**Trust access to the VBA project object model** disabled, so no VBA import can run.
Until `patriks-tab\tools\smoke_test.ps1` passes on a machine with that enabled, this
tree stays as the recovery copy. The repository has no commits yet, so deleting it
now would not be recoverable through git.

Once the smoke test passes, this directory can be removed.
