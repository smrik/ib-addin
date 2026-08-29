# FAS Command Palette and Precedent Cycling Design

## Goal

Make FAS commands discoverable without memorizing aliases, and provide reliable keyboard navigation through direct formula precedents when Excel's `Ctrl+[` shortcut is unavailable.

## Command Palette

`Ctrl+Shift+J` opens a dedicated modeless FAS command palette inspired by the VS Code command palette.

- Focus starts in the search box.
- The result list initially shows the complete command catalog.
- Typing filters commands by canonical name, aliases, category, keywords, and description.
- Up and Down move through results.
- Enter executes the selected command and closes the palette.
- Escape closes the palette without running anything.
- Each row shows a readable command name and a short description; aliases remain searchable but are visually secondary.
- `help`, `commands`, and `?` open the palette with all commands.
- Queries such as `help chart` and `? ledger` open the palette pre-filtered.

The existing sheet navigator remains separate. Sheet navigation and command execution are different tasks and should not share one result list.

## Command Catalog

Replace the partial private help list with one authoritative catalog exposed by `modCommandHelp`.

Each catalog entry contains:

- canonical command text
- aliases or accepted command forms
- search keywords
- category
- short description

The palette, help searches, unknown-command suggestions, and shortcut-builder command list consume this catalog. The existing dispatcher remains responsible for execution because several command families use parameter parsers rather than fixed command strings.

The catalog must cover every currently documented command family, including navigation, search, bookmarks, formulas, precedents, paste, charts, slide sizing, mappings, drilldown, shapes, and setup/debug commands.

## Direct Precedent Cycling

`Ctrl+Alt+Up` runs `FAS_CycleDirectPrecedent`.

- The first press reads the direct precedents of the active formula cell, stores that formula cell as the trace origin, and jumps to precedent 1.
- Further presses advance through the stored direct precedents.
- Pressing the shortcut after the final precedent wraps to precedent 1.
- Selecting a different formula cell starts a new cycle.
- If the stored workbook, worksheet, or precedent is no longer available, the cycle resets with a friendly status-bar message.
- Existing conservative parsing remains in place for same-sheet and cross-sheet A1 references. Unsupported formulas such as external links, `INDIRECT`, `OFFSET`, structured references, and spill references are reported without guessing.

`Ctrl+Alt+Down` runs `FAS_BackToTraceOrigin`, returning to the formula cell and resetting the active cycle. Existing `pre`, `pre2`, `pre3`, and `back` commands remain compatible.

All shortcut registrations use workbook-qualified macro names and are removed by `FAS_UnregisterShortcuts`.

## Error Handling

- An empty palette remains useful by showing every command.
- No search matches displays a clear empty state while keeping the search box active.
- A command failure is reported through the existing status-bar mechanism.
- Precedent cycling from a non-formula cell reports that a formula cell is required.
- No direct precedents reports that condition without changing the selection.

## Verification

Automated smoke coverage will verify:

- representative commands from every catalog category are present and searchable
- the catalog is larger than the old 12-entry display cap
- same-sheet precedent cycling advances and wraps
- cross-sheet precedent cycling activates the expected worksheet and cell
- return navigation restores the formula origin
- shortcut registration macros can be called without error

Manual Excel smoke will verify palette focus, live filtering, arrow-key movement, Enter execution, Escape dismissal, and both precedent shortcuts.

## Compatibility

The canonical artifact remains `FAS_Toolkit.xlam`. Existing typed command aliases and the separate `Ctrl+Shift+N` sheet navigator continue to work. No client workbook content or settings are modified by opening the command palette.
