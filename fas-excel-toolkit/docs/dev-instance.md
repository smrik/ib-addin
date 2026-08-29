# Dev Excel Instance

This repo supports a dedicated development Excel instance so agents and humans can edit VBA as text files, save, and test immediately without colliding with the user's normal Excel session.

## Source of truth

- `src/vba` is the source of truth.
- `dist/*.xlam` files are build artifacts or development targets.
- Do not edit the VBA project directly in the VBE unless the user explicitly asks for that workflow.

## Preferred dev loop

Use Rubberduck to reload source into the open development add-in:

1. Start a dedicated Excel test instance and open `dist\FAS_Toolkit.xlam`.
2. Edit files under `src\vba` in VS Code.
3. Save the file.
4. In VBE, run `Rubberduck Code Explorer -> Project -> Sync Project -> Update Components from Files...`.
5. Select the changed source files and test immediately.

This keeps Excel open and avoids rebuilding the add-in for normal VBA edits. The existing watcher scripts are legacy recovery code, not the default workflow.

## Why the separate instance matters

Do not rely on one Excel instance loading both the normal installed add-in and the dev add-in copy.

Reasons:

- public macro names can collide
- `Application.OnKey` bindings are global per Excel instance
- Excel can appear to run one add-in while code changes are landing in another
- the Add-ins dialog can become misleading when multiple copies share the same functionality

The separate test instance exists specifically to avoid that ambiguity.

## Tasks to use

- `Start test Excel instance`
  Starts the isolated Excel test session and loads the requested dev add-in.
- `Run macro in Excel`
  Runs a `Public Sub` with no required arguments in the targeted workbook/add-in.

## Targeting rules

- The watcher and macro runner should target the workbook/add-in by file path first.
- If there are multiple Excel instances open, the dev tools must still prefer the specific add-in path the task was launched with.
- When debugging a "change did not go through" report, first verify the watcher printed the sync line for the expected file.

## Shortcut behavior

- Shortcut registration is handled through `FAS_RegisterShortcuts`.
- Shortcut bindings should use workbook-qualified macro names like `'FAS_Toolkit.xlam'!FAS_JumpToSheet`.
- If shortcut wiring changes during live development, `Workbook_Open` will not rerun automatically. Run `FAS_RegisterShortcuts` again or restart the test instance.

## Macro testing rules

- Excel's macro dialog only runs `Public Sub` procedures with no required arguments.
- Helper `Function` procedures will not appear there even when sync succeeded.
- For command-palette features, test either through the actual palette flow or through the public macro entry point that reaches that logic.

## Forms and structural changes

- Keep real `.frm` and `.frx` pairs together.
- Live sync is good for editing existing modules, classes, document modules, and forms.
- If files are renamed, deleted, or structurally reorganized, prefer the normal import/build flow instead of assuming the watcher will clean up old VB components.

## Safety notes

- The isolated test-instance script should not uninstall the user's normal add-ins globally.
- It is acceptable to close other add-in workbooks inside the dedicated test session so only the dev add-in is active there.
- Avoid using the installed roaming add-in as the dev target unless the user explicitly wants that.

## Troubleshooting

If a change seems missing:

1. Confirm the edited file under `src/vba` was saved.
2. Confirm Rubberduck updated the matching component in the development project.
3. Confirm the test was performed in the `TEST INSTANCE` window, not the user's normal Excel window.
4. Run a direct public macro through `Run macro in Excel`.
5. If the behavior still looks stale, restart the dedicated test instance.

If a macro seems unavailable:

1. Check whether it is a `Public Sub` with no required arguments.
2. Check for a compile/runtime error in the VBA project.
3. Check whether shortcut registration needs to be rerun.
