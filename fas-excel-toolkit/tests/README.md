# Tests

The primary automated check for this repo is `tools\smoke_test.ps1`, which exercises a built `.xlam` through Excel COM automation on a local Windows machine with Excel installed.

Suggested manual checks after major UI changes:

- open the navigator with `Ctrl+Shift+N`
- confirm it is modeless
- confirm repeated toggles do not spawn duplicates
- confirm search filters the list
- confirm Enter and double-click activate the selected sheet
