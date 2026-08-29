"""Patch Excel Live Server's generated VBScript for reliable diagnostics.

The extension embeds the VBScript in dist/extension.js and rewrites the
temporary runnable file when VS Code activates. Both copies need the patch.
A sibling .fas-backup is created before the first edit.
"""

from __future__ import annotations

import argparse
import os
from pathlib import Path
import shutil


TEMP_FOLDER = "EXP-140880SDCEFEF-YURI-SDD-DALVA-EXFOLDERSAS-EDSDS"
SERVER_FILE = f"{TEMP_FOLDER}.vbs"


OLD_HANDLER = """Public Function ErrorHandler(message)
\tIf Err.Number <> 0 Then
\t\tWScript.Echo "Error in " & message & " : " & Err.Description
\t\tCall LogMessages(0, LOG_FILE_PATH, "Error in" & message)
\t\tErr.Clear
\tEnd If
\t
End Function"""

NEW_HANDLER = """Public Function ErrorHandler(message)
\t' Capture Err immediately: later COM calls can overwrite its useful details.
\tDim errNumber, errDescription, errSource, detail
\terrNumber = Err.Number
\terrDescription = Err.Description
\terrSource = Err.Source

\tIf errNumber <> 0 Then
\t\tdetail = "Error in " & message & " [" & CStr(errNumber) & " / &H" & Hex(errNumber) & "]"
\t\tIf Len(errSource) > 0 Then detail = detail & " Source=" & errSource
\t\tdetail = detail & ": " & errDescription
\t\tWScript.Echo detail
\t\tIf Len(LOG_FILE_PATH) > 0 Then Call LogMessages(0, LOG_FILE_PATH, detail)
\t\tErr.Clear
\tEnd If
End Function"""


PLAIN_REPLACEMENTS = {
    OLD_HANDLER: NEW_HANDLER,
    'Set book =  GetObject(pathWorkbook).Application.Workbooks(VSCODE_ENV("WorkbookName"))':
        "Set book = GetObject(pathWorkbook)",
    'Set book =  GetObject(VSCODE_ENV("WorkspacePath") & "\\" & VSCODE_ENV("WorkbookName")).Application.Workbooks(VSCODE_ENV("WorkbookName"))':
        'Set book = GetObject(VSCODE_ENV("WorkspacePath") & "\\" & VSCODE_ENV("WorkbookName"))',
    'Set book = GetObject(VSCODE_ENV("WorkspacePath") & "\\" & VSCODE_ENV("WorkbookName")).Application.Workbooks(VSCODE_ENV("WorkbookName"))':
        'Set book = GetObject(VSCODE_ENV("WorkspacePath") & "\\" & VSCODE_ENV("WorkbookName"))',
    "\txlApp.Calculation    = -4135  ' xlCalculationManual\n": "",
    "\txlApp.Calculation    = -4105  ' xlCalculationAutomatic\n": "",
    "\txlAppCh.Calculation    = -4135  ' xlCalculationManual\n": "",
    "\txlAppCh.Calculation    = -4105  ' xlCalculationAutomatic\n": "",
}


def js_string_fragment(text: str) -> str:
    """Encode VBScript as it appears inside the bundled JavaScript string."""
    return (
        text.replace("\\", "\\\\")
        .replace("'", "\\'")
        .replace("\r\n", "\\n")
        .replace("\n", "\\n")
        .replace("\t", "\\t")
    )


def patch_text(text: str, escaped_js: bool) -> tuple[str, list[str]]:
    changed: list[str] = []
    for index, (old, new) in enumerate(PLAIN_REPLACEMENTS.items(), start=1):
        if escaped_js:
            old, new = js_string_fragment(old), js_string_fragment(new)
        if old in text:
            text = text.replace(old, new)
            changed.append(f"replacement-{index}")
    return text, changed


def patch_file(path: Path, escaped_js: bool) -> list[str]:
    original = path.read_text(encoding="utf-8", errors="strict")
    patched, changes = patch_text(original, escaped_js)
    if not changes:
        if "Capture Err immediately" in original:
            return ["already-patched"]
        raise RuntimeError(f"No expected XVBA patterns found in {path}")

    backup = path.with_name(path.name + ".fas-backup")
    if not backup.exists():
        shutil.copy2(path, backup)
    path.write_text(patched, encoding="utf-8", newline="")
    return changes


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--extension-dir", type=Path, required=True)
    args = parser.parse_args()

    extension_js = args.extension_dir / "dist" / "extension.js"
    temp_script = Path(os.environ["TEMP"]) / TEMP_FOLDER / SERVER_FILE

    for target, escaped_js in ((extension_js, True), (temp_script, False)):
        changes = patch_file(target, escaped_js)
        print(f"PATCHED {target}: {', '.join(changes)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
