"""Run the source-only checks that are useful before XVBA sync.

This intentionally does not open Excel. It checks the contracts most likely to
break when a module, form, or Ribbon callback is edited.
"""

from pathlib import Path
import re
import xml.etree.ElementTree as ET


ROOT = Path(__file__).resolve().parents[1]
SOURCE_SUFFIXES = {".bas", ".cls", ".frm"}
source_paths = [
    path for path in (ROOT / "vba-files").rglob("*")
    if path.suffix.lower() in SOURCE_SUFFIXES
]
source_text = "\n".join(
    path.read_text(encoding="utf-8-sig", errors="replace")
    for path in source_paths
)
xml_path = ROOT / "ribbons" / "customUI14.xml"
xml_text = xml_path.read_text(encoding="utf-8-sig")
ET.parse(xml_path)

callbacks = set(re.findall(
    r'(?:onAction|getPressed|getEnabled|getVisible|getLabel)="([^"]+)"',
    xml_text,
))
public_names = re.findall(
    r"(?im)^\s*Public\s+(?:Sub|Function)\s+([A-Za-z_]\w*)",
    source_text,
)

missing_callbacks = sorted(callbacks - set(public_names))
duplicate_names = sorted({
    name for name in public_names if public_names.count(name) > 1
})
assert not missing_callbacks, f"Missing Ribbon callbacks: {missing_callbacks}"
assert not duplicate_names, f"Duplicate public procedures: {duplicate_names}"

required_symbols = [
    "FAS_TryRunFormatCycleCommand",
    "FAS_IsFormatCycleCommand",
    "FAS_CycleNumberFormat",
    "FAS_CyclePercentFormat",
    "FAS_CycleCurrencyFormat",
    "FAS_ShowSettings",
    "FAS_TryRunShortcutCommand",
    "FAS_IsKnownShortcutCommand",
]
missing_symbols = [name for name in required_symbols if name not in source_text]
assert not missing_symbols, f"Missing format/settings symbols: {missing_symbols}"

shortcut_text = (ROOT / "vba-files" / "Module" / "modShortcutManager.bas").read_text(
    encoding="utf-8-sig"
)
assert shortcut_text.count("definitions.Add Array(") == 10
assert "Ctrl+Shift+1" in shortcut_text
assert "Ctrl+Shift+5" in shortcut_text
assert "Ctrl+Shift+4" in shortcut_text

print(
    f"PASS XmlParsed=True SourceFiles={len(source_paths)} "
    f"RibbonCallbacks={len(callbacks)} MissingCallbacks=0 "
    f"DuplicatePublicNames=0 ShortcutDefinitions=10"
)
