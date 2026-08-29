"""Replace customUI14.xml inside an Office package, preserving a backup.

This is a recovery helper for cases where Excel cannot open the add-in far
enough for xVBA to synchronize source files.
"""

from __future__ import annotations

import shutil
import sys
import tempfile
import zipfile
from pathlib import Path


def main() -> None:
    package = Path(sys.argv[1]).resolve()
    ribbon_xml = Path(sys.argv[2]).resolve()
    backup = Path(sys.argv[3]).resolve()
    embedded_name = "customUI/customUI14.xml"

    if not package.is_file() or not ribbon_xml.is_file():
        raise FileNotFoundError("The add-in package or Ribbon XML is missing.")

    backup.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(package, backup)

    with tempfile.NamedTemporaryFile(
        dir=package.parent, suffix=".xlam.tmp", delete=False
    ) as handle:
        temporary_package = Path(handle.name)

    try:
        with zipfile.ZipFile(package, "r") as source:
            if embedded_name not in source.namelist():
                raise KeyError(f"{embedded_name} is missing from {package.name}")

            with zipfile.ZipFile(temporary_package, "w") as target:
                for item in source.infolist():
                    data = (
                        ribbon_xml.read_bytes()
                        if item.filename == embedded_name
                        else source.read(item.filename)
                    )
                    target.writestr(item, data)

        with zipfile.ZipFile(temporary_package, "r") as check:
            if check.testzip() is not None:
                raise zipfile.BadZipFile("The rebuilt Office package failed validation.")

        temporary_package.replace(package)
    finally:
        temporary_package.unlink(missing_ok=True)


if __name__ == "__main__":
    main()
