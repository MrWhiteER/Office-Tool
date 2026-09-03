"""
Single source of truth for the app's current version number.

Reads the plain-text VERSION file sitting next to this one — the SAME file
build.bat feeds into installer.iss (via ISCC's /DMyAppVersion=... override)
when it compiles the installer, so the running app and the installer it was
built from always agree on what version they are. To ship an update: bump
VERSION, rerun build.bat, tag+release on GitHub with that same number (see
update_checker.py) — nothing else needs editing.

Uses engine.BASE (the read-only bundled-resource root — see engine.py's own
comment on BASE/DATA_BASE) since VERSION ships as bundled data, not
user data; build.bat adds it to PyInstaller via --add-data "VERSION;.".
"""
import os
import engine


def _read_version():
    try:
        with open(os.path.join(engine.BASE, "VERSION"), "r", encoding="utf-8") as f:
            return f.read().strip()
    except Exception:
        return "0.0.0"


APP_VERSION = _read_version()
