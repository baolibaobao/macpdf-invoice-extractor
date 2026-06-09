"""Application entry point for local run and PyInstaller packaging."""

from __future__ import annotations

import os
import sys
from pathlib import Path


def _configure_packaged_tcl_tk() -> None:
    if not getattr(sys, "frozen", False):
        return

    try:
        import _tkinter

        tcl_version = str(_tkinter.TCL_VERSION)
        tk_version = str(_tkinter.TK_VERSION)
    except Exception:
        tcl_version = "8.6"
        tk_version = "8.6"

    bases = []
    if hasattr(sys, "_MEIPASS"):
        bases.append(Path(sys._MEIPASS))

    executable_path = Path(sys.executable).resolve()
    contents_dir = executable_path.parent.parent
    bases.extend(
        [
            contents_dir / "Frameworks",
            contents_dir / "Resources",
            contents_dir,
        ]
    )

    search_roots: list[Path] = []
    for base in bases:
        search_roots.extend([base / "_tcl_data", base])

    def find_library(script_name: str, preferred_dir: str) -> Path | None:
        for root in search_roots:
            candidates = [
                root / preferred_dir,
                root,
            ]
            for candidate in candidates:
                if (candidate / script_name).exists():
                    return candidate

        for root in search_roots:
            if not root.exists():
                continue
            for script_path in root.rglob(script_name):
                return script_path.parent
        return None

    tcl_dir = find_library("init.tcl", f"tcl{tcl_version}")
    tk_dir = find_library("tk.tcl", f"tk{tk_version}")

    if tcl_dir:
        os.environ["TCL_LIBRARY"] = str(tcl_dir)
    if tk_dir:
        os.environ["TK_LIBRARY"] = str(tk_dir)


_configure_packaged_tcl_tk()

import tkinter  # noqa: E402,F401 - ensure PyInstaller includes Tcl/Tk runtime hooks.

from src.main_ui import main


if __name__ == "__main__":
    raise SystemExit(main())
